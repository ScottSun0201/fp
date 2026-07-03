# -*- coding: utf-8 -*-
"""
发票-对账单自动匹配引擎 (REQ-025)

四级加权评分系统:
  1. 金额匹配 (权重 40%): 发票明细金额 vs 对账单明细金额
  2. 物料映射匹配 (权重 35%): 通过 sys_material_mapping 表查询映射关系
  3. 数量匹配 (权重 15%): 数量差异 <= 5% 视为完全匹配
  4. 日期匹配 (权重 10%): 发票日期在对账单交货日期 ±15 天内

评分阈值:
  - >= 80 分: 自动匹配 (match_level = 'full')
  - 50 ~ 79 分: 建议匹配 (match_level = 'partial')
  - < 50 分:  不匹配   (match_level = 'unmatched')

结果写入 rcn_reconciliation 表。
"""

import sqlite3
import logging
from datetime import datetime, date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from config import MATCH_SCORE_THRESHOLD, AMOUNT_TOLERANCE

logger = logging.getLogger(__name__)

# ===================================================================
# 权重与阈值常量
# ===================================================================

# 四项评分权重（合计 = 1.0）
WEIGHT_AMOUNT = 0.40       # 金额权重
WEIGHT_MATERIAL = 0.35     # 物料映射权重
WEIGHT_QUANTITY = 0.15     # 数量权重
WEIGHT_DATE = 0.10         # 日期权重

# 匹配等级阈值
THRESHOLD_AUTO = float(MATCH_SCORE_THRESHOLD)  # 默认 80，自动匹配最低分
THRESHOLD_SUGGEST = 50.0                        # 建议匹配最低分

# 数量容忍度：差异 <= 5% 视为满分
QUANTITY_TOLERANCE_PCT = 0.05

# 日期容忍度：±15 天内有分
DATE_TOLERANCE_DAYS = 15


# ===================================================================
# 工具函数
# ===================================================================

def _to_decimal(value):
    """
    将输入值安全转换为 Decimal。
    支持 int / float / str / Decimal，无法转换时返回 None。
    """
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return None


def _to_date(value):
    """
    将输入值安全转换为 date 对象。
    支持 date / datetime / 'YYYY-MM-DD' 等常见格式字符串。
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
    return None


def _safe_get(row, key, default=None):
    """
    从 sqlite3.Row 中安全取值，键不存在或值为 None 时返回默认值。
    """
    if row is None:
        return default
    try:
        val = row[key]
        return val if val is not None else default
    except (IndexError, KeyError):
        return default


def _round2(value):
    """将浮点数四舍五入保留两位小数。"""
    return round(float(value), 2)


# ===================================================================
# 单项评分函数
# ===================================================================

def calculate_amount_score(inv_amount, stm_amount):
    """
    计算金额匹配分数（0 ~ 100）。

    算法说明:
      - 两者相等（含均为 0） → 100 分
      - 否则按偏差比例线性扣减:
            score = max(0, 100 × (1 − |差额| / max(|inv|, |stm|)))
      - 任一金额无法解析 → 0 分

    Args:
        inv_amount: 发票明细含税金额（amount_excl + tax_amount）
        stm_amount: 对账单明细含税金额（amount_incl_tax）

    Returns:
        float: 0.0 ~ 100.0
    """
    inv = _to_decimal(inv_amount)
    stm = _to_decimal(stm_amount)

    if inv is None or stm is None:
        logger.warning("金额评分: 无法解析金额 inv=%r stm=%r", inv_amount, stm_amount)
        return 0.0

    if inv == stm:
        return 100.0

    if inv == 0 and stm == 0:
        return 100.0

    diff = abs(inv - stm)
    base = max(abs(inv), abs(stm))

    if base == 0:
        return 0.0

    score = max(Decimal("0"), Decimal("100") * (1 - diff / base))
    return float(score.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def calculate_material_score(inv_spec, stm_customer_code, conn):
    """
    计算物料映射匹配分数（0 或 100）。

    通过 sys_material_mapping 表查找映射关系:
      customer_material_code（对账单客户物料编码）
      → supplier_material_code（发票规格 / 供应商物料编码）
    找到映射记录 → 100 分，未找到 → 0 分。

    补充兜底逻辑: 若发票规格与客户物料编码直接相同（忽略大小写），也视为匹配。

    Args:
        inv_spec:          发票上的物料规格 (specification) 或物料名称
        stm_customer_code: 对账单上的客户物料编码 (customer_material_code)
        conn:              sqlite3 数据库连接

    Returns:
        float: 0.0 或 100.0
    """
    if not inv_spec or not stm_customer_code:
        logger.debug(
            "物料评分: 参数缺失 inv_spec=%r stm_customer_code=%r",
            inv_spec, stm_customer_code,
        )
        return 0.0

    inv_spec_str = str(inv_spec).strip()
    stm_code_str = str(stm_customer_code).strip()

    if not inv_spec_str or not stm_code_str:
        return 0.0

    try:
        # 主查询：通过映射表查找
        cursor = conn.execute(
            """
            SELECT 1
              FROM sys_material_mapping
             WHERE customer_material_code = ?
               AND supplier_material_code = ?
             LIMIT 1
            """,
            (stm_code_str, inv_spec_str),
        )
        if cursor.fetchone() is not None:
            logger.debug("物料评分: 映射匹配成功 [%s → %s]", stm_code_str, inv_spec_str)
            return 100.0

        # 兜底：反向查找（发票规格作为 customer_code，对账单编码作为 supplier_code）
        cursor = conn.execute(
            """
            SELECT 1
              FROM sys_material_mapping
             WHERE customer_material_code = ?
               AND supplier_material_code = ?
             LIMIT 1
            """,
            (inv_spec_str, stm_code_str),
        )
        if cursor.fetchone() is not None:
            logger.debug("物料评分: 反向映射匹配成功 [%s ↔ %s]", inv_spec_str, stm_code_str)
            return 100.0

        # 兜底：编码直接相同（忽略大小写）
        if inv_spec_str.upper() == stm_code_str.upper():
            logger.debug("物料评分: 编码直接匹配成功 [%s]", inv_spec_str)
            return 100.0

        logger.debug("物料评分: 未找到映射 [%s → %s]", stm_code_str, inv_spec_str)
        return 0.0

    except sqlite3.Error as exc:
        logger.error("物料评分: 数据库查询异常 - %s", exc)
        return 0.0


def calculate_quantity_score(inv_qty, stm_qty):
    """
    计算数量匹配分数（0 ~ 100）。

    算法说明:
      - 差异百分比 <= 5% → 100 分（完全匹配）
      - 差异百分比 > 5%  → 线性递减:
            score = max(0, 100 × (1 − (diff_pct − 0.05) / 0.45))
        即在 50% 差异时降至 0 分

    Args:
        inv_qty: 发票明细数量
        stm_qty: 对账单明细数量

    Returns:
        float: 0.0 ~ 100.0
    """
    inv = _to_decimal(inv_qty)
    stm = _to_decimal(stm_qty)

    if inv is None or stm is None:
        logger.warning("数量评分: 无法解析数量 inv=%r stm=%r", inv_qty, stm_qty)
        return 0.0

    if inv == 0 and stm == 0:
        return 100.0

    base = max(abs(inv), abs(stm))
    if base == 0:
        # 一方为 0，另一方不为 0 → 完全不匹配
        return 0.0

    diff = abs(inv - stm)
    diff_pct = float(diff / base)

    if diff_pct <= QUANTITY_TOLERANCE_PCT:
        return 100.0

    # 超出容忍范围，线性衰减至 0（diff_pct = 0.50 时归零）
    max_excess = 0.50 - QUANTITY_TOLERANCE_PCT  # 0.45
    excess = diff_pct - QUANTITY_TOLERANCE_PCT
    score = max(0.0, 100.0 * (1.0 - excess / max_excess))
    return _round2(score)


def calculate_date_score(inv_date, stm_date):
    """
    计算日期匹配分数（0 ~ 100）。

    算法说明:
      - 日期差 = 0 天       → 100 分
      - 日期差 1 ~ 15 天    → 线性递减 100 → 50:
            score = 100 − (days / 15) × 50
      - 日期差 16 ~ 30 天   → 线性递减 50 → 0:
            score = 50 − ((days − 15) / 15) × 50
      - 日期差 > 30 天      → 0 分

    Args:
        inv_date: 发票日期（inv_invoice.invoice_date）
        stm_date: 对账单交货日期（stm_statement_item.delivery_date）

    Returns:
        float: 0.0 ~ 100.0
    """
    d_inv = _to_date(inv_date)
    d_stm = _to_date(stm_date)

    if d_inv is None or d_stm is None:
        logger.warning("日期评分: 无法解析日期 inv=%r stm=%r", inv_date, stm_date)
        return 0.0

    days_diff = abs((d_inv - d_stm).days)
    tolerance = DATE_TOLERANCE_DAYS  # 15 天

    if days_diff == 0:
        return 100.0

    if days_diff <= tolerance:
        # 容忍范围内：100 → 50 线性递减
        score = 100.0 - (days_diff / tolerance) * 50.0
        return _round2(score)

    # 超出容忍范围：50 → 0 线性递减（再过 tolerance 天归零）
    excess = days_diff - tolerance
    score = max(0.0, 50.0 - (excess / tolerance) * 50.0)
    return _round2(score)


# ===================================================================
# 综合评分与等级判定
# ===================================================================

def _calculate_weighted_score(amount_score, material_score, quantity_score, date_score):
    """
    根据权重计算加权总分。

    Returns:
        float: 0.0 ~ 100.0
    """
    total = (
        amount_score * WEIGHT_AMOUNT
        + material_score * WEIGHT_MATERIAL
        + quantity_score * WEIGHT_QUANTITY
        + date_score * WEIGHT_DATE
    )
    return _round2(total)


def _determine_match_level(total_score):
    """
    根据总分判定匹配等级。

    Returns:
        str: 'full'（自动匹配）/ 'partial'（建议匹配）/ 'unmatched'（不匹配）
    """
    if total_score >= THRESHOLD_AUTO:
        return "full"
    elif total_score >= THRESHOLD_SUGGEST:
        return "partial"
    else:
        return "unmatched"


# ===================================================================
# 数据查询辅助函数
# ===================================================================

def _fetch_invoice(conn, invoice_id):
    """获取发票主表记录。"""
    cursor = conn.execute(
        """
        SELECT id, invoice_number, invoice_date, total_amount_incl
          FROM inv_invoice
         WHERE id = ?
        """,
        (invoice_id,),
    )
    return cursor.fetchone()


def _fetch_statement(conn, statement_id):
    """获取对账单主表记录。"""
    cursor = conn.execute(
        """
        SELECT id, statement_period, total_invoice_amount, customer_name
          FROM stm_statement
         WHERE id = ?
        """,
        (statement_id,),
    )
    return cursor.fetchone()


def _fetch_invoice_items(conn, invoice_id):
    """
    获取指定发票的全部明细行。
    计算含税金额 = amount_excl + tax_amount。
    """
    cursor = conn.execute(
        """
        SELECT id            AS item_id,
               material_name,
               specification,
               quantity,
               amount_excl,
               tax_amount,
               (amount_excl + COALESCE(tax_amount, 0)) AS amount_incl
          FROM inv_invoice_item
         WHERE invoice_id = ?
         ORDER BY line_number, id
        """,
        (invoice_id,),
    )
    return cursor.fetchall()


def _fetch_statement_items(conn, statement_id):
    """获取指定对账单的全部明细行。"""
    cursor = conn.execute(
        """
        SELECT id                    AS item_id,
               customer_material_code,
               product_name,
               quantity,
               amount_incl_tax,
               delivery_date
          FROM stm_statement_item
         WHERE statement_id = ?
         ORDER BY seq, id
        """,
        (statement_id,),
    )
    return cursor.fetchall()


# ===================================================================
# 对账结果写入
# ===================================================================

def _insert_reconciliation(conn, **kwargs):
    """
    将单条匹配结果写入 rcn_reconciliation 表。

    必需关键字参数:
        invoice_id, statement_id, invoice_item_id, statement_item_id,
        match_score, amount_score, material_score, quantity_score,
        date_score, match_level, difference_amount
    """
    conn.execute(
        """
        INSERT INTO rcn_reconciliation (
            invoice_id,
            statement_id,
            invoice_item_id,
            statement_item_id,
            match_type,
            match_score,
            amount_score,
            material_score,
            quantity_score,
            date_score,
            match_level,
            difference_amount,
            difference_reason,
            is_confirmed,
            created_at
        ) VALUES (?, ?, ?, ?, 'auto', ?, ?, ?, ?, ?, ?, ?, ?, 0, datetime('now','localtime'))
        """,
        (
            kwargs["invoice_id"],
            kwargs["statement_id"],
            kwargs["invoice_item_id"],
            kwargs["statement_item_id"],
            kwargs["match_score"],
            kwargs["amount_score"],
            kwargs["material_score"],
            kwargs["quantity_score"],
            kwargs["date_score"],
            kwargs["match_level"],
            kwargs["difference_amount"],
            kwargs.get("difference_reason"),
        ),
    )


def _build_difference_reason(amount_score, material_score, quantity_score, date_score):
    """
    根据各项评分生成分歧原因描述（仅列出低分项）。

    Returns:
        str or None
    """
    reasons = []
    if amount_score < 80:
        reasons.append(f"金额偏差(得分{amount_score})")
    if material_score < 100:
        reasons.append(f"物料未映射(得分{material_score})")
    if quantity_score < 80:
        reasons.append(f"数量偏差(得分{quantity_score})")
    if date_score < 60:
        reasons.append(f"日期偏差(得分{date_score})")
    return "; ".join(reasons) if reasons else None


# ===================================================================
# 主匹配函数
# ===================================================================

def match_invoice_statement(conn, invoice_id, statement_id):
    """
    执行发票与对账单的自动匹配（REQ-025 核心入口）。

    执行流程:
      1. 校验发票与对账单是否存在
      2. 获取双方明细行
      3. 对所有 (发票明细 × 对账单明细) 交叉进行四级评分
      4. 贪心策略选取最优配对（每次取总分最高的未占用配对）
      5. 未配对的明细行标记为 unmatched
      6. 全部结果写入 rcn_reconciliation 表
      7. 返回汇总结果

    Args:
        conn:         sqlite3 连接（row_factory = sqlite3.Row）
        invoice_id:   发票主表 ID（inv_invoice.id）
        statement_id: 对账单主表 ID（stm_statement.id）

    Returns:
        dict: {
            "invoice_id":       int,
            "statement_id":     int,
            "invoice_number":   str,
            "statement_period": str,
            "total_pairs":      int,   # 最终生成的匹配记录数
            "auto_matched":     int,   # match_level='full' 数量
            "suggested":        int,   # match_level='partial' 数量
            "unmatched":        int,   # match_level='unmatched' 数量
            "results":          list[dict],
        }

    Raises:
        ValueError:    发票或对账单不存在 / 无明细行
        sqlite3.Error: 数据库操作异常
    """

    # ------------------------------------------------------------------
    # 1. 校验主表记录是否存在
    # ------------------------------------------------------------------
    invoice = _fetch_invoice(conn, invoice_id)
    if invoice is None:
        raise ValueError(f"发票不存在: invoice_id={invoice_id}")

    statement = _fetch_statement(conn, statement_id)
    if statement is None:
        raise ValueError(f"对账单不存在: statement_id={statement_id}")

    invoice_number = _safe_get(invoice, "invoice_number", "")
    invoice_date = _safe_get(invoice, "invoice_date")
    statement_period = _safe_get(statement, "statement_period", "")

    # ------------------------------------------------------------------
    # 2. 获取明细行
    # ------------------------------------------------------------------
    inv_items = _fetch_invoice_items(conn, invoice_id)
    stm_items = _fetch_statement_items(conn, statement_id)

    if not inv_items:
        raise ValueError(f"发票无明细行: invoice_id={invoice_id}")
    if not stm_items:
        raise ValueError(f"对账单无明细行: statement_id={statement_id}")

    logger.info(
        "开始匹配: 发票[%s] %d条明细 × 对账单[%s] %d条明细",
        invoice_number, len(inv_items),
        statement_period, len(stm_items),
    )

    # ------------------------------------------------------------------
    # 3. 全量交叉评分（发票明细 × 对账单明细）
    # ------------------------------------------------------------------
    all_pairs = []

    for i, inv_item in enumerate(inv_items):
        # 发票含税金额 = 不含税金额 + 税额
        inv_amount = _safe_get(inv_item, "amount_incl", 0)
        # 物料规格：优先取 specification，其次取 material_name
        inv_spec = _safe_get(inv_item, "specification") or _safe_get(inv_item, "material_name", "")
        inv_qty = _safe_get(inv_item, "quantity", 0)

        for j, stm_item in enumerate(stm_items):
            stm_amount = _safe_get(stm_item, "amount_incl_tax", 0)
            stm_customer_code = _safe_get(stm_item, "customer_material_code", "")
            stm_qty = _safe_get(stm_item, "quantity", 0)
            stm_delivery_date = _safe_get(stm_item, "delivery_date")

            # 四级评分
            amt_score = calculate_amount_score(inv_amount, stm_amount)
            mat_score = calculate_material_score(inv_spec, stm_customer_code, conn)
            qty_score = calculate_quantity_score(inv_qty, stm_qty)
            dt_score = calculate_date_score(invoice_date, stm_delivery_date)

            total = _calculate_weighted_score(amt_score, mat_score, qty_score, dt_score)
            level = _determine_match_level(total)

            # 计算金额差异
            inv_dec = _to_decimal(inv_amount) or Decimal("0")
            stm_dec = _to_decimal(stm_amount) or Decimal("0")
            diff_amount = float((inv_dec - stm_dec).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

            all_pairs.append({
                "total_score": total,
                "inv_idx": i,
                "stm_idx": j,
                "amount_score": amt_score,
                "material_score": mat_score,
                "quantity_score": qty_score,
                "date_score": dt_score,
                "match_level": level,
                "difference_amount": diff_amount,
                "inv_item": inv_item,
                "stm_item": stm_item,
            })

    # ------------------------------------------------------------------
    # 4. 贪心最优匹配（每条明细最多匹配一次，优先选取高分配对）
    # ------------------------------------------------------------------
    all_pairs.sort(key=lambda p: p["total_score"], reverse=True)

    used_inv = set()   # 已匹配的发票明细索引
    used_stm = set()   # 已匹配的对账单明细索引
    matched_pairs = []  # 最终配对结果

    for pair in all_pairs:
        i_idx = pair["inv_idx"]
        j_idx = pair["stm_idx"]

        if i_idx in used_inv or j_idx in used_stm:
            continue

        used_inv.add(i_idx)
        used_stm.add(j_idx)
        matched_pairs.append(pair)

    # ------------------------------------------------------------------
    # 5. 收集未匹配的明细行，标记为 unmatched
    # ------------------------------------------------------------------
    for i, inv_item in enumerate(inv_items):
        if i not in used_inv:
            inv_amount = _safe_get(inv_item, "amount_incl", 0)
            matched_pairs.append({
                "total_score": 0.0,
                "inv_idx": i,
                "stm_idx": None,
                "amount_score": 0.0,
                "material_score": 0.0,
                "quantity_score": 0.0,
                "date_score": 0.0,
                "match_level": "unmatched",
                "difference_amount": float(_to_decimal(inv_amount) or 0),
                "inv_item": inv_item,
                "stm_item": None,
            })

    for j, stm_item in enumerate(stm_items):
        if j not in used_stm:
            stm_amount = _safe_get(stm_item, "amount_incl_tax", 0)
            matched_pairs.append({
                "total_score": 0.0,
                "inv_idx": None,
                "stm_idx": j,
                "amount_score": 0.0,
                "material_score": 0.0,
                "quantity_score": 0.0,
                "date_score": 0.0,
                "match_level": "unmatched",
                "difference_amount": -float(_to_decimal(stm_amount) or 0),
                "inv_item": None,
                "stm_item": stm_item,
            })

    # ------------------------------------------------------------------
    # 6. 写入 rcn_reconciliation 表
    # ------------------------------------------------------------------
    auto_count = 0
    suggest_count = 0
    unmatched_count = 0
    output_results = []

    try:
        for pair in matched_pairs:
            inv_item = pair["inv_item"]
            stm_item = pair["stm_item"]
            level = pair["match_level"]

            # 统计各等级数量
            if level == "full":
                auto_count += 1
            elif level == "partial":
                suggest_count += 1
            else:
                unmatched_count += 1

            inv_item_id = _safe_get(inv_item, "item_id") if inv_item else None
            stm_item_id = _safe_get(stm_item, "item_id") if stm_item else None

            # 生成分歧原因描述
            diff_reason = _build_difference_reason(
                pair["amount_score"],
                pair["material_score"],
                pair["quantity_score"],
                pair["date_score"],
            )

            # 写入数据库
            _insert_reconciliation(
                conn,
                invoice_id=invoice_id,
                statement_id=statement_id,
                invoice_item_id=inv_item_id,
                statement_item_id=stm_item_id,
                match_score=pair["total_score"],
                amount_score=pair["amount_score"],
                material_score=pair["material_score"],
                quantity_score=pair["quantity_score"],
                date_score=pair["date_score"],
                match_level=level,
                difference_amount=pair["difference_amount"],
                difference_reason=diff_reason,
            )

            # 构建返回结果
            output_results.append({
                "invoice_item_id": inv_item_id,
                "statement_item_id": stm_item_id,
                "inv_material": _safe_get(inv_item, "material_name", "") if inv_item else "",
                "stm_product": _safe_get(stm_item, "product_name", "") if stm_item else "",
                "amount_score": pair["amount_score"],
                "material_score": pair["material_score"],
                "quantity_score": pair["quantity_score"],
                "date_score": pair["date_score"],
                "total_score": pair["total_score"],
                "match_level": level,
                "difference_amount": pair["difference_amount"],
                "difference_reason": diff_reason,
            })

        conn.commit()
        logger.info(
            "匹配完成: 发票[%s] ↔ 对账单[%s] → 自动=%d, 建议=%d, 未匹配=%d",
            invoice_number, statement_period,
            auto_count, suggest_count, unmatched_count,
        )

    except sqlite3.Error as exc:
        conn.rollback()
        logger.error("匹配结果写入失败，已回滚: %s", exc)
        raise

    # ------------------------------------------------------------------
    # 7. 返回汇总结果
    # ------------------------------------------------------------------
    return {
        "invoice_id": invoice_id,
        "statement_id": statement_id,
        "invoice_number": invoice_number,
        "statement_period": statement_period,
        "total_pairs": len(matched_pairs),
        "auto_matched": auto_count,
        "suggested": suggest_count,
        "unmatched": unmatched_count,
        "results": output_results,
    }
