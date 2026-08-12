# -*- coding: utf-8 -*-
"""
匹配反馈闭环模块

功能：
1. 记录人工匹配反馈（确认/拒绝/手动关联）
2. 基于历史反馈提升匹配引擎准确率
3. 查询反馈统计
"""

import logging
from datetime import datetime
from models import get_db

logger = logging.getLogger(__name__)


def record_feedback(invoice_item_id, statement_item_id, feedback_type, 
                   original_score=None, original_level=None, 
                   feedback_reason=None, created_by=None):
    """
    记录匹配反馈
    
    Args:
        invoice_item_id: 发票明细ID
        statement_item_id: 对账单明细ID
        feedback_type: 反馈类型 ('confirm'/'reject'/'manual_link')
        original_score: 原始匹配分数
        original_level: 原始匹配等级
        feedback_reason: 反馈原因说明
        created_by: 操作用户ID
    
    Returns:
        dict: 反馈记录ID和状态
    """
    try:
        with get_db() as conn:
            # 获取发票明细信息
            inv_item = conn.execute(
                """SELECT material_name, specification, amount_excl + COALESCE(tax_amount, 0) as amount
                   FROM inv_invoice_item WHERE id = ?""",
                (invoice_item_id,)
            ).fetchone() if invoice_item_id else None
            
            # 获取对账单明细信息
            stm_item = conn.execute(
                """SELECT product_name, customer_material_code, amount_incl_tax as amount,
                          (SELECT supplier_code FROM stm_statement WHERE id = statement_id) as supplier_code,
                          (SELECT supplier_name FROM stm_statement WHERE id = statement_id) as supplier_name
                   FROM stm_statement_item WHERE id = ?""",
                (statement_item_id,)
            ).fetchone() if statement_item_id else None
            
            cursor = conn.execute(
                """INSERT INTO match_feedback (
                    invoice_item_id, statement_item_id, feedback_type,
                    original_score, original_level,
                    inv_material_name, inv_specification, inv_amount,
                    stm_product_name, stm_customer_code, stm_amount,
                    supplier_code, supplier_name,
                    feedback_reason, created_by
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    invoice_item_id,
                    statement_item_id,
                    feedback_type,
                    original_score,
                    original_level,
                    inv_item['material_name'] if inv_item else None,
                    inv_item['specification'] if inv_item else None,
                    inv_item['amount'] if inv_item else None,
                    stm_item['product_name'] if stm_item else None,
                    stm_item['customer_material_code'] if stm_item else None,
                    stm_item['amount'] if stm_item else None,
                    stm_item['supplier_code'] if stm_item else None,
                    stm_item['supplier_name'] if stm_item else None,
                    feedback_reason,
                    created_by
                )
            )
            
            feedback_id = cursor.lastrowid
            conn.commit()
            
            logger.info(
                "记录匹配反馈: feedback_id=%d, type=%s, inv_item=%s, stm_item=%s",
                feedback_id, feedback_type, invoice_item_id, statement_item_id
            )
            
            return {
                "feedback_id": feedback_id,
                "status": "success",
                "message": "反馈记录成功"
            }
            
    except Exception as e:
        logger.exception("记录反馈失败: %s", e)
        return {
            "feedback_id": None,
            "status": "error",
            "message": f"记录失败: {str(e)}"
        }


def get_feedback_history(supplier_code=None, limit=50):
    """
    查询反馈历史
    
    Args:
        supplier_code: 供应商编码（可选，不传则查全部）
        limit: 返回条数限制
    
    Returns:
        list: 反馈记录列表
    """
    try:
        with get_db() as conn:
            if supplier_code:
                rows = conn.execute(
                    """SELECT f.*, 
                              u.real_name as operator_name
                       FROM match_feedback f
                       LEFT JOIN sys_user u ON f.created_by = u.id
                       WHERE f.supplier_code = ?
                       ORDER BY f.created_at DESC
                       LIMIT ?""",
                    (supplier_code, limit)
                ).fetchall()
            else:
                rows = conn.execute(
                    """SELECT f.*, 
                              u.real_name as operator_name
                       FROM match_feedback f
                       LEFT JOIN sys_user u ON f.created_by = u.id
                       ORDER BY f.created_at DESC
                       LIMIT ?""",
                    (limit,)
                ).fetchall()
            
            return [dict(row) for row in rows]
            
    except Exception as e:
        logger.exception("查询反馈历史失败: %s", e)
        return []


def get_feedback_statistics(supplier_code=None):
    """
    获取反馈统计信息
    
    Args:
        supplier_code: 供应商编码（可选）
    
    Returns:
        dict: 统计数据
    """
    try:
        with get_db() as conn:
            where_clause = "WHERE f.supplier_code = ?" if supplier_code else ""
            params = (supplier_code,) if supplier_code else ()
            
            # 按类型统计
            type_stats = conn.execute(
                f"""SELECT feedback_type, COUNT(*) as count
                    FROM match_feedback f
                    {where_clause}
                    GROUP BY feedback_type""",
                params
            ).fetchall()
            
            # 总数
            total = conn.execute(
                f"""SELECT COUNT(*) as total FROM match_feedback f {where_clause}""",
                params
            ).fetchone()['total']
            
            # 最近7天反馈数
            recent = conn.execute(
                f"""SELECT COUNT(*) as count FROM match_feedback f
                    {where_clause} {'AND' if where_clause else 'WHERE'} 
                    created_at >= datetime('now', '-7 days', 'localtime')""",
                params
            ).fetchone()['count']
            
            return {
                "total": total,
                "recent_7days": recent,
                "by_type": {row['feedback_type']: row['count'] for row in type_stats}
            }
            
    except Exception as e:
        logger.exception("获取反馈统计失败: %s", e)
        return {"total": 0, "recent_7days": 0, "by_type": {}}


def get_similar_feedback(inv_material_name, inv_specification, stm_customer_code, supplier_code):
    """
    查询相似的历史反馈（用于匹配引擎增强）
    
    查找相同物料名称、规格、客户料号的历史反馈记录
    
    Args:
        inv_material_name: 发票物料名称
        inv_specification: 发票规格
        stm_customer_code: 对账单客户料号
        supplier_code: 供应商编码
    
    Returns:
        list: 相似反馈记录
    """
    try:
        with get_db() as conn:
            # 构建查询条件
            conditions = []
            params = []
            
            if inv_material_name:
                conditions.append("inv_material_name = ?")
                params.append(inv_material_name)
            
            if inv_specification:
                conditions.append("inv_specification = ?")
                params.append(inv_specification)
            
            if stm_customer_code:
                conditions.append("stm_customer_code = ?")
                params.append(stm_customer_code)
            
            if supplier_code:
                conditions.append("supplier_code = ?")
                params.append(supplier_code)
            
            if not conditions:
                return []
            
            where_clause = " AND ".join(conditions)
            
            rows = conn.execute(
                f"""SELECT * FROM match_feedback
                    WHERE {where_clause}
                      AND feedback_type IN ('confirm', 'manual_link')
                    ORDER BY created_at DESC
                    LIMIT 10""",
                params
            ).fetchall()
            
            return [dict(row) for row in rows]
            
    except Exception as e:
        logger.exception("查询相似反馈失败: %s", e)
        return []


def calculate_feedback_bonus(inv_material_name, inv_specification, stm_customer_code, supplier_code):
    """
    基于历史反馈计算匹配加分（0-20分）
    
    如果历史上相同物料组合被人工确认过，给予额外加分
    
    Args:
        inv_material_name: 发票物料名称
        inv_specification: 发票规格
        stm_customer_code: 对账单客户料号
        supplier_code: 供应商编码
    
    Returns:
        float: 加分值（0-20）
    """
    try:
        similar = get_similar_feedback(
            inv_material_name, inv_specification, stm_customer_code, supplier_code
        )
        
        if not similar:
            return 0.0
        
        # 统计确认次数
        confirm_count = sum(1 for f in similar if f['feedback_type'] in ('confirm', 'manual_link'))
        
        if confirm_count == 0:
            return 0.0
        
        # 根据确认次数给予加分（最多20分）
        # 1次=10分, 2次=15分, 3次及以上=20分
        if confirm_count >= 3:
            return 20.0
        elif confirm_count == 2:
            return 15.0
        else:
            return 10.0
            
    except Exception as e:
        logger.exception("计算反馈加分失败: %s", e)
        return 0.0
