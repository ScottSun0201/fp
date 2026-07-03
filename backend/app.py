# -*- coding: utf-8 -*-
"""
FP进销存财务系统 - Flask主应用
REQ-039: 端口统一5050
REQ-042: 集成全部模块 (config/models/invoice_parser/statement_parser/matching_engine/export_utils)
"""
import os
import io
import json
import logging
from datetime import datetime
from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

from config import (
    SECRET_KEY, UPLOAD_DIR, MAX_CONTENT_LENGTH,
    SESSION_LIFETIME_HOURS, DB_PATH
)
from models import init_db, get_db, audit_log, rows_to_list, dict_from_row
from invoice_parser import parse_invoice_pdf
from statement_parser import parse_statement_pdf
from matching_engine import match_invoice_statement
from export_utils import (
    export_invoices_csv, export_invoices_excel,
    export_statements_csv, export_statements_excel,
    export_match_results_csv,
)

# ─── 日志 ───
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s'
)
logger = logging.getLogger('fp-app')

# ─── Flask 初始化 ───
app = Flask(__name__)
app.config['SECRET_KEY'] = SECRET_KEY
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH
CORS(app)

# ─── 启动时初始化数据库 ───
init_db()

# ─── 前端静态文件 ───
FRONTEND_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'frontend')

@app.route('/')
def index():
    """首页 - 返回登录页"""
    return send_from_directory(FRONTEND_DIR, 'login.html')

@app.route('/<path:filename>')
def serve_static(filename):
    """静态文件服务"""
    filepath = os.path.join(FRONTEND_DIR, filename)
    if os.path.isfile(filepath):
        return send_from_directory(FRONTEND_DIR, filename)
    return send_from_directory(FRONTEND_DIR, 'login.html')


# ================================================================
#  工具函数
# ================================================================

def _get_client_ip():
    """获取客户端IP"""
    return request.headers.get('X-Forwarded-For', request.remote_addr)


def _paginate_query(conn, sql, count_sql, page=1, size=20, params=None):
    """通用分页查询"""
    offset = (page - 1) * size
    base_params = list(params or [])
    total = conn.execute(count_sql, base_params).fetchone()[0]
    rows = conn.execute(
        f"{sql} LIMIT ? OFFSET ?", base_params + [size, offset]
    ).fetchall()
    return rows_to_list(rows), total, page


# ================================================================
#  用户 API
# ================================================================

@app.route("/api/users", methods=["GET"])
def get_users():
    """获取所有活跃用户"""
    with get_db() as conn:
        rows = conn.execute(
            "SELECT id, username, real_name, role, is_active, created_at FROM sys_user WHERE is_active=1"
        ).fetchall()
    return jsonify({"code": 0, "data": rows_to_list(rows)})


@app.route("/api/users/<int:user_id>", methods=["GET"])
def get_user(user_id):
    """获取单个用户详情"""
    with get_db() as conn:
        row = conn.execute(
            "SELECT id, username, real_name, role, is_active, created_at FROM sys_user WHERE id=?",
            (user_id,)
        ).fetchone()
    if row:
        return jsonify({"code": 0, "data": dict_from_row(row)})
    return jsonify({"code": 404, "message": "用户不存在"}), 404


@app.route("/api/login", methods=["POST"])
def login():
    """用户登录 — bcrypt密码验证"""
    data = request.get_json()
    if not data:
        return jsonify({"code": 400, "message": "请求体为空"}), 400

    username = data.get("username", "").strip()
    password = data.get("password", "")

    if not username or not password:
        return jsonify({"code": 401, "message": "用户名或密码不能为空"}), 401

    import bcrypt
    with get_db() as conn:
        user = conn.execute(
            "SELECT * FROM sys_user WHERE username=? AND is_active=1", (username,)
        ).fetchone()

        if user and bcrypt.checkpw(password.encode('utf-8'), user['password_hash'].encode('utf-8')):
            # 登录成功，重置失败次数
            conn.execute(
                "UPDATE sys_user SET login_attempts=0, locked_until=NULL WHERE id=?",
                (user['id'],)
            )
            audit_log(conn, user['id'], username, 'LOGIN', 'user', user['id'],
                      ip=_get_client_ip())

            # 生成简易 token（生产环境应使用 JWT）
            import hashlib
            token_raw = f"{user['id']}:{username}:{datetime.now().isoformat()}"
            token = hashlib.sha256(token_raw.encode()).hexdigest()

            return jsonify({
                "code": 0,
                "message": "登录成功",
                "data": {
                    "token": token,
                    "username": username,
                    "real_name": user['real_name'],
                    "role": user['role'],
                    "user_id": user['id'],
                }
            })
        else:
            # 登录失败，累加失败次数
            if user:
                conn.execute(
                    "UPDATE sys_user SET login_attempts = login_attempts + 1 WHERE id=?",
                    (user['id'],)
                )
            return jsonify({"code": 401, "message": "用户名或密码错误"}), 401


# ================================================================
#  发票 API
# ================================================================

@app.route("/api/invoices/upload", methods=["POST"])
def upload_invoice():
    """上传并解析发票PDF"""
    if 'file' not in request.files:
        return jsonify({"code": 400, "message": "未上传文件"}), 400

    file = request.files['file']
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({"code": 400, "message": "仅支持PDF文件"}), 400

    filename = secure_filename(file.filename)
    # 防止重名：加时间戳
    ts = datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f"inv_{ts}_{filename}"
    filepath = str(UPLOAD_DIR / filename)
    file.save(filepath)

    try:
        data = parse_invoice_pdf(filepath)

        if data.get('errors') and not data.get('invoice_number'):
            return jsonify({
                "code": 500,
                "message": f"发票解析失败: {'; '.join(data['errors'])}",
                "data": data
            }), 500

        with get_db() as conn:
            cursor = conn.execute("""
                INSERT INTO inv_invoice (
                    invoice_number, invoice_date, invoice_type,
                    buyer_name, buyer_tax_id, seller_name, seller_tax_id,
                    total_amount_excl, total_tax, total_amount_incl,
                    amount_capital, source, pdf_path
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                data['invoice_number'], data['invoice_date'], data['invoice_type'],
                data['buyer_name'], data['buyer_tax_id'],
                data['seller_name'], data['seller_tax_id'],
                data['total_amount_excl'], data['total_tax'], data['total_amount_incl'],
                data['amount_capital'], 'ocr', filepath
            ))
            invoice_id = cursor.lastrowid

            # 写入明细行
            for idx, item in enumerate(data.get('items', []), start=1):
                conn.execute("""
                    INSERT INTO inv_invoice_item (
                        invoice_id, line_number, category_prefix, material_name,
                        specification, unit, quantity, unit_price_excl,
                        amount_excl, tax_rate, tax_amount
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    invoice_id, idx,
                    item.get('category_prefix', ''),
                    item.get('material_name', ''),
                    item.get('specification', ''),
                    item.get('unit', 'PCS'),
                    item.get('quantity', 0),
                    item.get('unit_price_excl', 0),
                    item.get('amount_excl', 0),
                    item.get('tax_rate', 13.0),
                    item.get('tax_amount', 0),
                ))

            audit_log(conn, None, 'system', 'CREATE', 'invoice', invoice_id,
                      new_values={"invoice_number": data['invoice_number']},
                      ip=_get_client_ip())

        data['id'] = invoice_id
        return jsonify({"code": 0, "message": "发票解析成功", "data": data})

    except Exception as e:
        logger.exception("发票上传处理异常")
        return jsonify({"code": 500, "message": f"服务器错误: {str(e)}"}), 500


@app.route("/api/invoices", methods=["GET"])
def list_invoices():
    """分页查询发票列表"""
    page = request.args.get('page', 1, type=int)
    size = request.args.get('size', 20, type=int)
    status = request.args.get('status', '')
    keyword = request.args.get('keyword', '')

    where_clauses = []
    params = []

    if status:
        where_clauses.append("status = ?")
        params.append(status)
    if keyword:
        where_clauses.append(
            "(invoice_number LIKE ? OR buyer_name LIKE ? OR seller_name LIKE ?)"
        )
        kw = f"%{keyword}%"
        params.extend([kw, kw, kw])

    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""

    with get_db() as conn:
        data, total, page = _paginate_query(
            conn,
            f"SELECT * FROM inv_invoice {where_sql} ORDER BY id DESC",
            f"SELECT count(*) FROM inv_invoice {where_sql}",
            page, size, params
        )

    return jsonify({"code": 0, "data": data, "total": total, "page": page, "size": size})


@app.route("/api/invoices/<int:invoice_id>", methods=["GET"])
def get_invoice(invoice_id):
    """获取发票详情（含明细行）"""
    with get_db() as conn:
        inv = conn.execute("SELECT * FROM inv_invoice WHERE id=?", (invoice_id,)).fetchone()
        if not inv:
            return jsonify({"code": 404, "message": "发票不存在"}), 404

        items = conn.execute(
            "SELECT * FROM inv_invoice_item WHERE invoice_id=? ORDER BY line_number",
            (invoice_id,)
        ).fetchall()

    result = dict_from_row(inv)
    result['items'] = rows_to_list(items)
    return jsonify({"code": 0, "data": result})


@app.route("/api/invoices/<int:invoice_id>", methods=["DELETE"])
def delete_invoice(invoice_id):
    """删除发票（级联删除明细）"""
    with get_db() as conn:
        inv = conn.execute("SELECT id, invoice_number FROM inv_invoice WHERE id=?", (invoice_id,)).fetchone()
        if not inv:
            return jsonify({"code": 404, "message": "发票不存在"}), 404

        conn.execute("DELETE FROM inv_invoice WHERE id=?", (invoice_id,))
        audit_log(conn, None, 'system', 'DELETE', 'invoice', invoice_id,
                  old_values={"invoice_number": inv['invoice_number']},
                  ip=_get_client_ip())

    return jsonify({"code": 0, "message": "删除成功"})


# ================================================================
#  对账单 API
# ================================================================

@app.route("/api/statements/upload", methods=["POST"])
def upload_statement():
    """上传并解析对账单PDF"""
    if 'file' not in request.files:
        return jsonify({"code": 400, "message": "未上传文件"}), 400

    file = request.files['file']
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({"code": 400, "message": "仅支持PDF文件"}), 400

    filename = secure_filename(file.filename)
    ts = datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f"stm_{ts}_{filename}"
    filepath = str(UPLOAD_DIR / filename)
    file.save(filepath)

    try:
        data = parse_statement_pdf(filepath)

        with get_db() as conn:
            cursor = conn.execute("""
                INSERT INTO stm_statement (
                    statement_period, statement_date,
                    customer_name, customer_tax_id,
                    supplier_name, supplier_tax_id,
                    settlement_days,
                    opening_balance, current_payment, closing_balance,
                    delivered_unpaid, total_invoice_amount, total_quantity,
                    balance_status, source_file, pdf_path
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                data.get('statement_month', ''),
                data.get('statement_date', ''),
                data['customer_name'], data['customer_tax_id'],
                data['supplier_name'], data['supplier_tax_id'],
                data.get('settlement_days', 30),
                data['opening_balance'], data['current_payment'],
                data['closing_balance'], data['delivered_unpaid'],
                data['total_invoice_amount'], data['total_quantity'],
                'balanced' if data.get('balance_check', True) else 'unbalanced',
                filename, filepath
            ))
            stmt_id = cursor.lastrowid

            # 写入明细行
            for idx, item in enumerate(data.get('items', []), start=1):
                conn.execute("""
                    INSERT INTO stm_statement_item (
                        statement_id, seq, customer_order_no,
                        customer_material_code, delivery_no, delivery_date,
                        product_name, quantity, unit,
                        unit_price_incl_tax, amount_incl_tax
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    stmt_id, idx,
                    item.get('customer_order_no', ''),
                    item.get('customer_material_code', ''),
                    item.get('delivery_no', ''),
                    item.get('delivery_date', ''),
                    item.get('product_name', ''),
                    item.get('quantity', 0),
                    item.get('unit', 'PCS'),
                    item.get('unit_price_incl_tax', 0),
                    item.get('amount_incl_tax', 0),
                ))

            audit_log(conn, None, 'system', 'CREATE', 'statement', stmt_id,
                      new_values={"period": data.get('statement_month', '')},
                      ip=_get_client_ip())

        data['id'] = stmt_id
        return jsonify({"code": 0, "message": "对账单解析成功", "data": data})

    except Exception as e:
        logger.exception("对账单上传处理异常")
        return jsonify({"code": 500, "message": f"服务器错误: {str(e)}"}), 500


@app.route("/api/statements", methods=["GET"])
def list_statements():
    """分页查询对账单列表"""
    page = request.args.get('page', 1, type=int)
    size = request.args.get('size', 20, type=int)
    status = request.args.get('status', '')
    customer = request.args.get('customer', '')

    where_clauses = []
    params = []

    if status:
        where_clauses.append("status = ?")
        params.append(status)
    if customer:
        where_clauses.append("customer_name LIKE ?")
        params.append(f"%{customer}%")

    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""

    with get_db() as conn:
        data, total, page = _paginate_query(
            conn,
            f"SELECT * FROM stm_statement {where_sql} ORDER BY id DESC",
            f"SELECT count(*) FROM stm_statement {where_sql}",
            page, size, params
        )

    return jsonify({"code": 0, "data": data, "total": total, "page": page, "size": size})


@app.route("/api/statements/<int:stmt_id>", methods=["GET"])
def get_statement(stmt_id):
    """获取对账单详情（含明细行）"""
    with get_db() as conn:
        stmt = conn.execute("SELECT * FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
        if not stmt:
            return jsonify({"code": 404, "message": "对账单不存在"}), 404

        items = conn.execute(
            "SELECT * FROM stm_statement_item WHERE statement_id=? ORDER BY seq",
            (stmt_id,)
        ).fetchall()

    result = dict_from_row(stmt)
    result['items'] = rows_to_list(items)
    return jsonify({"code": 0, "data": result})


@app.route("/api/statements/<int:stmt_id>", methods=["DELETE"])
def delete_statement(stmt_id):
    """删除对账单"""
    with get_db() as conn:
        stmt = conn.execute("SELECT id, statement_period FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
        if not stmt:
            return jsonify({"code": 404, "message": "对账单不存在"}), 404

        conn.execute("DELETE FROM stm_statement WHERE id=?", (stmt_id,))
        audit_log(conn, None, 'system', 'DELETE', 'statement', stmt_id,
                  old_values={"period": stmt['statement_period']},
                  ip=_get_client_ip())

    return jsonify({"code": 0, "message": "删除成功"})


# ================================================================
#  对账核销/匹配 API
# ================================================================

@app.route("/api/match", methods=["POST"])
def run_match():
    """
    执行发票-对账单自动匹配
    REQ-043: 调用 matching_engine.match_invoice_statement
    """
    data = request.get_json()
    if not data:
        return jsonify({"code": 400, "message": "请提供invoice_id和statement_id"}), 400

    invoice_id = data.get('invoice_id')
    statement_id = data.get('statement_id')

    if not invoice_id or not statement_id:
        return jsonify({"code": 400, "message": "invoice_id 和 statement_id 必填"}), 400

    try:
        with get_db() as conn:
            result = match_invoice_statement(conn, invoice_id, statement_id)

        return jsonify({"code": 0, "message": "匹配完成", "data": result})

    except ValueError as e:
        return jsonify({"code": 404, "message": str(e)}), 404
    except Exception as e:
        logger.exception("匹配执行异常")
        return jsonify({"code": 500, "message": f"匹配异常: {str(e)}"}), 500


@app.route("/api/match/batch", methods=["POST"])
def run_batch_match():
    """批量匹配：对所有未匹配的发票和对账单执行匹配"""
    try:
        with get_db() as conn:
            invoices = conn.execute(
                "SELECT id FROM inv_invoice WHERE status='normal'"
            ).fetchall()
            statements = conn.execute(
                "SELECT id FROM stm_statement WHERE status IN ('draft','confirmed')"
            ).fetchall()

            results = []
            for inv in invoices:
                for stmt in statements:
                    try:
                        r = match_invoice_statement(conn, inv['id'], stmt['id'])
                        results.append(r)
                    except ValueError:
                        continue

        total_auto = sum(r['auto_matched'] for r in results)
        total_suggest = sum(r['suggested'] for r in results)
        total_unmatched = sum(r['unmatched'] for r in results)

        return jsonify({
            "code": 0,
            "message": f"批量匹配完成: {len(results)}组",
            "data": {
                "total_pairs": len(results),
                "auto_matched": total_auto,
                "suggested": total_suggest,
                "unmatched": total_unmatched,
                "details": results
            }
        })

    except Exception as e:
        logger.exception("批量匹配异常")
        return jsonify({"code": 500, "message": f"批量匹配异常: {str(e)}"}), 500


@app.route("/api/match/results", methods=["GET"])
def get_match_results():
    """查询匹配结果"""
    page = request.args.get('page', 1, type=int)
    size = request.args.get('size', 20, type=int)
    level = request.args.get('level', '')

    where_clauses = []
    params = []

    if level:
        where_clauses.append("r.match_level = ?")
        params.append(level)

    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""

    with get_db() as conn:
        data, total, page = _paginate_query(
            conn,
            f"""SELECT r.*, 
                       i.invoice_number, i.invoice_date,
                       s.statement_period, s.customer_name
                FROM rcn_reconciliation r
                LEFT JOIN inv_invoice i ON r.invoice_id = i.id
                LEFT JOIN stm_statement s ON r.statement_id = s.id
                {where_sql}
                ORDER BY r.match_score DESC""",
            f"SELECT count(*) FROM rcn_reconciliation r {where_sql}",
            page, size, params
        )

    return jsonify({"code": 0, "data": data, "total": total, "page": page})


@app.route("/api/match/<int:match_id>/confirm", methods=["POST"])
def confirm_match(match_id):
    """人工确认匹配结果"""
    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM rcn_reconciliation WHERE id=?", (match_id,)
        ).fetchone()
        if not row:
            return jsonify({"code": 404, "message": "匹配记录不存在"}), 404

        conn.execute(
            "UPDATE rcn_reconciliation SET is_confirmed=1, confirmed_by=? WHERE id=?",
            ('manual', match_id)
        )
        audit_log(conn, None, 'system', 'MATCH', 'reconciliation', match_id,
                  ip=_get_client_ip())

    return jsonify({"code": 0, "message": "确认成功"})


# ================================================================
#  回款 API
# ================================================================

@app.route("/api/payments", methods=["GET"])
def list_payments():
    """查询回款记录"""
    page = request.args.get('page', 1, type=int)
    size = request.args.get('size', 20, type=int)

    with get_db() as conn:
        data, total, page = _paginate_query(
            conn,
            """SELECT p.*, s.statement_period, s.customer_name
               FROM stm_payment p
               LEFT JOIN stm_statement s ON p.statement_id = s.id
               ORDER BY p.payment_date DESC""",
            "SELECT count(*) FROM stm_payment",
            page, size
        )

    return jsonify({"code": 0, "data": data, "total": total, "page": page})


@app.route("/api/payments", methods=["POST"])
def create_payment():
    """录入回款"""
    data = request.get_json()
    if not data:
        return jsonify({"code": 400, "message": "请求体为空"}), 400

    required = ['statement_id', 'payment_date', 'amount']
    for field in required:
        if not data.get(field):
            return jsonify({"code": 400, "message": f"缺少必填字段: {field}"}), 400

    with get_db() as conn:
        cursor = conn.execute("""
            INSERT INTO stm_payment (
                statement_id, invoice_id, payment_date, amount,
                payment_method, bill_number, bill_maturity, bank_ref_no, remark
            ) VALUES (?,?,?,?,?,?,?,?,?)
        """, (
            data['statement_id'],
            data.get('invoice_id'),
            data['payment_date'],
            data['amount'],
            data.get('payment_method', 'bank_transfer'),
            data.get('bill_number', ''),
            data.get('bill_maturity', ''),
            data.get('bank_ref_no', ''),
            data.get('remark', ''),
        ))

        audit_log(conn, None, 'system', 'CREATE', 'payment', cursor.lastrowid,
                  new_values=data, ip=_get_client_ip())

    return jsonify({"code": 0, "message": "回款录入成功", "data": {"id": cursor.lastrowid}})


# ================================================================
#  企业信息 API
# ================================================================

@app.route("/api/enterprises", methods=["GET"])
def list_enterprises():
    """查询企业列表"""
    page = request.args.get('page', 1, type=int)
    size = request.args.get('size', 20, type=int)
    etype = request.args.get('type', '')

    where_sql = ""
    params = []
    if etype:
        where_sql = "WHERE enterprise_type = ? OR enterprise_type = 'both'"
        params.append(etype)

    with get_db() as conn:
        data, total, page = _paginate_query(
            conn,
            f"SELECT * FROM sys_enterprise {where_sql} ORDER BY id DESC",
            f"SELECT count(*) FROM sys_enterprise {where_sql}",
            page, size, params
        )

    return jsonify({"code": 0, "data": data, "total": total, "page": page})


@app.route("/api/enterprises", methods=["POST"])
def create_enterprise():
    """新建企业"""
    data = request.get_json()
    if not data or not data.get('enterprise_name'):
        return jsonify({"code": 400, "message": "企业名称必填"}), 400

    with get_db() as conn:
        cursor = conn.execute("""
            INSERT INTO sys_enterprise (
                enterprise_name, tax_id, address, phone,
                bank_name, bank_account, seal_number,
                enterprise_type, contact_person
            ) VALUES (?,?,?,?,?,?,?,?,?)
        """, (
            data['enterprise_name'],
            data.get('tax_id', ''),
            data.get('address', ''),
            data.get('phone', ''),
            data.get('bank_name', ''),
            data.get('bank_account', ''),
            data.get('seal_number', ''),
            data.get('enterprise_type', 'both'),
            data.get('contact_person', ''),
        ))
    return jsonify({"code": 0, "message": "创建成功", "data": {"id": cursor.lastrowid}})


# ================================================================
#  物料管理 API
# ================================================================

@app.route("/api/materials", methods=["GET"])
def list_materials():
    """查询物料列表"""
    page = request.args.get('page', 1, type=int)
    size = request.args.get('size', 20, type=int)
    keyword = request.args.get('keyword', '')

    where_sql = ""
    params = []
    if keyword:
        where_sql = "WHERE material_name LIKE ? OR material_code LIKE ?"
        params = [f"%{keyword}%", f"%{keyword}%"]

    with get_db() as conn:
        data, total, page = _paginate_query(
            conn,
            f"SELECT * FROM sys_material {where_sql} ORDER BY id DESC",
            f"SELECT count(*) FROM sys_material {where_sql}",
            page, size, params
        )

    return jsonify({"code": 0, "data": data, "total": total, "page": page})


@app.route("/api/materials", methods=["POST"])
def create_material():
    """新建物料"""
    data = request.get_json()
    if not data or not data.get('material_code') or not data.get('material_name'):
        return jsonify({"code": 400, "message": "物料编码和名称必填"}), 400

    with get_db() as conn:
        try:
            cursor = conn.execute("""
                INSERT INTO sys_material (
                    material_code, material_name, category,
                    specification, unit, tax_rate
                ) VALUES (?,?,?,?,?,?)
            """, (
                data['material_code'],
                data['material_name'],
                data.get('category', ''),
                data.get('specification', ''),
                data.get('unit', 'PCS'),
                data.get('tax_rate', 13.0),
            ))
        except Exception as e:
            return jsonify({"code": 409, "message": f"物料编码已存在: {str(e)}"}), 409

    return jsonify({"code": 0, "message": "创建成功", "data": {"id": cursor.lastrowid}})


# ================================================================
#  料号映射 API
# ================================================================

@app.route("/api/material-mappings", methods=["GET"])
def list_material_mappings():
    """查询料号映射"""
    page = request.args.get('page', 1, type=int)
    size = request.args.get('size', 50, type=int)

    with get_db() as conn:
        data, total, page = _paginate_query(
            conn,
            "SELECT m.*, e.enterprise_name FROM sys_material_mapping m LEFT JOIN sys_enterprise e ON m.enterprise_id = e.id ORDER BY m.id DESC",
            "SELECT count(*) FROM sys_material_mapping",
            page, size
        )

    return jsonify({"code": 0, "data": data, "total": total, "page": page})


@app.route("/api/material-mappings", methods=["POST"])
def create_material_mapping():
    """新建料号映射"""
    data = request.get_json()
    if not data:
        return jsonify({"code": 400, "message": "请求体为空"}), 400

    with get_db() as conn:
        cursor = conn.execute("""
            INSERT INTO sys_material_mapping (
                enterprise_id, customer_material_code, supplier_material_code,
                customer_name, supplier_name, spec, unit, tax_rate
            ) VALUES (?,?,?,?,?,?,?,?)
        """, (
            data.get('enterprise_id'),
            data['customer_material_code'],
            data['supplier_material_code'],
            data.get('customer_name', ''),
            data.get('supplier_name', ''),
            data.get('spec', ''),
            data.get('unit', 'PCS'),
            data.get('tax_rate', '13%'),
        ))

    return jsonify({"code": 0, "message": "映射创建成功", "data": {"id": cursor.lastrowid}})


# ================================================================
#  Dashboard API
# ================================================================

@app.route("/api/dashboard/summary", methods=["GET"])
def dashboard_summary():
    """首页仪表盘汇总数据"""
    with get_db() as conn:
        invoice_count = conn.execute("SELECT count(*) FROM inv_invoice").fetchone()[0]
        invoice_total = conn.execute(
            "SELECT COALESCE(SUM(total_amount_incl),0) FROM inv_invoice"
        ).fetchone()[0]

        stmt_count = conn.execute("SELECT count(*) FROM stm_statement").fetchone()[0]
        payment_total = conn.execute(
            "SELECT COALESCE(SUM(current_payment),0) FROM stm_statement"
        ).fetchone()[0]

        # 匹配统计
        match_full = conn.execute(
            "SELECT count(*) FROM rcn_reconciliation WHERE match_level='full'"
        ).fetchone()[0]
        match_partial = conn.execute(
            "SELECT count(*) FROM rcn_reconciliation WHERE match_level='partial'"
        ).fetchone()[0]
        match_unmatched = conn.execute(
            "SELECT count(*) FROM rcn_reconciliation WHERE match_level='unmatched'"
        ).fetchone()[0]

        # 异常统计
        anomaly_open = conn.execute(
            "SELECT count(*) FROM sys_anomaly WHERE status='open'"
        ).fetchone()[0]

        # 企业数
        enterprise_count = conn.execute("SELECT count(*) FROM sys_enterprise").fetchone()[0]

        # 物料数
        material_count = conn.execute("SELECT count(*) FROM sys_material WHERE is_active=1").fetchone()[0]

    return jsonify({"code": 0, "data": {
        "invoice_count": invoice_count,
        "invoice_total": round(invoice_total, 2),
        "statement_count": stmt_count,
        "payment_total": round(payment_total, 2),
        "unpaid": round(invoice_total - payment_total, 2),
        "match_stats": {
            "auto_matched": match_full,
            "suggested": match_partial,
            "unmatched": match_unmatched,
        },
        "anomaly_open": anomaly_open,
        "enterprise_count": enterprise_count,
        "material_count": material_count,
    }})


@app.route("/api/dashboard/anomalies", methods=["GET"])
def dashboard_anomalies():
    """查询未处理异常"""
    with get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM sys_anomaly WHERE status='open' ORDER BY created_at DESC LIMIT 20"
        ).fetchall()
    return jsonify({"code": 0, "data": rows_to_list(rows)})


# ================================================================
#  导出 API
# ================================================================

@app.route("/api/export/invoices", methods=["GET"])
def export_invoices():
    """导出发票 — 支持 CSV 和 Excel"""
    fmt = request.args.get('format', 'xlsx')

    with get_db() as conn:
        rows = conn.execute("SELECT * FROM inv_invoice ORDER BY id DESC").fetchall()
    invoices = rows_to_list(rows)

    if fmt == 'csv':
        csv_str = export_invoices_csv(invoices)
        return send_file(
            io.BytesIO(csv_str.encode('utf-8-sig')),
            mimetype='text/csv',
            as_attachment=True,
            download_name='invoices.csv'
        )
    else:
        wb = export_invoices_excel(invoices)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='invoices.xlsx'
        )


@app.route("/api/export/statements", methods=["GET"])
def export_statements():
    """导出对账单"""
    fmt = request.args.get('format', 'xlsx')

    with get_db() as conn:
        rows = conn.execute("SELECT * FROM stm_statement ORDER BY id DESC").fetchall()
    statements = rows_to_list(rows)

    if fmt == 'csv':
        csv_str = export_statements_csv(statements)
        return send_file(
            io.BytesIO(csv_str.encode('utf-8-sig')),
            mimetype='text/csv',
            as_attachment=True,
            download_name='statements.csv'
        )
    else:
        wb = export_statements_excel(statements)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='statements.xlsx'
        )


@app.route("/api/export/match-results", methods=["GET"])
def export_match_results():
    """导出匹配结果"""
    fmt = request.args.get('format', 'csv')

    with get_db() as conn:
        rows = conn.execute("""
            SELECT r.*, i.invoice_number, i.total_amount_incl AS invoice_amount,
                   s.statement_period, s.closing_balance AS statement_amount
            FROM rcn_reconciliation r
            LEFT JOIN inv_invoice i ON r.invoice_id = i.id
            LEFT JOIN stm_statement s ON r.statement_id = s.id
            ORDER BY r.match_score DESC
        """).fetchall()
    results = rows_to_list(rows)

    csv_str = export_match_results_csv(results)
    return send_file(
        io.BytesIO(csv_str.encode('utf-8-sig')),
        mimetype='text/csv',
        as_attachment=True,
        download_name='match_results.csv'
    )


# ================================================================
#  系统配置 API
# ================================================================

@app.route("/api/config", methods=["GET"])
def get_config():
    """获取系统配置"""
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM sys_config ORDER BY id").fetchall()
    return jsonify({"code": 0, "data": rows_to_list(rows)})


@app.route("/api/config/<key>", methods=["PUT"])
def update_config(key):
    """更新系统配置"""
    data = request.get_json()
    value = data.get('value', '')

    with get_db() as conn:
        existing = conn.execute(
            "SELECT id FROM sys_config WHERE config_key=?", (key,)
        ).fetchone()
        if existing:
            conn.execute(
                "UPDATE sys_config SET config_value=?, updated_at=datetime('now','localtime') WHERE config_key=?",
                (value, key)
            )
        else:
            conn.execute(
                "INSERT INTO sys_config (config_key, config_value) VALUES (?,?)",
                (key, value)
            )
        audit_log(conn, None, 'system', 'UPDATE', 'config', None,
                  new_values={key: value}, ip=_get_client_ip())

    return jsonify({"code": 0, "message": "配置更新成功"})


# ================================================================
#  审计日志 API
# ================================================================

@app.route("/api/audit-logs", methods=["GET"])
def get_audit_logs():
    """查询审计日志"""
    page = request.args.get('page', 1, type=int)
    size = request.args.get('size', 50, type=int)
    action = request.args.get('action', '')

    where_sql = ""
    params = []
    if action:
        where_sql = "WHERE action = ?"
        params.append(action)

    with get_db() as conn:
        data, total, page = _paginate_query(
            conn,
            f"SELECT * FROM sys_audit_log {where_sql} ORDER BY id DESC",
            f"SELECT count(*) FROM sys_audit_log {where_sql}",
            page, size, params
        )

    return jsonify({"code": 0, "data": data, "total": total, "page": page})


# ================================================================
#  健康检查
# ================================================================

@app.route("/api/health", methods=["GET"])
def health_check():
    """健康检查端点"""
    try:
        with get_db() as conn:
            conn.execute("SELECT 1")
        return jsonify({"code": 0, "status": "healthy", "timestamp": datetime.now().isoformat()})
    except Exception as e:
        return jsonify({"code": 500, "status": "unhealthy", "error": str(e)}), 500


# ================================================================
#  启动 — REQ-039: 端口统一 5050
# ================================================================

if __name__ == "__main__":
    port = int(os.getenv('PORT', 5050))
    logger.info("FP进销存财务系统启动 → port=%d", port)
    debug = os.getenv('FLASK_DEBUG', 'true').lower() == 'true'
    app.run(host="0.0.0.0", port=port, debug=debug)
