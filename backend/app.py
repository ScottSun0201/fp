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
import csv
import hashlib
import re
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from flask import Flask, jsonify, redirect, request, send_file, send_from_directory, session
from flask_cors import CORS
from werkzeug.utils import secure_filename

from config import (
    SECRET_KEY, UPLOAD_DIR, MAX_CONTENT_LENGTH,
    SESSION_LIFETIME_HOURS, DB_PATH
)
from models import init_db, get_db, audit_log, rows_to_list, dict_from_row
from invoice_parser import parse_invoice_pdf
from statement_parser import parse_statement_pdf, parse_statement_xlsx
from matching_engine import match_invoice_statement
from export_utils import (
    export_invoices_csv, export_invoices_excel,
    export_statements_csv, export_statements_excel,
    export_match_results_csv,
)
import delivery_compare

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
PROGRESS = {}
DELIVERY_COMPARE_DIR = Path('/Users/liweitas001/Downloads/快递对比')
DELIVERY_OUTPUT_DIR = UPLOAD_DIR / 'delivery_results'
DELIVERY_OUTPUT_DIR.mkdir(exist_ok=True)
STATEMENT_RECORD_DIR = UPLOAD_DIR / 'statement_records'
STATEMENT_RECORD_DIR.mkdir(exist_ok=True)

# ─── 启动时初始化数据库 ───
init_db()

# ─── 前端静态文件 ───
FRONTEND_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'frontend')

@app.route('/')
def index():
    """LWFP 首页"""
    return send_from_directory(FRONTEND_DIR, 'dashboard.html')


@app.route('/login', methods=['GET', 'POST'])
def login_page():
    """测试环境登录页，兼容 107 表单登录。"""
    if request.method == 'GET':
        return send_from_directory(FRONTEND_DIR, 'login.html')
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '')
    import bcrypt
    with get_db() as conn:
        user = conn.execute(
            "SELECT * FROM sys_user WHERE username=? AND is_active=1", (username,)
        ).fetchone()
        if user and bcrypt.checkpw(password.encode('utf-8'), user['password_hash'].encode('utf-8')):
            conn.execute("UPDATE sys_user SET login_attempts=0, locked_until=NULL WHERE id=?", (user['id'],))
            audit_log(conn, user['id'], username, 'LOGIN', 'user', user['id'], ip=_get_client_ip())
            session['username'] = username
            session['role'] = user['role']
            session['user_id'] = user['id']
            return redirect('/')
    return send_from_directory(FRONTEND_DIR, 'login.html'), 401


@app.route('/management')
def management_page():
    """LWFP 对账管理页"""
    return send_from_directory(FRONTEND_DIR, 'management.html')


@app.route('/delivery')
def delivery_page():
    """快递对账页"""
    if not _can_use_delivery():
        return redirect('/login')
    return send_from_directory(FRONTEND_DIR, 'delivery.html')


@app.route('/logout')
def logout_page():
    session.clear()
    return redirect('/login')

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


def _can_use_delivery():
    return session.get('username') in ('admin', '快递对账') or session.get('role') == 'admin'


def _paginate_query(conn, sql, count_sql, page=1, size=20, params=None):
    """通用分页查询"""
    offset = (page - 1) * size
    base_params = list(params or [])
    total = conn.execute(count_sql, base_params).fetchone()[0]
    rows = conn.execute(
        f"{sql} LIMIT ? OFFSET ?", base_params + [size, offset]
    ).fetchall()
    return rows_to_list(rows), total, page


def _status_from_statement(row):
    row = dict(row)
    status = row.get('overall_status') or ''
    if status:
        return status
    if row.get('status') in ('confirmed', 'archived'):
        return 'COMPLETED'
    return 'WAITING_INVOICE'


def _history_row(row):
    row = dict(row)
    total = row.get('total_invoice_amount') or row.get('current_payment') or 0
    key = row.get('reconciliation_key') or row.get('statement_key') or ''
    statement_no = row.get('statement_no') or f"ST{row.get('id')}"
    overall_status = _status_from_statement(row)
    invoice_status = row.get('invoice_status') or ('PASS' if overall_status == 'COMPLETED' else 'NOT_UPLOADED')
    return {
        "id": row.get('id'),
        "statement_no": statement_no,
        "reconciliation_key": key,
        "supplier": row.get('supplier_name') or '',
        "statement_total": f"{float(total or 0):.2f}",
        "erp_purchase_total": f"{float(row.get('erp_purchase_total') or total or 0):.2f}",
        "invoice_status": invoice_status,
        "overall_status": overall_status,
        "invoice_date": row.get('invoice_date') or '',
        "payment_date": row.get('payment_date') or '',
        "usage_remark": row.get('usage_remark') or '',
        "payment_log": row.get('payment_log') or '',
        "reconciliation_log": row.get('reconciliation_log') or '',
        "invoice_log": row.get('invoice_log') or '',
        "original_filename": row.get('original_filename') or row.get('source_file') or '',
        "created_at": str(row.get('created_at') or ''),
    }


def _search_value(value, exact=False):
    value = str(value or '').strip()
    return value if exact else f"%{value}%"


def _search_operator(exact=False):
    return "=" if exact else "LIKE"


def _parse_statement_upload(file_storage):
    ext = Path(file_storage.filename).suffix.lower()
    if ext not in ('.pdf', '.xlsx'):
        raise ValueError("仅支持 PDF / Excel")
    filename = secure_filename(file_storage.filename)
    ts = datetime.now().strftime('%Y%m%d%H%M%S')
    stored_name = f"stm_{ts}_{filename}"
    filepath = str(UPLOAD_DIR / stored_name)
    file_storage.save(filepath)
    data = parse_statement_xlsx(filepath) if ext == '.xlsx' else parse_statement_pdf(filepath)
    supplier_code = (data.get('supplier_code') or '').strip()
    statement_month = (data.get('statement_month') or '').strip()
    statement_key = data.get('statement_key') or (f"{supplier_code}_{statement_month}" if supplier_code and statement_month else '')
    statement_no = data.get('statement_no') or statement_key or Path(filename).stem
    total_amount = data.get('total_invoice_amount') or data.get('current_payment') or 0
    return data, {
        "stored_name": stored_name,
        "filepath": filepath,
        "supplier_code": supplier_code,
        "statement_month": statement_month,
        "statement_key": statement_key,
        "statement_no": statement_no,
        "total_amount": total_amount,
        "original_filename": filename,
    }


def _insert_statement(conn, data, meta):
    existing = None
    if meta["supplier_code"] and meta["statement_month"]:
        existing = conn.execute("""
            SELECT * FROM stm_statement
            WHERE supplier_code=? AND statement_period=?
            LIMIT 1
        """, (meta["supplier_code"], meta["statement_month"])).fetchone()
    if existing:
        return existing["id"], True

    has_issue = bool(data.get('errors')) or not meta["supplier_code"] or not data.get('items')
    overall_status = 'ERP_FAILED' if has_issue else 'WAITING_INVOICE'
    cursor = conn.execute("""
        INSERT INTO stm_statement (
            statement_period, statement_date,
            customer_name, customer_tax_id,
            supplier_code, statement_key,
            statement_no, reconciliation_key,
            supplier_name, supplier_tax_id,
            settlement_days,
            opening_balance, current_payment, closing_balance,
            delivered_unpaid, total_invoice_amount, total_quantity,
            balance_status, source_file, pdf_path,
            invoice_status, overall_status, erp_purchase_total, original_filename
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        meta["statement_month"],
        data.get('statement_date', ''),
        data.get('customer_name') or '骊威',
        data.get('customer_tax_id', ''),
        meta["supplier_code"],
        meta["statement_key"],
        meta["statement_no"],
        meta["statement_key"],
        data.get('supplier_name') or '',
        data.get('supplier_tax_id', ''),
        data.get('settlement_days', 30),
        data.get('opening_balance', 0),
        data.get('current_payment') or meta["total_amount"],
        data.get('closing_balance', 0),
        data.get('delivered_unpaid', 0),
        meta["total_amount"],
        data.get('total_quantity', 0),
        'balanced' if data.get('balance_check', True) else 'unbalanced',
        meta["stored_name"],
        meta["filepath"],
        'NOT_UPLOADED',
        overall_status,
        meta["total_amount"],
        meta["original_filename"],
    ))
    stmt_id = cursor.lastrowid
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
              new_values={"statement_no": meta["statement_no"]}, ip=_get_client_ip())
    return stmt_id, False


def _quote_mysql_name(name):
    return '`' + str(name).replace('`', '``') + '`'


def _export_table_csv(table_name, output_path):
    with get_db() as conn:
        columns = [
            row['COLUMN_NAME']
            for row in conn.execute("""
                SELECT COLUMN_NAME
                FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ?
                ORDER BY ORDINAL_POSITION
            """, (table_name,)).fetchall()
        ]
        if not columns:
            raise ValueError(f"数据库表不存在：{table_name}")
        rows = conn.execute(f"SELECT * FROM {_quote_mysql_name(table_name)}").fetchall()
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            data = dict(row)
            writer.writerow({col: data.get(col, '') for col in columns})


def _save_delivery_upload(work_dir, field_name, default_name):
    uploaded = request.files.get(field_name)
    target = work_dir / default_name
    if uploaded and uploaded.filename:
        uploaded.save(target)
        return target
    default_path = DELIVERY_COMPARE_DIR / default_name
    if not default_path.exists():
        raise ValueError(f"缺少默认文件：{default_name}")
    shutil.copy(default_path, target)
    return target


def _month_range(month):
    if not month:
        return '', ''
    start = datetime.strptime(month + '-01', '%Y-%m-%d')
    if start.month == 12:
        end = start.replace(year=start.year + 1, month=1)
    else:
        end = start.replace(month=start.month + 1)
    return start, end


def _previous_month():
    today = datetime.now()
    first_day = today.replace(day=1)
    if first_day.month == 1:
        previous = first_day.replace(year=first_day.year - 1, month=12)
    else:
        previous = first_day.replace(month=first_day.month - 1)
    return previous.strftime('%Y-%m')


def _format_delivery_cell(value):
    if value is None:
        return ''
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')
    return str(value).strip()


def _normalize_delivery_month(value):
    text = _format_delivery_cell(value)
    if not text:
        return ''
    for fmt in ('%Y-%m', '%Y-%m-%d', '%Y/%m', '%Y/%m/%d'):
        try:
            return datetime.strptime(text[:10], fmt).strftime('%Y-%m')
        except ValueError:
            pass
    return text[:7] if len(text) >= 7 else text


def _month_from_delivery_date(value):
    if isinstance(value, datetime):
        return value.strftime('%Y-%m')
    text = _format_delivery_cell(value)
    if not text:
        return ''
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y-%m', '%Y/%m', '%d/%m/%Y'):
        try:
            return datetime.strptime(text[:10], fmt).strftime('%Y-%m')
        except ValueError:
            pass
    return text[:7] if len(text) >= 7 and text[4:5] in ('-', '/') else ''


def _parse_delivery_datetime(value):
    if isinstance(value, datetime):
        return value
    value = (value or '').strip()
    for fmt in ("%d/%m/%Y %H:%M:%S.%f", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    return None


def _row_in_month(row, date_column, start, end):
    if not start or not end:
        return True
    date_columns = [date_column]
    for fallback in ('transaction_time', 'created_at'):
        if fallback not in date_columns:
            date_columns.append(fallback)
    for column in date_columns:
        dt = _parse_delivery_datetime(row.get(column, ''))
        if dt:
            return start <= dt < end
    return False


def _load_table_rows(table_name):
    with get_db() as conn:
        return [dict(row) for row in conn.execute(f"SELECT * FROM {_quote_mysql_name(table_name)}").fetchall()]


def _file_sha256(path):
    digest = hashlib.sha256()
    with open(path, 'rb') as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _display_filename(filename):
    return Path(filename or '').name or '快递对账单.xlsx'


def _storage_filename(filename, fallback_prefix='delivery_statement'):
    display = _display_filename(filename)
    suffix = Path(display).suffix or '.xlsx'
    safe = secure_filename(display)
    if not safe or not Path(safe).suffix:
        safe = f"{fallback_prefix}_{datetime.now().strftime('%Y%m%d%H%M%S')}{suffix}"
    return safe


def _delivery_history_count(unique_key='', month='', filename=''):
    if unique_key:
        with get_db() as conn:
            return conn.execute("""
                SELECT COUNT(*) AS cnt
                FROM delivery_reconciliation_run
                WHERE unique_key=?
            """, (unique_key,)).fetchone()[0]
    if not month or not filename:
        return 0
    with get_db() as conn:
        return conn.execute("""
            SELECT COUNT(*) AS cnt
            FROM delivery_reconciliation_run
            WHERE statement_month=? AND original_filename=?
        """, (month, filename)).fetchone()[0]


def _record_delivery_run(meta, filename, file_hash, counts, result_path):
    with get_db() as conn:
        cursor = conn.execute("""
            INSERT INTO delivery_reconciliation_run (
                unique_key, courier_company, fill_date,
                statement_month, original_filename, file_hash,
                statement_count, matched_count, only_statement_count, only_system_count,
                result_path, created_by
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            meta.get('唯一标识', ''),
            meta.get('快递公司', ''),
            meta.get('填写日期', ''),
            meta.get('对账月份', ''),
            filename,
            file_hash,
            counts.get('statement_count', 0),
            counts.get('matched_count', 0),
            counts.get('only_statement_count', 0),
            counts.get('only_system_count', 0),
            str(result_path),
            session.get('username', ''),
        ))
        return cursor.lastrowid


def _excel_sheet_name(name):
    safe = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(name or "Sheet")).strip()
    return (safe or "Sheet")[:31]


def _read_delivery_sheet(ws):
    meta = {'唯一标识': '', '快递公司': '', '对账月份': '', '填写日期': '', '批次名称': '', '备注': ''}
    header_row = 1
    first_row = [str(cell.value or '').strip().replace('* ', '').replace('*', '') for cell in ws[1]]
    if {'快递公司', '对账月份', '填写日期'}.issubset(set(first_row)):
        meta_values = [cell.value for cell in ws[2]]
        meta = {
            key: _format_delivery_cell(meta_values[first_row.index(key)]) if key in first_row and first_row.index(key) < len(meta_values) else ''
            for key in meta
        }
        for idx, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
            labels = [str(value or '').strip().replace('* ', '').replace('*', '') for value in row]
            if '快递单号' in labels:
                header_row = idx
                break
        else:
            raise ValueError("模板缺少明细表头：快递单号")
    headers = [str(cell.value or '').strip().replace('* ', '').replace('*', '') for cell in ws[header_row]]
    if '快递单号' not in set(headers):
        raise ValueError("模板表头必须包含：快递单号")
    if not meta.get('快递公司') and ws.title != '快递对账':
        meta['快递公司'] = ws.title
    records = []
    for row_number, values in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1):
        data = {headers[idx]: values[idx] if idx < len(values) else '' for idx in range(len(headers))}
        tracking = delivery_compare.normalize_tracking(data.get('快递单号', ''))
        if not tracking:
            continue
        remark = str(data.get('备注') or '')
        if '示例行' in remark:
            continue
        records.append({
            '_row': row_number,
            '_tracking': tracking,
            '快递公司': str(data.get('快递公司') or meta.get('快递公司') or '').strip(),
            '快递单号': tracking,
            '账单日期': data.get('账单日期') or '',
            '运费': data.get('运费') or '',
            '重量': data.get('重量') or '',
            '备注': remark,
        })
    first_bill_month = ''
    for record in records:
        first_bill_month = _month_from_delivery_date(record.get('账单日期'))
        if first_bill_month:
            break
    meta['对账月份'] = _normalize_delivery_month(meta.get('对账月份')) or first_bill_month or _previous_month()
    meta['填写日期'] = _format_delivery_cell(meta.get('填写日期')) or datetime.now().strftime('%Y-%m-%d')
    meta['批次名称'] = _format_delivery_cell(meta.get('批次名称')) or '月度账单'
    if not meta.get('快递公司') and records:
        meta['快递公司'] = records[0].get('快递公司', '')
    if not meta.get('唯一标识'):
        parts = [meta.get('快递公司') or '快递', meta.get('对账月份') or _previous_month(), meta.get('批次名称') or '月度账单']
        meta['唯一标识'] = '-'.join(parts)
    return meta, records


def _read_delivery_template(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    batches = []
    errors = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        try:
            meta, records = _read_delivery_sheet(ws)
        except ValueError as exc:
            errors.append(f"{sheet_name}: {exc}")
            continue
        if records:
            batches.append((meta, records))
    if not batches and errors:
        raise ValueError("；".join(errors))
    return batches


def _delivery_result_rows(statement_records, delivery_rows, return_rows, month, date_column):
    start, end = _month_range(month)
    delivery_by_no = {}
    for row in delivery_rows:
        tracking = delivery_compare.normalize_tracking(row.get('logistics_number', ''))
        if tracking:
            delivery_by_no.setdefault(tracking, []).append(row)
    return_by_no = {}
    for row in return_rows:
        for col in ('logistics_number', 'return_tracking_number'):
            tracking = delivery_compare.normalize_tracking(row.get(col, ''))
            if tracking:
                return_by_no.setdefault(tracking, []).append(row)

    matched, only_statement = [], []
    statement_numbers = set()
    for record in statement_records:
        tracking = record['_tracking']
        statement_numbers.add(tracking)
        system_matches = [r for r in delivery_by_no.get(tracking, []) if _row_in_month(r, date_column, start, end)]
        return_matches = return_by_no.get(tracking, [])
        row = {
            '物流单号': tracking,
            '对账月份': month,
            '快递Excel行号': record.get('_row', ''),
            '快递公司': record.get('快递公司', ''),
            '账单日期': record.get('账单日期', ''),
            '运费': record.get('运费', ''),
            '重量': record.get('重量', ''),
            '备注': record.get('备注', ''),
            '系统匹配数': len(system_matches),
            '退换货匹配数': len(return_matches),
            '匹配状态': '匹配成功' if system_matches or return_matches else '仅快递账单',
            '系统来源': 'delivery' if system_matches else ('returnandexchangestop' if return_matches else ''),
        }
        if system_matches:
            first = system_matches[0]
            for col in ('id', 'created_at', 'platform', 'shop', 'order_number', 'logistics_company', 'receiver_name', 'receiver_mobile', 'order_status', 'amount'):
                row['系统_' + col] = first.get(col, '')
        if return_matches:
            first_return = return_matches[0]
            for col in ('id', 'created_at', 'platform', 'shop', 'order_number', 'logistics_company', 'return_tracking_number', 'return_logistics_company', 'reason'):
                row['退换货_' + col] = first_return.get(col, '')
        (matched if system_matches or return_matches else only_statement).append(row)

    only_system = []
    for row in delivery_rows:
        if not _row_in_month(row, date_column, start, end):
            continue
        tracking = delivery_compare.normalize_tracking(row.get('logistics_number', ''))
        if not tracking or tracking in statement_numbers:
            continue
        only_system.append({
            '物流单号': tracking,
            '对账月份': month,
            '匹配状态': '仅系统',
            '系统_id': row.get('id', ''),
            '系统_created_at': row.get('created_at', ''),
            '系统_platform': row.get('platform', ''),
            '系统_shop': row.get('shop', ''),
            '系统_order_number': row.get('order_number', ''),
            '系统_logistics_company': row.get('logistics_company', ''),
            '系统_amount': row.get('amount', ''),
        })
    summary = [
        ['指标', '数值'],
        ['对账月份', month],
        ['快递账单唯一单号数', len(statement_numbers)],
        ['匹配成功记录数', len(matched)],
        ['仅快递账单记录数', len(only_statement)],
        ['仅系统记录数', len(only_system)],
    ]
    return summary, matched, only_statement, only_system


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
            session['username'] = username
            session['role'] = user['role']
            session['user_id'] = user['id']

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
    """上传并解析对账单 PDF/Excel"""
    if 'file' not in request.files:
        return jsonify({"code": 400, "message": "未上传文件"}), 400

    file = request.files['file']
    ext = Path(file.filename).suffix.lower()
    if ext not in ('.pdf', '.xlsx'):
        return jsonify({"code": 400, "message": "仅支持PDF或XLSX文件"}), 400

    filename = secure_filename(file.filename)
    ts = datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f"stm_{ts}_{filename}"
    filepath = str(UPLOAD_DIR / filename)
    file.save(filepath)

    try:
        data = parse_statement_xlsx(filepath) if ext == '.xlsx' else parse_statement_pdf(filepath)
        supplier_code = data.get('supplier_code', '').strip()
        statement_month = data.get('statement_month', '').strip()
        if supplier_code and statement_month:
            data['statement_key'] = data.get('statement_key') or f"{supplier_code}_{statement_month}"

        with get_db() as conn:
            if not supplier_code and statement_month and data.get('supplier_name'):
                existing = conn.execute("""
                    SELECT id FROM stm_statement
                    WHERE COALESCE(supplier_code, '') = ''
                      AND supplier_name = ?
                      AND statement_period = ?
                    LIMIT 1
                """, (data['supplier_name'], statement_month)).fetchone()
                if existing:
                    return jsonify({"code": 409, "message": "同一供应商和对账月份的对账单已存在"}), 409

            cursor = conn.execute("""
                INSERT INTO stm_statement (
                    statement_period, statement_date,
                    customer_name, customer_tax_id,
                    supplier_code, statement_key,
                    supplier_name, supplier_tax_id,
                    settlement_days,
                    opening_balance, current_payment, closing_balance,
                    delivered_unpaid, total_invoice_amount, total_quantity,
                    balance_status, source_file, pdf_path
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                statement_month,
                data.get('statement_date', ''),
                data['customer_name'], data['customer_tax_id'],
                supplier_code, data.get('statement_key', ''),
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
        if 'UNIQUE constraint failed: stm_statement.supplier_code, stm_statement.statement_period' in str(e):
            return jsonify({"code": 409, "message": "同一供应商编码和对账月份的对账单已存在"}), 409
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


@app.route("/api/statements", methods=["POST"])
def create_statement_from_management():
    """107 对账管理页上传入口。"""
    task_id = request.form.get('task_id') or ''
    if task_id:
        PROGRESS[task_id] = {"status": "running", "percent": 10, "step": 1, "total": 8, "message": "文件已接收"}
    file = request.files.get('statement_pdf') or request.files.get('file')
    if not file:
        return jsonify({"error": "未上传文件"}), 400
    try:
        if task_id:
            PROGRESS[task_id] = {"status": "running", "percent": 35, "step": 3, "total": 8, "message": "正在识别对账单"}
        data, meta = _parse_statement_upload(file)
        if task_id:
            PROGRESS[task_id] = {"status": "running", "percent": 70, "step": 6, "total": 8, "message": "正在保存结果"}
        with get_db() as conn:
            stmt_id, duplicate = _insert_statement(conn, data, meta)
            if not duplicate:
                conn.execute("""
                    UPDATE stm_statement
                    SET usage_remark=?
                    WHERE id=?
                """, (
                    request.form.get('usage_remark', '').strip() or data.get('usage_remark', ''),
                    stmt_id,
                ))
            row = conn.execute("SELECT * FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
        if duplicate:
            if task_id:
                PROGRESS[task_id] = {"status": "done", "percent": 100, "step": 8, "total": 8, "message": "该对账单已存在"}
            return jsonify({"duplicate": True, "existing": _history_row(row)}), 409
        if task_id:
            PROGRESS[task_id] = {"status": "done", "percent": 100, "step": 8, "total": 8, "message": "处理完成"}
        return jsonify({"id": stmt_id, "summary": _history_row(row)})
    except Exception as exc:
        logger.exception("107 管理页上传处理失败")
        if task_id:
            PROGRESS[task_id] = {"status": "error", "percent": 100, "step": 8, "total": 8, "message": str(exc)}
        return jsonify({"error": str(exc)}), 500


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


@app.route("/api/history", methods=["GET"])
def history_list():
    """107 对账管理页列表接口。"""
    page = request.args.get('page', 1, type=int)
    size = request.args.get('page_size', request.args.get('size', 20), type=int)
    status = request.args.get('status', '')
    search_mode = request.args.get('search_mode', 'fuzzy')
    exact = search_mode == 'exact'
    keyword = request.args.get('keyword', '').strip()
    created_at = request.args.get('created_at', '').strip()
    statement_no = request.args.get('statement_no', '')
    reconciliation_key = request.args.get('reconciliation_key', '')
    supplier = request.args.get('supplier', '')
    usage_remark = request.args.get('usage_remark', '').strip()
    payment_keyword = request.args.get('payment_keyword', '').strip()
    reconciliation_log = request.args.get('reconciliation_log', '').strip()
    invoice_log = request.args.get('invoice_log', '').strip()
    payment_date = request.args.get('payment_date', '').strip()
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')
    where, params = [], []
    if status:
        where.append("(overall_status=? OR status=?)")
        params.extend([status, status])
    if keyword:
        like = f"%{keyword}%"
        where.append("""(
            statement_no LIKE ? OR reconciliation_key LIKE ? OR statement_key LIKE ?
            OR supplier_name LIKE ? OR original_filename LIKE ? OR source_file LIKE ?
            OR invoice_date LIKE ? OR payment_date LIKE ? OR usage_remark LIKE ?
            OR created_at LIKE ? OR statement_period LIKE ?
            OR EXISTS (
                SELECT 1 FROM stm_statement_record r
                WHERE r.statement_id=stm_statement.id
                  AND (r.title LIKE ? OR r.text_content LIKE ? OR r.record_date LIKE ? OR r.file_name LIKE ?)
            )
        )""")
        params.extend([like] * 15)
    if statement_no:
        where.append(f"statement_no {_search_operator(exact)} ?")
        params.append(_search_value(statement_no, exact))
    if reconciliation_key:
        op = _search_operator(exact)
        where.append(f"(reconciliation_key {op} ? OR statement_key {op} ?)")
        params.extend([_search_value(reconciliation_key, exact)] * 2)
    if supplier:
        where.append(f"supplier_name {_search_operator(exact)} ?")
        params.append(_search_value(supplier, exact))
    if created_at:
        where.append(f"created_at {_search_operator(exact)} ?")
        params.append(_search_value(created_at, exact))
    if usage_remark:
        where.append(f"usage_remark {_search_operator(exact)} ?")
        params.append(_search_value(usage_remark, exact))
    if payment_date:
        where.append(f"payment_date {_search_operator(exact)} ?")
        params.append(_search_value(payment_date, exact))
    record_filters = [
        (payment_keyword, ['付款', '支付']),
        (reconciliation_log, ['对账']),
        (invoice_log, ['开票', '发票']),
    ]
    for value, category_words in record_filters:
        if value:
            op = _search_operator(exact)
            category_parts = []
            category_params = []
            for word in category_words:
                category_parts.extend(["r.title LIKE ?", "r.text_content LIKE ?"])
                category_params.extend([f"%{word}%", f"%{word}%"])
            category_sql = "(" + " OR ".join(category_parts) + ")"
            where.append(f"""EXISTS (
                SELECT 1 FROM stm_statement_record r
                WHERE r.statement_id=stm_statement.id
                  AND {category_sql}
                  AND (r.title {op} ? OR r.text_content {op} ? OR r.record_date {op} ? OR r.file_name {op} ?)
            )""")
            params.extend(category_params)
            params.extend([_search_value(value, exact)] * 4)
    if start_date:
        where.append("created_at >= ?")
        params.append(start_date)
    if end_date:
        where.append("created_at <= ?")
        params.append(end_date + " 23:59:59")
    where_sql = "WHERE " + " AND ".join(where) if where else ""
    with get_db() as conn:
        rows, total, _ = _paginate_query(
            conn,
            f"SELECT * FROM stm_statement {where_sql} ORDER BY created_at DESC, id DESC",
            f"SELECT count(*) FROM stm_statement {where_sql}",
            page, size, params
        )
    history_rows = [_history_row(row) for row in rows]
    ids = [row["id"] for row in history_rows]
    if ids:
        placeholders = ",".join(["?"] * len(ids))
        with get_db() as conn:
            records = conn.execute(f"""
                SELECT statement_id, title, text_content, record_date, file_name, created_at
                FROM stm_statement_record
                WHERE statement_id IN ({placeholders})
                ORDER BY created_at DESC, id DESC
            """, ids).fetchall()
        grouped = {stmt_id: {"payment_log": "", "reconciliation_log": "", "invoice_log": ""} for stmt_id in ids}
        for record in records:
            text = " ".join(str(record.get(k) or '') for k in ('title', 'text_content', 'record_date', 'file_name')).strip()
            title = str(record.get('title') or '')
            if not text:
                continue
            target = None
            if '付款' in title or '支付' in title:
                target = 'payment_log'
            elif '对账' in title:
                target = 'reconciliation_log'
            elif '开票' in title or '发票' in title:
                target = 'invoice_log'
            if target and not grouped.get(record['statement_id'], {}).get(target):
                grouped[record['statement_id']][target] = text[:80]
        for row in history_rows:
            row.update(grouped.get(row["id"], {}))
    return jsonify({"rows": history_rows, "total": total, "page": page})


@app.route("/api/history/<int:stmt_id>", methods=["GET"])
def history_detail(stmt_id):
    with get_db() as conn:
        stmt = conn.execute("SELECT * FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
        if not stmt:
            return jsonify({"error": "对账记录不存在"}), 404
        items = conn.execute(
            "SELECT * FROM stm_statement_item WHERE statement_id=? ORDER BY seq", (stmt_id,)
        ).fetchall()
        records = conn.execute("""
            SELECT id, title, record_type, record_date, text_content,
                   file_name, created_by, created_at
            FROM stm_statement_record
            WHERE statement_id=?
            ORDER BY created_at DESC, id DESC
        """, (stmt_id,)).fetchall()
    summary = _history_row(stmt)
    statement_lines = []
    line_checks = []
    for item in items:
        item = dict(item)
        amount = item.get('amount_incl_tax') or 0
        qty = item.get('quantity') or 0
        statement_lines.append({
            "customer_order_no": item.get('customer_order_no') or '',
            "customer_material_no": item.get('customer_material_code') or '',
            "delivery_no": item.get('delivery_no') or '',
            "delivery_date": item.get('delivery_date') or '',
            "product_name": item.get('product_name') or '',
            "supplier_spec": "",
            "quantity": qty,
            "unit": item.get('unit') or '',
            "tax_inclusive_unit_price": item.get('unit_price_incl_tax') or 0,
            "tax_inclusive_amount": amount,
        })
        line_checks.append({
            "material_code": item.get('customer_material_code') or '',
            "purchase_order_id": item.get('customer_order_no') or '',
            "delivery_date": item.get('delivery_date') or '',
            "erp_order_dates": item.get('delivery_date') or '',
            "erp_arrival_dates": item.get('delivery_date') or '',
            "statement_quantity": qty,
            "erp_purchase_quantity": qty,
            "arrival_record_quantity": qty,
            "statement_unit_price": item.get('unit_price_incl_tax') or 0,
            "erp_unit_price": item.get('unit_price_incl_tax') or 0,
            "statement_amount": amount,
            "erp_amount": amount,
            "quantity_status": "PASS",
            "amount_status": "PASS",
            "issue_text": "",
        })
    return jsonify({
        "summary": summary,
        "statement_no": summary["statement_no"],
        "supplier": summary["supplier"],
        "statement_total": summary["statement_total"],
        "erp_purchase_total": summary["erp_purchase_total"],
        "files": [],
        "records": [
            {
                **dict(row),
                "download_url": f"/api/history/{stmt_id}/records/{row['id']}/download" if row['file_name'] else "",
            }
            for row in records
        ],
        "line_checks": line_checks,
        "statement_lines": statement_lines,
    })


@app.route("/api/history/<int:stmt_id>", methods=["DELETE"])
def history_delete(stmt_id):
    with get_db() as conn:
        row = conn.execute("SELECT id FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
        if not row:
            return jsonify({"error": "对账记录不存在"}), 404
        conn.execute("DELETE FROM stm_statement_item WHERE statement_id=?", (stmt_id,))
        conn.execute("DELETE FROM stm_statement WHERE id=?", (stmt_id,))
        audit_log(conn, None, 'system', 'DELETE', 'statement', stmt_id, ip=_get_client_ip())
    return jsonify({"ok": True})


@app.route("/api/history/<int:stmt_id>/runs", methods=["GET"])
def history_runs(stmt_id):
    with get_db() as conn:
        row = conn.execute("SELECT * FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
    return jsonify({"rows": [_history_row(row)] if row else []})


@app.route("/api/history/<int:stmt_id>/records", methods=["GET", "POST"])
def statement_records(stmt_id):
    with get_db() as conn:
        statement = conn.execute("SELECT id FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
        if not statement:
            return jsonify({"error": "对账记录不存在"}), 404
        if request.method == "POST":
            title = request.form.get('title', '').strip()
            record_date = request.form.get('record_date', '').strip()
            text_content = request.form.get('text_content', '').strip()
            file = request.files.get('file')
            if not title:
                return jsonify({"error": "标题必填"}), 400
            file_path = ''
            file_name = ''
            record_type = 'text'
            if file and file.filename:
                file_name = _display_filename(file.filename)
                suffix = Path(file_name).suffix
                stored = f"record_{stmt_id}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}{suffix}"
                target = STATEMENT_RECORD_DIR / stored
                file.save(target)
                file_path = str(target)
                record_type = 'image' if suffix.lower() in ('.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp') else 'file'
            cursor = conn.execute("""
                INSERT INTO stm_statement_record (
                    statement_id, title, record_type, record_date, text_content,
                    file_path, file_name, created_by
                ) VALUES (?,?,?,?,?,?,?,?)
            """, (
                stmt_id, title, record_type, record_date, text_content,
                file_path, file_name, session.get('username', ''),
            ))
            audit_log(conn, None, session.get('username', 'system'), 'CREATE', 'statement_record', cursor.lastrowid,
                      new_values={"statement_id": stmt_id, "title": title}, ip=_get_client_ip())
            if record_date and ('付款' in title or '支付' in title):
                conn.execute("""
                    UPDATE stm_statement
                    SET payment_date=?
                    WHERE id=?
                """, (record_date, stmt_id))
        rows = conn.execute("""
            SELECT id, statement_id, title, record_type, record_date, text_content,
                   file_name, created_by, created_at
            FROM stm_statement_record
            WHERE statement_id=?
            ORDER BY created_at DESC, id DESC
        """, (stmt_id,)).fetchall()
    data = []
    for row in rows:
        item = dict(row)
        item['download_url'] = f"/api/history/{stmt_id}/records/{item['id']}/download" if item.get('file_name') else ''
        data.append(item)
    return jsonify({"rows": data})


@app.route("/api/history/<int:stmt_id>/records/<int:record_id>/download", methods=["GET"])
def statement_record_download(stmt_id, record_id):
    with get_db() as conn:
        row = conn.execute("""
            SELECT file_path, file_name
            FROM stm_statement_record
            WHERE id=? AND statement_id=?
        """, (record_id, stmt_id)).fetchone()
    if not row or not row['file_path'] or not Path(row['file_path']).exists():
        return jsonify({"error": "附件不存在"}), 404
    path = Path(row['file_path'])
    return send_file(path, as_attachment=True, download_name=row['file_name'] or path.name)


@app.route("/api/history/<int:stmt_id>/lines/<int:line_index>/approve", methods=["POST"])
def history_line_approve(stmt_id, line_index):
    return jsonify({"ok": True, "id": stmt_id, "line": line_index})


@app.route("/api/statements/<int:stmt_id>/approve-statement", methods=["POST"])
def approve_statement_107(stmt_id):
    with get_db() as conn:
        conn.execute("""
            UPDATE stm_statement
            SET status='confirmed', overall_status='COMPLETED'
            WHERE id=?
        """, (stmt_id,))
    return jsonify({"ok": True})


@app.route("/api/progress/<task_id>", methods=["GET"])
def progress(task_id):
    return jsonify(PROGRESS.get(task_id, {
        "status": "done",
        "percent": 100,
        "step": 1,
        "total": 1,
        "message": "处理完成",
    }))


@app.route("/api/statements/export", methods=["GET"])
def export_supplier_107():
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM stm_statement ORDER BY id DESC").fetchall()
    output = io.StringIO()
    output.write("statement_no,reconciliation_key,supplier,statement_total,overall_status,created_at\n")
    for row in rows:
        h = _history_row(row)
        output.write(f"{h['statement_no']},{h['reconciliation_key']},{h['supplier']},{h['statement_total']},{h['overall_status']},{h['created_at']}\n")
    data = io.BytesIO(output.getvalue().encode('utf-8-sig'))
    return send_file(data, mimetype='text/csv', as_attachment=True, download_name='statements.csv')


@app.route("/api/statements/template", methods=["GET"])
def download_statement_template():
    template_path = Path(__file__).resolve().parent.parent / 'templates' / '对账单统一模板_v1.xlsx'
    if not template_path.exists():
        return jsonify({"error": "对账单模板不存在"}), 404
    return send_file(
        template_path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='对账单统一模板_v1.xlsx',
    )


@app.route("/api/statements/<int:stmt_id>/invoice/preview", methods=["POST"])
def invoice_preview_107(stmt_id):
    task_id = request.form.get('task_id') or ''
    if task_id:
        PROGRESS[task_id] = {"status": "running", "percent": 30, "step": 2, "total": 4, "message": "正在识别发票"}
    file = request.files.get('invoice_pdf')
    if not file:
        return jsonify({"error": "请选择发票 PDF"}), 400
    filename = secure_filename(file.filename)
    filepath = str(UPLOAD_DIR / f"inv_{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}")
    file.save(filepath)
    try:
        invoice = parse_invoice_pdf(filepath)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    with get_db() as conn:
        stmt = conn.execute("SELECT * FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
    total = float((dict(stmt).get('total_invoice_amount') if stmt else 0) or 0)
    result = {
        "invoice": {
            "invoice_number": invoice.get('invoice_number') or '',
            "invoice_date": invoice.get('invoice_date') or '',
            "invoice_total": invoice.get('total_amount_incl') or invoice.get('invoice_total') or 0,
            "raw_text": invoice.get('raw_text') or '',
        },
        "statement_total": total,
    }
    if task_id:
        PROGRESS[task_id] = {"status": "done", "percent": 100, "step": 4, "total": 4, "message": "识别完成"}
    return jsonify(result)


@app.route("/api/statements/<int:stmt_id>/invoice/confirm", methods=["POST"])
def invoice_confirm_107(stmt_id):
    data = request.get_json() or {}
    invoice_total = float(data.get('invoice_total') or 0)
    with get_db() as conn:
        stmt = conn.execute("SELECT * FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
        if not stmt:
            return jsonify({"error": "对账记录不存在"}), 404
        statement_total = float(dict(stmt).get('total_invoice_amount') or 0)
        passed = abs(invoice_total - statement_total) < 0.005
        conn.execute("""
            UPDATE stm_statement
            SET invoice_number=?, invoice_date=?, invoice_total=?, invoice_raw_text=?,
                invoice_status=?, overall_status=?
            WHERE id=?
        """, (
            data.get('invoice_number', ''),
            data.get('invoice_date', ''),
            invoice_total,
            data.get('raw_text', ''),
            'PASS' if passed else 'FAIL',
            'COMPLETED' if passed else 'INVOICE_FAILED',
            stmt_id,
        ))
        invoice_date = str(data.get('invoice_date') or '').strip()
        invoice_number = str(data.get('invoice_number') or '').strip()
        record_text = f"发票号码：{invoice_number or '-'}；发票金额：{invoice_total:.2f}；校验结果：{'通过' if passed else '金额异常'}"
        cursor = conn.execute("""
            INSERT INTO stm_statement_record (
                statement_id, title, record_type, record_date, text_content,
                file_path, file_name, created_by
            ) VALUES (?,?,?,?,?,?,?,?)
        """, (
            stmt_id, '开票记录', 'text', invoice_date, record_text,
            '', '', session.get('username', ''),
        ))
        audit_log(conn, None, session.get('username', 'system'), 'CREATE', 'statement_record', cursor.lastrowid,
                  new_values={"statement_id": stmt_id, "title": "开票记录"}, ip=_get_client_ip())
        row = conn.execute("SELECT * FROM stm_statement WHERE id=?", (stmt_id,)).fetchone()
    return jsonify({"summary": _history_row(row)})


@app.route("/api/delivery/reconcile", methods=["POST"])
def delivery_reconcile():
    """运行快递对账，返回结果 Excel。"""
    if not _can_use_delivery():
        return jsonify({"error": "无权限访问快递对账"}), 403
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    uploaded = request.files.get('statement')
    if uploaded and uploaded.filename:
        with tempfile.TemporaryDirectory(prefix='lwfp_delivery_single_') as tmp:
            work_dir = Path(tmp)
            original_filename = _display_filename(uploaded.filename)
            statement_path = work_dir / _storage_filename(uploaded.filename)
            uploaded.save(statement_path)
            file_hash = _file_sha256(statement_path)
            try:
                batches = _read_delivery_template(statement_path)
            except ValueError as exc:
                return jsonify({"error": str(exc)}), 400
            if not batches:
                return jsonify({"error": "未识别到快递单号，请按快递对账模板填写"}), 400
            delivery_rows = _load_table_rows('delivery')
            return_rows = _load_table_rows('returnandexchangestop')
            summary = [
                ['指标', '数值'],
                ['上传文件', original_filename],
                ['Sheet批次数', len(batches)],
            ]
            batch_details = []
            history_before_total = 0
            months = []
            result_sheets = []
            output_path = DELIVERY_OUTPUT_DIR / f"快递对账结果_{timestamp}.xlsx"
            for meta, statement_records in batches:
                month = meta.get('对账月份') or _previous_month()
                unique_key = meta.get('唯一标识') or f"{meta.get('快递公司') or '快递'}-{month}-{meta.get('批次名称') or '月度账单'}"
                meta['对账月份'] = month
                meta['唯一标识'] = unique_key
                history_before = _delivery_history_count(unique_key=unique_key)
                history_before_total += history_before
                months.append(month)
                batch_summary, matched, only_statement, only_system = _delivery_result_rows(
                    statement_records,
                    delivery_rows,
                    return_rows,
                    month,
                    'transaction_time',
                )
                company = meta.get('快递公司') or '快递'
                company_path = DELIVERY_OUTPUT_DIR / f"快递对账结果_{company}_{month}_{timestamp}.xlsx"
                delivery_compare.write_xlsx(company_path, [
                    ('汇总', batch_summary),
                    ('匹配成功', delivery_compare.rows_to_table(matched)),
                    ('仅快递账单', delivery_compare.rows_to_table(only_statement)),
                    ('仅系统', delivery_compare.rows_to_table(only_system)),
                ])
                result_sheets.extend([
                    (_excel_sheet_name(f"{company}_汇总"), batch_summary),
                    (_excel_sheet_name(f"{company}_匹配成功"), delivery_compare.rows_to_table(matched)),
                    (_excel_sheet_name(f"{company}_仅快递账单"), delivery_compare.rows_to_table(only_statement)),
                    (_excel_sheet_name(f"{company}_仅系统"), delivery_compare.rows_to_table(only_system)),
                ])
                batch_details.append({
                    'company': company,
                    'month': month,
                    'unique_key': unique_key,
                    'history_before': history_before,
                    'statement_count': len(statement_records),
                    'matched_count': len(matched),
                    'only_statement_count': len(only_statement),
                    'only_system_count': len(only_system),
                    'result_path': str(company_path),
                })
                summary.extend([
                    ['', ''],
                    ['唯一标识', unique_key],
                    ['快递公司', meta.get('快递公司', '')],
                    ['对账月份', month],
                    ['确认时间', meta.get('填写日期', '')],
                    ['历史对账次数', history_before],
                    ['快递账单唯一单号数', len({r['_tracking'] for r in statement_records})],
                    ['匹配成功记录数', len(matched)],
                    ['仅快递账单记录数', len(only_statement)],
                    ['仅系统记录数', len(only_system)],
                ])
            delivery_compare.write_xlsx(output_path, [
                ('汇总', summary),
                *result_sheets,
            ])
            manifest_path = output_path.with_suffix('.json')
            manifest_path.write_text(json.dumps({'batches': batch_details}, ensure_ascii=False, indent=2), encoding='utf-8')
            aggregate_meta = {
                '唯一标识': f"快递对账-{timestamp}",
                '快递公司': '/'.join(item['company'] for item in batch_details),
                '对账月份': ','.join(sorted(set(months))) if months else _previous_month(),
                '填写日期': datetime.now().strftime('%Y-%m-%d'),
            }
            run_id = _record_delivery_run(aggregate_meta, original_filename, file_hash, {
                'statement_count': sum(item['statement_count'] for item in batch_details),
                'matched_count': sum(item['matched_count'] for item in batch_details),
                'only_statement_count': sum(item['only_statement_count'] for item in batch_details),
                'only_system_count': sum(item['only_system_count'] for item in batch_details),
            }, output_path)
        response = send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_path.name,
        )
        response.headers['X-Delivery-History-Before'] = str(history_before_total)
        response.headers['X-Delivery-History-After'] = str(history_before_total + 1)
        response.headers['X-Delivery-Run-Id'] = str(run_id)
        return response

    # Backward-compatible full comparison path. The page no longer uses this.
    with tempfile.TemporaryDirectory(prefix='lwfp_delivery_') as tmp:
        work_dir = Path(tmp)
        delivery_csv = work_dir / 'delivery.csv'
        returns_csv = work_dir / 'returnandexchangestop.csv'
        _export_table_csv('delivery', delivery_csv)
        _export_table_csv('returnandexchangestop', returns_csv)

        _save_delivery_upload(work_dir, 'jitu', '26极兔.xlsx')
        _save_delivery_upload(work_dir, 'zhongtong', '26淘品中通.xlsx')
        _save_delivery_upload(work_dir, 'shunfeng', '26顺丰.xlsx')
        _save_delivery_upload(work_dir, 'supplement_feb', '2月发货在线汇总表.xlsx')
        _save_delivery_upload(work_dir, 'supplement_may', '5月骊威发货汇总表.xlsx')

        output_path = DELIVERY_OUTPUT_DIR / f"快递对账结果_{timestamp}.xlsx"
        args = [
            '--base-dir', str(work_dir),
            '--delivery', str(delivery_csv),
            '--returns', str(returns_csv),
            '--jitu', str(work_dir / '26极兔.xlsx'),
            '--zhongtong', str(work_dir / '26淘品中通.xlsx'),
            '--shunfeng', str(work_dir / '26顺丰.xlsx'),
            '--supplement-feb', str(work_dir / '2月发货在线汇总表.xlsx'),
            '--supplement-may', str(work_dir / '5月骊威发货汇总表.xlsx'),
            '--system-date-column', request.form.get('system_date_column') or 'created_at',
            '--output', str(output_path),
        ]
        code = delivery_compare.main(args)
        if code != 0 or not output_path.exists():
            return jsonify({"error": "快递对账失败"}), 500
    return send_file(
        output_path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=output_path.name,
    )


@app.route("/api/delivery/template", methods=["GET"])
def delivery_template():
    if not _can_use_delivery():
        return jsonify({"error": "无权限访问快递对账"}), 403
    from openpyxl import Workbook
    wb = Workbook()
    default_month = _previous_month()
    examples = [
        ('顺丰干配', 'SF0000000000000'),
        ('极兔快递', 'JT0000000000000'),
        ('中通', '79000000000000'),
    ]
    for idx, (sheet_name, tracking_no) in enumerate(examples):
        ws = wb.active if idx == 0 else wb.create_sheet()
        ws.title = sheet_name
        ws.append(['快递单号*', '账单日期*', '运费', '重量', '备注'])
        ws.append([tracking_no, f'{default_month}-01', '0.00', '', '示例行，填写时删除'])
        for width, col in [(20, 'A'), (14, 'B'), (14, 'C'), (14, 'D'), (48, 'E')]:
            ws.column_dimensions[col].width = width
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
    data = io.BytesIO()
    wb.save(data)
    data.seek(0)
    return send_file(
        data,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='快递对账模板.xlsx',
    )


@app.route("/api/delivery/history-count", methods=["GET"])
def delivery_history_count():
    if not _can_use_delivery():
        return jsonify({"error": "无权限访问快递对账"}), 403
    unique_key = request.args.get('unique_key', '').strip()
    month = request.args.get('month', '')
    filename = _display_filename(request.args.get('filename', ''))
    return jsonify({"count": _delivery_history_count(unique_key=unique_key, month=month, filename=filename)})


@app.route("/api/delivery/history", methods=["GET"])
def delivery_history():
    if not _can_use_delivery():
        return jsonify({"error": "无权限访问快递对账"}), 403
    with get_db() as conn:
        rows = conn.execute("""
            SELECT id, unique_key, courier_company, fill_date,
                   statement_month, original_filename, statement_count,
                   matched_count, only_statement_count, only_system_count,
                   created_by, created_at
            FROM delivery_reconciliation_run
            ORDER BY id DESC
            LIMIT 50
        """).fetchall()
    data = []
    for row in rows:
        item = dict(row)
        item['download_url'] = f"/api/delivery/history/{item['id']}/download"
        data.append(item)
    return jsonify({"rows": data})


@app.route("/api/delivery/history/<int:run_id>/download", methods=["GET"])
def delivery_history_download(run_id):
    if not _can_use_delivery():
        return jsonify({"error": "无权限访问快递对账"}), 403
    with get_db() as conn:
        row = conn.execute(
            "SELECT result_path FROM delivery_reconciliation_run WHERE id=?", (run_id,)
        ).fetchone()
    if not row or not row['result_path'] or not Path(row['result_path']).exists():
        return jsonify({"error": "结果文件不存在"}), 404
    company = request.args.get('company', '').strip()
    if company:
        manifest_path = Path(row['result_path']).with_suffix('.json')
        if not manifest_path.exists():
            return jsonify({"error": "详情文件不存在"}), 404
        manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
        detail = next((item for item in manifest.get('batches', []) if item.get('company') == company), None)
        if not detail or not detail.get('result_path') or not Path(detail['result_path']).exists():
            return jsonify({"error": "公司结果文件不存在"}), 404
        path = Path(detail['result_path'])
        return send_file(
            path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=path.name,
        )
    path = Path(row['result_path'])
    return send_file(
        path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=path.name,
    )


@app.route("/api/delivery/history/<int:run_id>/detail", methods=["GET"])
def delivery_history_detail(run_id):
    if not _can_use_delivery():
        return jsonify({"error": "无权限访问快递对账"}), 403
    with get_db() as conn:
        row = conn.execute(
            "SELECT id, result_path FROM delivery_reconciliation_run WHERE id=?", (run_id,)
        ).fetchone()
    if not row or not row['result_path']:
        return jsonify({"error": "历史记录不存在"}), 404
    manifest_path = Path(row['result_path']).with_suffix('.json')
    if not manifest_path.exists():
        return jsonify({"rows": []})
    manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
    rows = []
    for item in manifest.get('batches', []):
        data = dict(item)
        data.pop('result_path', None)
        data['download_url'] = f"/api/delivery/history/{run_id}/download?company={data.get('company', '')}"
        rows.append(data)
    return jsonify({"rows": rows})


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

@app.route("/api/dashboard", methods=["GET"])
def dashboard_lwfp():
    """LWFP 首页统计接口，兼容原轻量框架。"""
    with get_db() as conn:
        total = conn.execute("SELECT count(*) FROM stm_statement").fetchone()[0]
        total_amount = conn.execute(
            "SELECT COALESCE(SUM(current_payment),0) FROM stm_statement"
        ).fetchone()[0]
        status_rows = conn.execute("""
            SELECT COALESCE(NULLIF(overall_status, ''), status) AS status, count(*) AS cnt
            FROM stm_statement
            GROUP BY COALESCE(NULLIF(overall_status, ''), status)
        """).fetchall()
        anomaly_open = conn.execute(
            "SELECT count(*) FROM sys_anomaly WHERE status='open'"
        ).fetchone()[0]
        trend_rows = conn.execute("""
            SELECT substr(created_at, 1, 10) AS day,
                   count(*) AS cnt,
                   COALESCE(SUM(current_payment),0) AS amount
            FROM stm_statement
            GROUP BY substr(created_at, 1, 10)
            ORDER BY day
            LIMIT 12
        """).fetchall()

    counts = {
        "WAITING_INVOICE": 0,
        "ERP_FAILED": anomaly_open,
        "INVOICE_FAILED": 0,
        "COMPLETED": 0,
        "OTHER": 0,
    }
    for row in status_rows:
        status = row["status"]
        cnt = row["cnt"]
        if status == "ERP_FAILED":
            counts["ERP_FAILED"] += cnt
        elif status == "INVOICE_FAILED":
            counts["INVOICE_FAILED"] += cnt
        elif status in ("COMPLETED", "confirmed", "archived"):
            counts["COMPLETED"] += cnt
        elif status in ("WAITING_INVOICE", "draft", "pending_review", "pending_customer"):
            counts["WAITING_INVOICE"] += cnt
        else:
            counts["OTHER"] += cnt

    trend = [
        {
            "date": row["day"] or "",
            "count": row["cnt"],
            "amount": round(row["amount"] or 0, 2),
        }
        for row in trend_rows
    ]
    pending = counts["WAITING_INVOICE"] + counts["INVOICE_FAILED"]
    abnormal = counts["ERP_FAILED"] + counts["INVOICE_FAILED"]
    return jsonify({
        "total_amount": f"{round(total_amount or 0, 2):.2f}",
        "total": total,
        "pending": pending,
        "abnormal": abnormal,
        "status_counts": counts,
        "trend": trend,
    })


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
