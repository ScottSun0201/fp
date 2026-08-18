#!/usr/bin/env python3
"""
数据库模型与初始化
REQ-022/031/032/033/036 全部表结构
"""
import json
import re
from datetime import datetime
from contextlib import contextmanager
from config import DB_ENGINE, MYSQL_CONFIG


class DbRow(dict):
    """Row object that supports both key and numeric access."""
    def __getitem__(self, key):
        if isinstance(key, int):
            return list(self.values())[key]
        return super().__getitem__(key)


class DbCursor:
    def __init__(self, cursor):
        self.cursor = cursor
        self.lastrowid = cursor.lastrowid
        self.rowcount = cursor.rowcount

    def fetchone(self):
        row = self.cursor.fetchone()
        return DbRow(row) if row else None

    def fetchall(self):
        return [DbRow(row) for row in self.cursor.fetchall()]


class MysqlCompatConnection:
    def __init__(self):
        import pymysql
        self.conn = pymysql.connect(
            **MYSQL_CONFIG,
            autocommit=False,
            cursorclass=pymysql.cursors.DictCursor,
        )

    def execute(self, sql, params=None):
        stripped = sql.strip().upper()
        if stripped.startswith("PRAGMA"):
            return DbCursor(self.conn.cursor())
        cursor = self.conn.cursor()
        cursor.execute(_mysql_sql(sql), params)
        return DbCursor(cursor)

    def executescript(self, sql):
        for stmt in _split_sql(sql):
            self.execute(stmt)

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()

    def close(self):
        self.conn.close()


@contextmanager
def get_db():
    """上下文管理器获取数据库连接"""
    conn = MysqlCompatConnection()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """初始化所有表"""
    with get_db() as conn:
        conn.executescript(MYSQL_SCHEMA_SQL)
        _ensure_column(conn, 'stm_statement', 'supplier_code', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'statement_key', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'statement_no', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'reconciliation_key', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'invoice_status', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'overall_status', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'erp_purchase_total', 'REAL DEFAULT 0')
        _ensure_column(conn, 'stm_statement', 'original_filename', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'invoice_number', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'invoice_date', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'invoice_total', 'REAL DEFAULT 0')
        _ensure_column(conn, 'stm_statement', 'invoice_raw_text', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'usage_remark', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'payment_date', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'statement_fingerprint', 'TEXT')
        _ensure_column(conn, 'stm_statement', 'recognition_metadata', 'LONGTEXT')
        if DB_ENGINE == 'mysql':
            # 识别元数据包含动态列映射及逐行错误，VARCHAR(255) 无法容纳真实单据。
            conn.execute(
                "ALTER TABLE stm_statement "
                "MODIFY COLUMN recognition_metadata LONGTEXT NULL"
            )
            # OCR 解析可能返回超长客户/供应商名称，扩大列宽防止 1406 错误
            conn.execute(
                "ALTER TABLE stm_statement "
                "MODIFY COLUMN customer_name VARCHAR(512) NOT NULL"
            )
            conn.execute(
                "ALTER TABLE stm_statement "
                "MODIFY COLUMN supplier_name VARCHAR(512) NOT NULL"
            )
        _ensure_column(conn, 'stm_statement_item', 'specification', 'TEXT')
        _ensure_index(
            conn,
            'stm_statement',
            'idx_stm_supplier_period_unique',
            """
            CREATE UNIQUE INDEX idx_stm_supplier_period_unique
            ON stm_statement(supplier_code, statement_period)
            """,
        )
        # 插入默认管理员
        import bcrypt
        pw_hash = bcrypt.hashpw(b'admin123', bcrypt.gensalt(rounds=12)).decode()
        conn.execute("""
            INSERT OR IGNORE INTO sys_user (username, password_hash, real_name, role, is_active)
            VALUES (?, ?, ?, ?, 1)
        """, ('admin', pw_hash, '系统管理员', 'admin'))
        delivery_pw_hash = bcrypt.hashpw(b'delivery123', bcrypt.gensalt(rounds=12)).decode()
        conn.execute("""
            INSERT OR IGNORE INTO sys_user (username, password_hash, real_name, role, is_active)
            VALUES (?, ?, ?, ?, 1)
        """, ('快递对账', delivery_pw_hash, '快递对账', 'viewer'))
        # 插入默认系统配置
        defaults = [
            ('default_tax_rate', '13.0', '默认税率%'),
            ('settlement_days', '30', '月结天数'),
            ('currency', 'CNY', '币种'),
            ('amount_precision', '2', '金额精度'),
        ]
        for key, val, desc in defaults:
            conn.execute("""
                INSERT OR IGNORE INTO sys_config (config_key, config_value, description)
                VALUES (?, ?, ?)
            """, (key, val, desc))
        conn.execute("""
            CREATE TABLE IF NOT EXISTS delivery_reconciliation_run (
                id INTEGER PRIMARY KEY AUTO_INCREMENT,
                unique_key VARCHAR(255),
                courier_company VARCHAR(255),
                fill_date VARCHAR(64),
                statement_month VARCHAR(32) NOT NULL,
                original_filename VARCHAR(255) NOT NULL,
                file_hash VARCHAR(128),
                statement_count INTEGER NOT NULL DEFAULT 0,
                matched_count INTEGER NOT NULL DEFAULT 0,
                only_statement_count INTEGER NOT NULL DEFAULT 0,
                only_system_count INTEGER NOT NULL DEFAULT 0,
                result_path VARCHAR(1024),
                created_by VARCHAR(255),
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                KEY idx_delivery_run_lookup (statement_month, original_filename),
                KEY idx_delivery_run_unique_key (unique_key),
                KEY idx_delivery_run_created (created_at)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """)
        _ensure_column(conn, 'delivery_reconciliation_run', 'unique_key', 'TEXT')
        _ensure_column(conn, 'delivery_reconciliation_run', 'courier_company', 'TEXT')
        _ensure_column(conn, 'delivery_reconciliation_run', 'fill_date', 'TEXT')
        conn.execute("""
            CREATE TABLE IF NOT EXISTS stm_statement_record (
                id INTEGER PRIMARY KEY AUTO_INCREMENT,
                statement_id INTEGER NOT NULL,
                title VARCHAR(255) NOT NULL,
                record_type VARCHAR(64) NOT NULL DEFAULT 'text',
                record_date VARCHAR(64),
                amount REAL DEFAULT 0,
                text_content TEXT,
                file_path VARCHAR(1024),
                file_name VARCHAR(255),
                created_by VARCHAR(255),
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                KEY idx_statement_record_statement (statement_id),
                KEY idx_statement_record_created (created_at)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """)
        _ensure_column(conn, 'stm_statement_record', 'amount', 'REAL DEFAULT 0')
        conn.execute("""
            CREATE TABLE IF NOT EXISTS stm_statement_allocation (
                id INTEGER PRIMARY KEY AUTO_INCREMENT,
                statement_id INTEGER NOT NULL,
                statement_item_id INTEGER NOT NULL,
                supplier_code VARCHAR(255),
                supplier_name VARCHAR(255),
                line_key VARCHAR(128) NOT NULL,
                purchase_order_id VARCHAR(255),
                material_code VARCHAR(255),
                delivery_date VARCHAR(64),
                unit_price REAL DEFAULT 0,
                current_quantity REAL DEFAULT 0,
                current_amount REAL DEFAULT 0,
                historical_quantity REAL DEFAULT 0,
                historical_amount REAL DEFAULT 0,
                cumulative_quantity REAL DEFAULT 0,
                cumulative_amount REAL DEFAULT 0,
                erp_quantity REAL DEFAULT 0,
                erp_amount REAL DEFAULT 0,
                remaining_quantity REAL DEFAULT 0,
                remaining_amount REAL DEFAULT 0,
                allocation_status VARCHAR(64) NOT NULL DEFAULT 'INFO',
                issue_text TEXT,
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                KEY idx_alloc_statement (statement_id),
                KEY idx_alloc_line_key (line_key),
                KEY idx_alloc_supplier (supplier_code)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS stm_statement_line_override (
                id INTEGER PRIMARY KEY AUTO_INCREMENT,
                statement_id INTEGER NOT NULL,
                statement_item_id INTEGER NOT NULL,
                override_type VARCHAR(64) NOT NULL,
                reason TEXT,
                approved_by VARCHAR(255),
                approved_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE KEY uq_statement_line_override (statement_id, statement_item_id),
                KEY idx_line_override_statement (statement_id)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """)
        # 差异工单表：把每一笔"对不上"变成一张可跟踪的工单（闭环核心）
        conn.execute("""
            CREATE TABLE IF NOT EXISTS stm_diff_ticket (
                id INTEGER PRIMARY KEY AUTO_INCREMENT,
                statement_id INTEGER NOT NULL,
                statement_item_id INTEGER,
                supplier_code VARCHAR(255),
                supplier_name VARCHAR(255),
                line_key VARCHAR(128),
                material_code VARCHAR(255),
                purchase_order_id VARCHAR(255),
                delivery_date VARCHAR(64),
                diff_type VARCHAR(64) NOT NULL DEFAULT 'OTHER',
                diff_amount REAL DEFAULT 0,
                issue_text TEXT,
                status VARCHAR(32) NOT NULL DEFAULT 'OPEN',
                action_type VARCHAR(64),
                action_note TEXT,
                assignee VARCHAR(255),
                source_period VARCHAR(64),
                is_carried_forward INTEGER NOT NULL DEFAULT 0,
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                resolved_at DATETIME,
                KEY idx_diff_statement (statement_id),
                KEY idx_diff_supplier_period (supplier_code, source_period),
                KEY idx_diff_status (status)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """)

    print("✅ 数据库初始化完成")


def _ensure_column(conn, table_name, column_name, column_def):
    """Add a column for existing SQLite databases if it is missing."""
    if DB_ENGINE == 'mysql':
        columns = {
            row['COLUMN_NAME']
            for row in conn.execute("""
                SELECT COLUMN_NAME
                FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ?
            """, (table_name,)).fetchall()
        }
        mysql_def = 'VARCHAR(255)' if column_def.upper() == 'TEXT' else column_def
        if column_name not in columns:
            conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {mysql_def}")
        return
    columns = {row['name'] for row in conn.execute(f"PRAGMA table_info({table_name})")}
    if column_name not in columns:
        conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_def}")


def _ensure_index(conn, table_name, index_name, create_sql):
    exists = conn.execute("""
        SELECT COUNT(*) AS cnt
        FROM information_schema.STATISTICS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = ?
          AND INDEX_NAME = ?
    """, (table_name, index_name)).fetchone()[0]
    if not exists:
        conn.execute(create_sql)


def _drop_index(conn, table_name, index_name):
    exists = conn.execute("""
        SELECT COUNT(*) AS cnt
        FROM information_schema.STATISTICS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = ?
          AND INDEX_NAME = ?
    """, (table_name, index_name)).fetchone()[0]
    if exists:
        conn.execute(f"DROP INDEX {index_name} ON {table_name}")


def _mysql_sql(sql):
    sql = sql.replace("datetime('now','localtime')", "NOW()")
    sql = re.sub(r"\bINSERT\s+OR\s+IGNORE\b", "INSERT IGNORE", sql, flags=re.I)
    return sql.replace("?", "%s")


def _split_sql(sql):
    return [part.strip() for part in sql.split(';') if part.strip()]


MYSQL_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS sys_user (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    username VARCHAR(255) NOT NULL UNIQUE,
    password_hash VARCHAR(255) NOT NULL,
    real_name VARCHAR(255) NOT NULL DEFAULT '',
    role VARCHAR(64) NOT NULL DEFAULT 'finance_staff',
    is_active INTEGER NOT NULL DEFAULT 1,
    login_attempts INTEGER NOT NULL DEFAULT 0,
    locked_until VARCHAR(64),
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS sys_enterprise (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    enterprise_name VARCHAR(255) NOT NULL,
    tax_id VARCHAR(255) UNIQUE,
    address VARCHAR(255),
    phone VARCHAR(255),
    bank_name VARCHAR(255),
    bank_account VARCHAR(255),
    seal_number VARCHAR(255),
    enterprise_type VARCHAR(64) NOT NULL DEFAULT 'both',
    contact_person VARCHAR(255),
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS sys_material (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    material_code VARCHAR(255) NOT NULL UNIQUE,
    material_name VARCHAR(255) NOT NULL,
    category VARCHAR(255),
    specification VARCHAR(255),
    unit VARCHAR(64) NOT NULL DEFAULT 'PCS',
    tax_rate REAL NOT NULL DEFAULT 13.0,
    is_active INTEGER NOT NULL DEFAULT 1,
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS sys_material_mapping (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    enterprise_id INTEGER,
    customer_material_code VARCHAR(255) NOT NULL,
    supplier_material_code VARCHAR(255) NOT NULL,
    customer_name VARCHAR(255),
    supplier_name VARCHAR(255),
    spec VARCHAR(255),
    unit VARCHAR(64) DEFAULT 'PCS',
    tax_rate VARCHAR(64) DEFAULT '13%',
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uniq_material_mapping (customer_material_code, supplier_material_code),
    KEY idx_mm_customer (customer_material_code),
    KEY idx_mm_supplier (supplier_material_code)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS inv_invoice (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    invoice_number VARCHAR(255) UNIQUE,
    invoice_date VARCHAR(64),
    invoice_type VARCHAR(64) NOT NULL DEFAULT 'VAT_SPECIAL',
    buyer_name VARCHAR(255),
    buyer_tax_id VARCHAR(255),
    seller_name VARCHAR(255),
    seller_tax_id VARCHAR(255),
    total_amount_excl REAL NOT NULL DEFAULT 0,
    total_tax REAL NOT NULL DEFAULT 0,
    total_amount_incl REAL NOT NULL DEFAULT 0,
    amount_capital VARCHAR(255),
    status VARCHAR(64) NOT NULL DEFAULT 'normal',
    original_invoice_id INTEGER,
    source VARCHAR(64) NOT NULL DEFAULT 'manual',
    pdf_path VARCHAR(512),
    remark VARCHAR(512),
    created_by INTEGER,
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY idx_inv_date (invoice_date),
    KEY idx_inv_status (status)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS inv_invoice_item (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    invoice_id INTEGER NOT NULL,
    line_number INTEGER NOT NULL,
    category_prefix VARCHAR(255),
    material_name VARCHAR(255) NOT NULL,
    specification VARCHAR(255),
    unit VARCHAR(64) NOT NULL,
    quantity REAL NOT NULL,
    unit_price_excl REAL NOT NULL,
    amount_excl REAL NOT NULL,
    tax_rate REAL NOT NULL DEFAULT 13.0,
    tax_amount REAL NOT NULL DEFAULT 0,
    material_id INTEGER,
    KEY idx_inv_item_inv (invoice_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS stm_statement (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    statement_period VARCHAR(64) NOT NULL,
    statement_date VARCHAR(64),
    customer_name VARCHAR(512) NOT NULL,
    customer_tax_id VARCHAR(255),
    supplier_code VARCHAR(255),
    statement_key VARCHAR(255),
    supplier_name VARCHAR(512) NOT NULL,
    supplier_tax_id VARCHAR(255),
    settlement_days INTEGER NOT NULL DEFAULT 30,
    opening_balance REAL NOT NULL DEFAULT 0,
    current_payment REAL NOT NULL DEFAULT 0,
    closing_balance REAL NOT NULL DEFAULT 0,
    delivered_unpaid REAL NOT NULL DEFAULT 0,
    total_invoice_amount REAL NOT NULL DEFAULT 0,
    total_quantity INTEGER NOT NULL DEFAULT 0,
    status VARCHAR(64) NOT NULL DEFAULT 'draft',
    prepared_by INTEGER,
    reviewed_by INTEGER,
    confirmed_by INTEGER,
    confirmed_at VARCHAR(64),
    pdf_path VARCHAR(512),
    balance_status VARCHAR(64) DEFAULT 'balanced',
    source_file VARCHAR(512),
    version INTEGER NOT NULL DEFAULT 1,
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY idx_stm_period (statement_period),
    KEY idx_stm_customer (customer_name)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS stm_statement_item (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    statement_id INTEGER NOT NULL,
    seq INTEGER NOT NULL,
    customer_order_no VARCHAR(255),
    customer_material_code VARCHAR(255),
    delivery_no VARCHAR(255),
    delivery_date VARCHAR(64),
    product_name VARCHAR(255) NOT NULL,
    specification VARCHAR(255),
    quantity REAL NOT NULL,
    unit VARCHAR(64) NOT NULL DEFAULT 'PCS',
    unit_price_incl_tax REAL NOT NULL,
    amount_incl_tax REAL NOT NULL,
    material_id INTEGER,
    KEY idx_stm_item_stm (statement_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS stm_payment (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    statement_id INTEGER,
    invoice_id INTEGER,
    payment_date VARCHAR(64) NOT NULL,
    amount REAL NOT NULL,
    payment_method VARCHAR(64) NOT NULL DEFAULT 'bank_transfer',
    bill_number VARCHAR(255),
    bill_maturity VARCHAR(64),
    bank_ref_no VARCHAR(255),
    remark VARCHAR(512),
    created_by INTEGER,
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY idx_pay_stm (statement_id),
    KEY idx_pay_date (payment_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS rcn_reconciliation (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    invoice_id INTEGER,
    statement_id INTEGER,
    invoice_item_id INTEGER,
    statement_item_id INTEGER,
    match_type VARCHAR(64) NOT NULL DEFAULT 'auto',
    match_score REAL,
    amount_score REAL,
    material_score REAL,
    quantity_score REAL,
    date_score REAL,
    match_level VARCHAR(64) DEFAULT 'unmatched',
    difference_amount REAL DEFAULT 0,
    difference_reason VARCHAR(512),
    is_confirmed INTEGER NOT NULL DEFAULT 0,
    confirmed_by VARCHAR(255),
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY idx_rcn_inv (invoice_id),
    KEY idx_rcn_stm (statement_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS sys_anomaly (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    anomaly_type VARCHAR(64) NOT NULL,
    ref_type VARCHAR(64) NOT NULL,
    ref_id INTEGER,
    description VARCHAR(1024) NOT NULL,
    severity VARCHAR(64) NOT NULL DEFAULT 'warning',
    status VARCHAR(64) NOT NULL DEFAULT 'open',
    resolved_at VARCHAR(64),
    resolved_by INTEGER,
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY idx_anom_status (status),
    KEY idx_anom_type (anomaly_type)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS sys_audit_log (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    user_id INTEGER,
    username VARCHAR(255),
    action VARCHAR(64) NOT NULL,
    target_type VARCHAR(64),
    target_id INTEGER,
    old_values TEXT,
    new_values TEXT,
    ip_address VARCHAR(128),
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY idx_audit_user (user_id),
    KEY idx_audit_action (action)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS sys_config (
    id INTEGER PRIMARY KEY AUTO_INCREMENT,
    config_key VARCHAR(255) NOT NULL UNIQUE,
    config_value VARCHAR(1024) NOT NULL,
    description VARCHAR(512),
    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
"""


def dict_from_row(row):
    """sqlite3.Row → dict"""
    if row is None:
        return None
    return dict(row)


def rows_to_list(rows):
    """多行 Row → list[dict]"""
    return [dict(r) for r in rows]


def audit_log(conn, user_id, username, action, target_type, target_id, old_values=None, new_values=None, ip=None):
    """写入审计日志"""
    conn.execute("""
        INSERT INTO sys_audit_log (user_id, username, action, target_type, target_id, old_values, new_values, ip_address)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (user_id, username, action, target_type, target_id,
          json.dumps(old_values, ensure_ascii=False) if old_values else None,
          json.dumps(new_values, ensure_ascii=False) if new_values else None,
          ip))
