#!/usr/bin/env python3
"""
数据库模型与初始化
REQ-022/031/032/033/036 全部表结构
"""
import sqlite3
import json
from datetime import datetime
from contextlib import contextmanager
from config import DB_PATH


@contextmanager
def get_db():
    """上下文管理器获取数据库连接"""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=DELETE")
    conn.execute("PRAGMA foreign_keys=ON")
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
        conn.executescript(SCHEMA_SQL)
        # 插入默认管理员
        import bcrypt
        pw_hash = bcrypt.hashpw(b'admin123', bcrypt.gensalt(rounds=12)).decode()
        conn.execute("""
            INSERT OR IGNORE INTO sys_user (username, password_hash, real_name, role, is_active)
            VALUES (?, ?, ?, ?, 1)
        """, ('admin', pw_hash, '系统管理员', 'admin'))

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

    print("✅ 数据库初始化完成")


SCHEMA_SQL = """
-- ═══ 用户表 ═══
CREATE TABLE IF NOT EXISTS sys_user (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    username        TEXT NOT NULL UNIQUE,
    password_hash   TEXT NOT NULL,
    real_name       TEXT NOT NULL DEFAULT '',
    role            TEXT NOT NULL DEFAULT 'finance_staff'
                    CHECK(role IN ('admin','finance_manager','finance_staff','sales','viewer')),
    is_active       INTEGER NOT NULL DEFAULT 1,
    login_attempts  INTEGER NOT NULL DEFAULT 0,
    locked_until    TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);

-- ═══ 企业信息表 ═══
CREATE TABLE IF NOT EXISTS sys_enterprise (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    enterprise_name TEXT NOT NULL,
    tax_id          TEXT UNIQUE,
    address         TEXT,
    phone           TEXT,
    bank_name       TEXT,
    bank_account    TEXT,
    seal_number     TEXT,
    enterprise_type TEXT NOT NULL DEFAULT 'both'
                    CHECK(enterprise_type IN ('customer','supplier','both')),
    contact_person  TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);

-- ═══ 产品物料表 ═══
CREATE TABLE IF NOT EXISTS sys_material (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    material_code   TEXT NOT NULL UNIQUE,
    material_name   TEXT NOT NULL,
    category        TEXT,
    specification   TEXT,
    unit            TEXT NOT NULL DEFAULT 'PCS',
    tax_rate        REAL NOT NULL DEFAULT 13.0,
    is_active       INTEGER NOT NULL DEFAULT 1,
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);

-- ═══ 料号映射表 REQ-033 ═══
CREATE TABLE IF NOT EXISTS sys_material_mapping (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    enterprise_id           INTEGER REFERENCES sys_enterprise(id),
    customer_material_code  TEXT NOT NULL,
    supplier_material_code  TEXT NOT NULL,
    customer_name           TEXT,
    supplier_name           TEXT,
    spec                    TEXT,
    unit                    TEXT DEFAULT 'PCS',
    tax_rate                TEXT DEFAULT '13%',
    created_at              TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    UNIQUE(customer_material_code, supplier_material_code)
);
CREATE INDEX IF NOT EXISTS idx_mm_customer ON sys_material_mapping(customer_material_code);
CREATE INDEX IF NOT EXISTS idx_mm_supplier ON sys_material_mapping(supplier_material_code);

-- ═══ 发票主表 ═══
CREATE TABLE IF NOT EXISTS inv_invoice (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_number      TEXT UNIQUE,
    invoice_date        TEXT,
    invoice_type        TEXT NOT NULL DEFAULT 'VAT_SPECIAL'
                        CHECK(invoice_type IN ('VAT_SPECIAL','VAT_NORMAL')),
    buyer_name          TEXT,
    buyer_tax_id        TEXT,
    seller_name         TEXT,
    seller_tax_id       TEXT,
    total_amount_excl   REAL NOT NULL DEFAULT 0,
    total_tax           REAL NOT NULL DEFAULT 0,
    total_amount_incl   REAL NOT NULL DEFAULT 0,
    amount_capital      TEXT,
    status              TEXT NOT NULL DEFAULT 'normal'
                        CHECK(status IN ('normal','red_rushed','cancelled')),
    original_invoice_id INTEGER REFERENCES inv_invoice(id),
    source              TEXT NOT NULL DEFAULT 'manual'
                        CHECK(source IN ('manual','ocr','api')),
    pdf_path            TEXT,
    remark              TEXT,
    created_by          INTEGER REFERENCES sys_user(id),
    created_at          TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at          TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_inv_date ON inv_invoice(invoice_date);
CREATE INDEX IF NOT EXISTS idx_inv_status ON inv_invoice(status);

-- ═══ 发票明细行 ═══
CREATE TABLE IF NOT EXISTS inv_invoice_item (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id          INTEGER NOT NULL REFERENCES inv_invoice(id) ON DELETE CASCADE,
    line_number         INTEGER NOT NULL,
    category_prefix     TEXT,
    material_name       TEXT NOT NULL,
    specification       TEXT,
    unit                TEXT NOT NULL,
    quantity            REAL NOT NULL,
    unit_price_excl     REAL NOT NULL,
    amount_excl         REAL NOT NULL,
    tax_rate            REAL NOT NULL DEFAULT 13.0,
    tax_amount          REAL NOT NULL DEFAULT 0,
    material_id         INTEGER REFERENCES sys_material(id)
);
CREATE INDEX IF NOT EXISTS idx_inv_item_inv ON inv_invoice_item(invoice_id);

-- ═══ 对账单主表 REQ-032 ═══
CREATE TABLE IF NOT EXISTS stm_statement (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    statement_period        TEXT NOT NULL,
    statement_date          TEXT,
    customer_name           TEXT NOT NULL,
    customer_tax_id         TEXT,
    supplier_name           TEXT NOT NULL,
    supplier_tax_id         TEXT,
    settlement_days         INTEGER NOT NULL DEFAULT 30,
    opening_balance         REAL NOT NULL DEFAULT 0,
    current_payment         REAL NOT NULL DEFAULT 0,
    closing_balance         REAL NOT NULL DEFAULT 0,
    delivered_unpaid        REAL NOT NULL DEFAULT 0,
    total_invoice_amount    REAL NOT NULL DEFAULT 0,
    total_quantity          INTEGER NOT NULL DEFAULT 0,
    status                  TEXT NOT NULL DEFAULT 'draft'
                            CHECK(status IN ('draft','pending_review','pending_customer','confirmed','archived')),
    prepared_by             INTEGER REFERENCES sys_user(id),
    reviewed_by             INTEGER REFERENCES sys_user(id),
    confirmed_by            INTEGER REFERENCES sys_user(id),
    confirmed_at            TEXT,
    pdf_path                TEXT,
    balance_status          TEXT DEFAULT 'balanced'
                            CHECK(balance_status IN ('balanced','unbalanced')),
    source_file             TEXT,
    version                 INTEGER NOT NULL DEFAULT 1,
    created_at              TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at              TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_stm_period ON stm_statement(statement_period);
CREATE INDEX IF NOT EXISTS idx_stm_customer ON stm_statement(customer_name);

-- ═══ 对账单明细行 ═══
CREATE TABLE IF NOT EXISTS stm_statement_item (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    statement_id            INTEGER NOT NULL REFERENCES stm_statement(id) ON DELETE CASCADE,
    seq                     INTEGER NOT NULL,
    customer_order_no       TEXT,
    customer_material_code  TEXT,
    delivery_no             TEXT,
    delivery_date           TEXT,
    product_name            TEXT NOT NULL,
    quantity                REAL NOT NULL,
    unit                    TEXT NOT NULL DEFAULT 'PCS',
    unit_price_incl_tax     REAL NOT NULL,
    amount_incl_tax         REAL NOT NULL,
    material_id             INTEGER REFERENCES sys_material(id)
);
CREATE INDEX IF NOT EXISTS idx_stm_item_stm ON stm_statement_item(statement_id);

-- ═══ 回款记录表 REQ-036 ═══
CREATE TABLE IF NOT EXISTS stm_payment (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    statement_id    INTEGER REFERENCES stm_statement(id),
    invoice_id      INTEGER REFERENCES inv_invoice(id),
    payment_date    TEXT NOT NULL,
    amount          REAL NOT NULL,
    payment_method  TEXT NOT NULL DEFAULT 'bank_transfer'
                    CHECK(payment_method IN ('bank_transfer','acceptance_bill','cash','other')),
    bill_number     TEXT,
    bill_maturity   TEXT,
    bank_ref_no     TEXT,
    remark          TEXT,
    created_by      INTEGER REFERENCES sys_user(id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_pay_stm ON stm_payment(statement_id);
CREATE INDEX IF NOT EXISTS idx_pay_date ON stm_payment(payment_date);

-- ═══ 对账核销/匹配表 REQ-025 ═══
CREATE TABLE IF NOT EXISTS rcn_reconciliation (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id          INTEGER REFERENCES inv_invoice(id),
    statement_id        INTEGER REFERENCES stm_statement(id),
    invoice_item_id     INTEGER REFERENCES inv_invoice_item(id),
    statement_item_id   INTEGER REFERENCES stm_statement_item(id),
    match_type          TEXT NOT NULL DEFAULT 'auto'
                        CHECK(match_type IN ('auto','manual')),
    match_score         REAL,
    amount_score        REAL,
    material_score      REAL,
    quantity_score      REAL,
    date_score          REAL,
    match_level         TEXT DEFAULT 'unmatched'
                        CHECK(match_level IN ('full','partial','unmatched')),
    difference_amount   REAL DEFAULT 0,
    difference_reason   TEXT,
    is_confirmed        INTEGER NOT NULL DEFAULT 0,
    confirmed_by        TEXT,
    created_at          TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_rcn_inv ON rcn_reconciliation(invoice_id);
CREATE INDEX IF NOT EXISTS idx_rcn_stm ON rcn_reconciliation(statement_id);

-- ═══ 异常记录表 REQ-029 ═══
CREATE TABLE IF NOT EXISTS sys_anomaly (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    anomaly_type    TEXT NOT NULL
                    CHECK(anomaly_type IN ('amount_mismatch','overdue','parse_error','missing_items','duplicate','unmatched')),
    ref_type        TEXT NOT NULL CHECK(ref_type IN ('invoice','statement','payment','match')),
    ref_id          INTEGER,
    description     TEXT NOT NULL,
    severity        TEXT NOT NULL DEFAULT 'warning'
                    CHECK(severity IN ('info','warning','error','critical')),
    status          TEXT NOT NULL DEFAULT 'open'
                    CHECK(status IN ('open','resolved','ignored')),
    resolved_at     TEXT,
    resolved_by     INTEGER REFERENCES sys_user(id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_anom_status ON sys_anomaly(status);
CREATE INDEX IF NOT EXISTS idx_anom_type ON sys_anomaly(anomaly_type);

-- ═══ 审计日志 ═══
CREATE TABLE IF NOT EXISTS sys_audit_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id         INTEGER REFERENCES sys_user(id),
    username        TEXT,
    action          TEXT NOT NULL CHECK(action IN ('CREATE','UPDATE','DELETE','EXPORT','LOGIN','LOGOUT','MATCH')),
    target_type     TEXT,
    target_id       INTEGER,
    old_values      TEXT,
    new_values      TEXT,
    ip_address      TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_audit_user ON sys_audit_log(user_id);
CREATE INDEX IF NOT EXISTS idx_audit_action ON sys_audit_log(action);

-- ═══ 系统配置 ═══
CREATE TABLE IF NOT EXISTS sys_config (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    config_key      TEXT NOT NULL UNIQUE,
    config_value    TEXT NOT NULL,
    description     TEXT,
    updated_at      TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
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
