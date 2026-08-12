#!/usr/bin/env python3
"""配置模块"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / 'uploads'
DB_ENGINE = os.environ.get('DB_ENGINE', 'mysql').lower()

# ── 环境检测：Docker 容器内 = 107 服务器，否则 = 本地开发 ──
_IN_DOCKER = os.path.exists('/.dockerenv') or os.environ.get('DOCKER_CONTAINER') == '1'

# ── FP 业务库（默认 127.0.0.1，Docker 下也是 127.0.0.1 因为 host 网络） ──
MYSQL_CONFIG = {
    'host': os.environ.get('FP_DB_HOST', '127.0.0.1'),
    'port': int(os.environ.get('FP_DB_PORT', '3306')),
    'user': os.environ.get('FP_DB_USER', ''),
    'password': os.environ.get('FP_DB_PASSWORD', ''),
    'database': os.environ.get('FP_DB_NAME', 'fp'),
    'charset': 'utf8mb4',
}

# ── ERP 库：本地开发 → 107 的 nbgl_test；107 服务器 Docker → 阿里云只读库 ──
if _IN_DOCKER:
    _ERP_DEFAULT_HOST = 'rm-2ze5p26084l3gu9ljzo.mysql.rds.aliyuncs.com'
    _ERP_DEFAULT_USER = 'lwgw_pro_query'
    _ERP_DEFAULT_PASS = 'zchtech_123456'
    _ERP_DEFAULT_DB   = 'nbgl_pro'
else:
    _ERP_DEFAULT_HOST = '192.168.1.107'
    _ERP_DEFAULT_USER = 'nbgl_test'
    _ERP_DEFAULT_PASS = 'J3bNttJSjkdJiNzr'
    _ERP_DEFAULT_DB   = 'nbgl_test'

ERP_MYSQL_CONFIG = {
    'host': os.environ.get('ERP_DB_HOST', _ERP_DEFAULT_HOST),
    'port': int(os.environ.get('ERP_DB_PORT', '3306')),
    'user': os.environ.get('ERP_DB_USER', _ERP_DEFAULT_USER),
    'password': os.environ.get('ERP_DB_PASSWORD', _ERP_DEFAULT_PASS),
    'database': os.environ.get('ERP_DB_NAME', _ERP_DEFAULT_DB),
    'charset': 'utf8mb4',
}
QWEN_OCR_API_KEY = os.environ.get(
    'QWEN_OCR_API_KEY',
    os.environ.get('DASHSCOPE_API_KEY', ''),
)
QWEN_OCR_MODEL = os.environ.get('QWEN_OCR_MODEL', 'qwen3.7-plus')
QWEN_OCR_API_URL = os.environ.get(
    'QWEN_OCR_API_URL',
    'https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions',
)
QWEN_OCR_MAX_PAGES = int(os.environ.get('QWEN_OCR_MAX_PAGES', '20'))

UPLOAD_DIR.mkdir(exist_ok=True)

# Flask
SECRET_KEY = os.environ.get('SECRET_KEY', 'fp-dev-secret-key-change-in-prod')
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB
SESSION_LIFETIME_HOURS = 24

# 业务参数
DEFAULT_TAX_RATE = 13.0
SETTLEMENT_DAYS = 30
AMOUNT_TOLERANCE = 0.02  # 金额容差
MATCH_SCORE_THRESHOLD = 80  # 自动匹配最低分
