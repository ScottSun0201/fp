#!/usr/bin/env python3
"""配置模块"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / 'uploads'
DB_PATH = BASE_DIR / 'fp_system.db'
DB_ENGINE = os.environ.get('DB_ENGINE', 'sqlite').lower()
MYSQL_CONFIG = {
    'host': os.environ.get('FP_DB_HOST', '127.0.0.1'),
    'port': int(os.environ.get('FP_DB_PORT', '3306')),
    'user': os.environ.get('FP_DB_USER', ''),
    'password': os.environ.get('FP_DB_PASSWORD', ''),
    'database': os.environ.get('FP_DB_NAME', 'fp'),
    'charset': 'utf8mb4',
}

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
