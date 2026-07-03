FROM python:3.12-slim

WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y --no-install-recommends curl && rm -rf /var/lib/apt/lists/*

# 安装 Python 依赖
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制后端代码
COPY backend/ ./backend/

# 复制前端文件
COPY frontend/ ./frontend/

# 创建上传目录
RUN mkdir -p /app/uploads

WORKDIR /app/backend

EXPOSE 5050

HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD curl -f http://localhost:5050/api/health || exit 1

CMD ["python", "app.py"]
