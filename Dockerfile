FROM python:3.10-slim

WORKDIR /app

# 安装系统依赖（akshare/curl_cffi编译需要）
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    build-essential \
    libcurl4-openssl-dev \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# 安装 Python 依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    && pip install --no-cache-dir flask flask-cors gunicorn

# 复制项目文件
COPY futures_scanner.py .
COPY config.json .
COPY cloud_server.py .
COPY web_dashboard.html .
COPY database/ ./database/
RUN mkdir -p output

# 启动命令：使用 PORT 环境变量（Render自动注入）
CMD exec gunicorn --bind 0.0.0.0:$PORT --timeout 300 --workers 1 --threads 4 cloud_server:app
