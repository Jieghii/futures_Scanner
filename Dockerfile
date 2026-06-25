FROM python:3.10-slim

WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y --no-install-recommends gcc && rm -rf /var/lib/apt/lists/*

# 安装 Python 依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    && pip install --no-cache-dir flask flask-cors gunicorn

# 复制项目文件
COPY futures_scanner.py .
COPY config.json .
COPY web_server.py .
COPY web_dashboard.html .
COPY database/ ./database/
RUN mkdir -p output

# 暴露端口
EXPOSE 5000

# 数据持久化：output 和 database 目录
VOLUME ["/app/output", "/app/database"]

# 启动命令
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "300", "--workers", "1", "--threads", "4", "web_server:app"]
