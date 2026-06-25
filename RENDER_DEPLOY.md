# Render 部署指南

## 方案一：Docker 方式（推荐修复后使用）

在 Render 后台 → New → Blueprint → 连接 GitHub 仓库。

## 方案二：Python 原生方式（更简单，推荐）

不需要 Dockerfile 和 render.yaml，直接用 Render 的 Python 环境：

1. 在 Render 后台点 **New** → **Web Service**
2. 连接 GitHub 仓库
3. 配置：
   - **Name**: futures-scanner（或其他名字）
   - **Region**: Singapore（离你最近）
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt && pip install flask flask-cors gunicorn`
   - **Start Command**: `gunicorn --bind 0.0.0.0:$PORT --timeout 300 cloud_server:app`
4. 点 **Create Web Service**

## 方案三：查看部署失败原因

如果部署失败，点击 `(deploy failed)` 旁边的按钮查看具体日志，根据日志修复问题。

## 注意

- 免费套餐 512MB 内存，首次启动需要下载数据，可能比较慢
- 每次重新部署（代码更新）会清空容器，output 目录数据丢失，但运行模型后会自动重新生成
