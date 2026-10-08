#
# 默认镜像（新前端 SPA 模式）
# =====================================================================
# 本文件构建「新前端」版本：先构建 ant-design-pro（Ant Design Pro + Umi Max）
# 静态产物，再与 Flask 后端一起打包，SPA_ENABLED=true 由 Flask 直接托管前端。
# 如需「旧前端」镜像（无 Node 构建、templates/static 旧界面），使用：
#   docker build -f Dockerfile.legacy -t outlook-email-plus-legacy .
# =====================================================================

FROM node:22-bookworm-slim AS frontend
WORKDIR /frontend

COPY ant-design-pro/ ./
ENV HUSKY=0 CI=true
RUN npm install --no-audit --no-fund --include=dev
ENV NODE_ENV=production
RUN npm run build

FROM python:3.11-slim

WORKDIR /app

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    GUNICORN_WORKERS=1 \
    GUNICORN_THREADS=8 \
    GUNICORN_TIMEOUT=120 \
    SPA_ENABLED=true

COPY requirements.txt .
RUN pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple && \
    pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple && \
    pip install gunicorn -i https://pypi.tuna.tsinghua.edu.cn/simple

COPY . .
COPY --from=frontend /frontend/dist /app/ant-design-pro/dist

RUN mkdir -p /app/data && chmod +x /app/scripts/start-gunicorn.sh && \
    if [ -f /app/.build-info ]; then cp /app/.build-info /app/.build_info_copy; fi

EXPOSE 5000

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s CMD ["python","-c","import urllib.request as u; u.urlopen('http://localhost:5000/healthz', timeout=4).read()"]

CMD ["scripts/start-gunicorn.sh"]
