FROM python:3.13-slim

WORKDIR /app
COPY agent-backend/requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
COPY agent-backend /app/agent-backend
COPY AI穿搭导购_商品更新源.xlsx /app/AI穿搭导购_商品更新源.xlsx

ENV PYTHONUNBUFFERED=1
ENV DATABASE_PATH=/data/agent.db
ENV CATALOG_XLSX_PATH=/app/AI穿搭导购_商品更新源.xlsx
EXPOSE 8080
CMD ["sh", "-c", "uvicorn app.main:app --app-dir /app/agent-backend --host 0.0.0.0 --port ${PORT:-8080}"]
