FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# 1. ติดตั้งฟอนต์ไทยก่อน (ต้องอยู่ก่อน pip install)
RUN apt-get update && apt-get install -y \
    fonts-thai-tlwg \
    fontconfig \
    && fc-cache -fv \
    && rm -rf /var/lib/apt/lists/*

# 2. ติดตั้ง Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 3. คัดลอกโค้ด
COPY screportserver.py mapping.py graph_mcp.py .

EXPOSE 8080

CMD ["sh", "-c", "uvicorn screportserver:app --host 0.0.0.0 --port ${PORT:-8080}"]