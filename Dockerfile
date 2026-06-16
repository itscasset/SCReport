# 1. ใช้ Python เวอร์ชันเสถียรและน้ำหนักเบา
FROM python:3.11-slim
# ตั้งค่าไม่ให้ Python สร้างไฟล์ .pyc บ่นรบกวนในคอนเทนเนอร์
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
WORKDIR /app
# 2. ติดตั้ง Dependencies (รวม FastAPI, Uvicorn, และ BigQuery คลีนๆ)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
# 3. ก๊อปปี้โค้ดหลักเข้าไปรัน
COPY screportserver.py mapping.py .
# 4. เปิดพอร์ต 8080 มาตรฐานสำหรับ Google Cloud Run
EXPOSE 8080
# 5. สั่งสตาร์ท FastAPI ด้วย Uvicorn โดยใช้ PORT จาก Cloud Run
CMD ["sh", "-c", "uvicorn screportserver:app --host 0.0.0.0 --port ${PORT:-8080}"]