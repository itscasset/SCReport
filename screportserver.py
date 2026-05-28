import os
import re
from pathlib import Path
from typing import Optional
 
import pandas as pd
from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from google.cloud import bigquery
 
 
# --------------------------------------------------------
# FastAPI App
# --------------------------------------------------------
 
app = FastAPI(
    title="SC Report MCP API",
    description="API สำหรับตรวจสอบสิทธิ์และสร้าง Excel Report จาก BigQuery",
    version="1.0.0",
)
 
 
# --------------------------------------------------------
# Config
# --------------------------------------------------------
 
PROJECT_ID = os.getenv("GCP_PROJECT_ID", "sc-ai-uat")
DATASET_ID = os.getenv("BQ_DATASET_ID", "SCReport")
AUTH_TABLE = os.getenv("BQ_AUTH_TABLE", "AuthenByMenu")
 
REPORT_DIR = Path("/tmp/reports")
REPORT_DIR.mkdir(parents=True, exist_ok=True)
 
 
# --------------------------------------------------------
# Request Model
# --------------------------------------------------------
 
class GenerateReportRequest(BaseModel):
    table_id: str
    user_email: Optional[str] = None
 
 
# --------------------------------------------------------
# Health Check
# --------------------------------------------------------
 
@app.get("/health")
def health():
    return {
        "status": "ok",
        "service": "sc-report-api",
    }
 
 
# --------------------------------------------------------
# Utility: Validate table_id
# กัน table_id แปลก ๆ เช่น `xxx`; DROP TABLE
# --------------------------------------------------------
 
def validate_table_id(table_id: str) -> bool:
    return bool(re.match(r"^[A-Za-z0-9_]+$", table_id))
 
 
# --------------------------------------------------------
# Authen Function
# --------------------------------------------------------
 
def check_user_permission(email: str, table_id: str):
    print(f"🔍 [Authentication] ตรวจสอบสิทธิ์: {email} สำหรับตาราง: {table_id}")
 
    try:
        client = bigquery.Client(project=PROJECT_ID)
 
        # หมายเหตุ:
        # ตอนนี้เช็คจาก Email ตามโค้ดเดิมของหนุ่ย
        # ถ้าใน AuthenByMenu มี field ชื่อตาราง เช่น TableName / TableId
        # แนะนำเพิ่ม WHERE เข้าไปด้วย
        auth_query = f"""
            SELECT ReportName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE Email = @email
            LIMIT 1
        """
 
        job_config = bigquery.QueryJobConfig(
            query_parameters=[
                bigquery.ScalarQueryParameter("email", "STRING", email),
            ]
        )
 
        auth_df = client.query(auth_query, job_config=job_config).to_dataframe()
 
        if not auth_df.empty:
            return auth_df["ReportName"].iloc[0]
 
        return None
 
    except Exception as auth_error:
        print(f"🚨 ไม่สามารถตรวจสอบสิทธิ์กับ BigQuery ได้: {str(auth_error)}")
        return None
 
 
# --------------------------------------------------------
# Generate Excel Report
# --------------------------------------------------------
 
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest):
    table_id = request.table_id
 
    user_email = (
        request.user_email
        or os.getenv("CURRENT_USER_EMAIL")
        or "test_user@scasset.com"
    )
 
    # 1. Validate table_id ก่อน
    if not validate_table_id(table_id):
        return {
            "success": False,
            "message": "❌ table_id ไม่ถูกต้อง อนุญาตเฉพาะตัวอักษร ตัวเลข และ underscore เท่านั้น",
        }
 
    # 2. ตรวจสอบสิทธิ์
    report_name = check_user_permission(user_email, table_id)
 
    if not report_name:
        return {
            "success": False,
            "message": (
                f"🙏 ขออภัยในความไม่สะดวกครับคุณ {user_email.split('@')[0]} "
                f"เนื่องจากระบบตรวจสอบพบว่าคุณยังไม่มีสิทธิ์เข้าถึงรายงานตัวนี้ในขณะนี้\n\n"
                f"💡 หากต้องการตรวจสอบหรือดูข้อมูลรายงานเพิ่มเติม สามารถเข้าชมได้ที่ระบบ **sc system** ครับ"
            ),
        }
 
    try:
        client = bigquery.Client(project=PROJECT_ID)
 
        query = f"""
            SELECT *
            FROM `{PROJECT_ID}.{DATASET_ID}.{table_id}`
            LIMIT 2000
        """
 
        df = client.query(query).to_dataframe()
 
        if df.empty:
            return {
                "success": False,
                "message": f"❌ ไม่พบข้อมูลในรายงาน {report_name} ({table_id})",
            }
 
        file_name = f"Report_{table_id}.xlsx"
        file_path = REPORT_DIR / file_name
 
        df.to_excel(file_path, index=False)
 
        base_url = os.getenv("PUBLIC_BASE_URL", "")
        if base_url:
            download_url = f"{base_url.rstrip()}/download/{file_name}"
        else:
            download_url = f"/download/{file_name}"
 
        return {
            "success": True,
            "report_name": report_name,
            "table_id": table_id,
            "rows": len(df),
            "file_name": file_name,
            "download_url": download_url,
            "message": (
                f"✅ ตรวจสอบสิทธิ์ผ่านระบบ AuthenByMenu เรียบร้อยแล้วครับ\n"
                f"📊 ระบบได้จัดเตรียม **รายงาน{report_name}** จำนวนทั้งหมด {len(df)} แถวให้คุณเรียบร้อยแล้ว\n\n"
                f"📥 **คลิกลิงก์เพื่อดาวน์โหลดไฟล์ลงเครื่อง:**\n"
                f"🔗 {download_url}"
            ),
        }
 
    except Exception as e:
        return {
            "success": False,
            "message": f"🚨 เกิดข้อผิดพลาดระหว่างดึงข้อมูล: {str(e)}",
        }
 
 
# --------------------------------------------------------
# Download Excel
# --------------------------------------------------------
 
@app.get("/download/{file_name}")
def download_report(file_name: str):
    if not file_name.endswith(".xlsx"):
        return {
            "success": False,
            "message": "❌ อนุญาตให้ดาวน์โหลดเฉพาะไฟล์ .xlsx เท่านั้น",
        }
 
    file_path = REPORT_DIR / file_name
 
    if not file_path.exists():
        return {
            "success": False,
            "message": "❌ ไม่พบไฟล์ที่ต้องการดาวน์โหลด",
        }
 
    return FileResponse(
        path=file_path,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )