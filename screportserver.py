import os
import re
import uuid
import logging
import io
from datetime import timedelta
from contextlib import asynccontextmanager
from contextvars import ContextVar
from pathlib import Path
from typing import Optional, Tuple

from openpyxl import Workbook
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from google.cloud import bigquery
from google.cloud import storage  # 👈 เพิ่มเข้ามาสำหรับบริหารจัดการ GCS
from mcp.server.fastmcp import FastMCP
from fastapi.middleware.cors import CORSMiddleware

# --------------------------------------------------------
# Configuration
# --------------------------------------------------------
PROJECT_ID = "sc-ai-uat"
DATASET_ID = "SCReport"
AUTH_TABLE = "AuthenByMenu"
DEFAULT_USERNAME = ""

# 🪣 ดึงชื่อสลอตถังเก็บรายงาน GCS จาก Environment Variable (หรือระบุ Default)
GCS_BUCKET_NAME = os.environ.get("GCS_REPORT_BUCKET", "sc-report-files-uat")

bq_client = bigquery.Client(project=PROJECT_ID)
# ☁️ ประกาศใช้งาน Storage Client สำหรับ Cloud Storage
gcs_client = storage.Client(project=PROJECT_ID)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("SCReportSecurity")

CURRENT_REQUEST_USERNAME: ContextVar[Optional[str]] = ContextVar("current_request_username", default=None)

# --------------------------------------------------------
# Models (รองรับทุกพารามิเตอร์ของ Agent เพื่อป้องกัน 422 Error)
# --------------------------------------------------------
class GenerateReportRequest(BaseModel):
    table_id: str
    limit: int = 2000
    filter_column: Optional[str] = None
    filter_value: Optional[str] = None
    condition: Optional[str] = None
    username: Optional[str] = None
    user_email: Optional[str] = None

    class Config:
        extra = "allow" # บล็อกปัญหากรณี Agent ส่งฟิลด์ประหลาดๆ แนบมา

# --------------------------------------------------------
# Mapping ตารางรายงานและสิทธิ์ ###ควรเพิ่มการหาชื่อรายงานในฐานข้อมูล###
# --------------------------------------------------------
TABLE_TO_REPORT_NAMES: dict[str, list[str]] = {
    "vrptexpension": ["รายงานตรวจสอบการจ่ายเงิน (ต้นทุน)", "รายงานตรวจสอบการจ่ายเงิน(ต้นทุน)", "ทะเบียนคุมการเบิกจ่าย (ต้นทุน)", "ทะเบียนคุมการเบิกจ่าย(ต้นทุน)"],
    "vrptexpensionexpmodule": ["รายงานตรวจสอบการจ่ายเงิน (ค่าใช้จ่าย)", "รายงานตรวจสอบการจ่ายเงิน(ค่าใช้จ่าย)", "ทะเบียนคุมการเบิกจ่าย (ค่าใช้จ่าย)", "ทะเบียนคุมการเบิกจ่าย(ค่าใช้จ่าย)"],
}

TABLE_TO_BQ_TABLE_NAME: dict[str, str] = {
    "vrptexpension": "VRptExpension",
    "vrptexpensionexpmodule": "VRptExpensionExpModule",
}

# --------------------------------------------------------
# Application Lifespan Setup
# --------------------------------------------------------
mcp = FastMCP("sc-report-server", stateless_http=True, json_response=True, host="0.0.0.0")
mcp_asgi_app = mcp.streamable_http_app()

@asynccontextmanager
async def lifespan(app: FastAPI):
    async with mcp.session_manager.run():
        yield

app = FastAPI(title="Secure SC Report API", version="2.0.0", lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

@app.middleware("http")
async def capture_username_header(request: Request, call_next):
    username_val = None
    for header_name in ["current-user", "current_user", "x-user-email"]:
        val = request.headers.get(header_name)
        if val:
            username_val = val
            break
    token_username = CURRENT_REQUEST_USERNAME.set(username_val)
    try:
        response = await call_next(request)
    finally:
        CURRENT_REQUEST_USERNAME.reset(token_username)
    return response

# --------------------------------------------------------
# Security Helpers
# --------------------------------------------------------
def clean_table_id(raw: str) -> str:
    return raw.strip().strip("`").split(".")[-1].strip()

def validate_table_id(table_id: str) -> bool:
    return bool(re.match(r"^[A-Za-z0-9_]+$", table_id))

def is_valid_user(user: Optional[str]) -> bool:
    if not user: return False
    user_str = user.strip()
    if not user_str or "${" in user_str or "}" in user_str or "{{" in user_str: return False
    return True

def resolve_username(fallback_user: Optional[str] = None, fallback_email: Optional[str] = None) -> str:
    context_username = CURRENT_REQUEST_USERNAME.get()
    if is_valid_user(context_username):
        return context_username
    if is_valid_user(fallback_user):
        return fallback_user
    if is_valid_user(fallback_email):
        return fallback_email
    return DEFAULT_USERNAME if DEFAULT_USERNAME else ""

def check_user_permission(username: str, table_id: str) -> Optional[str]:
    clean = clean_table_id(table_id).lower()
    if "AuthenByMenu" in clean:
        logger.warning("🚨 Security Blocked Access Attempt to AuthenByMenu table")
        return None

    report_names = TABLE_TO_REPORT_NAMES.get(clean)
    if not report_names:
        return None

    try:
        placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
        query = f"""
            SELECT ReportName FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}` 
            WHERE LOWER(TRIM(UserName)) = LOWER(TRIM(@username)) 
            AND TRIM(ReportName) IN ({placeholders}) LIMIT 1
        """
        params = [bigquery.ScalarQueryParameter("username", "STRING", username.strip())]
        for i, name in enumerate(report_names):
            params.append(bigquery.ScalarQueryParameter(f"name{i}", "STRING", name))

        job_config = bigquery.QueryJobConfig(query_parameters=params)
        job = bq_client.query(query, job_config=job_config)
        results = list(job)
        return results[0]["ReportName"] if results else None
    except Exception as e:
        logger.error(f"Auth verification failed: {e}")
        return None

# --------------------------------------------------------
# 🔥 REFACTORED: ฟังก์ชันดึงข้อมูลแบบเขียนลง GCS โดยตรงผ่าน RAM
# --------------------------------------------------------
def fetch_and_generate_excel_gcs(
    table_id: str, 
    limit: int = 2000, 
    filter_column: Optional[str] = None, 
    filter_value: Optional[str] = None,
    condition: Optional[str] = None
) -> Tuple[Optional[int], Optional[str]]:
    """
    🚀 ฟังก์ชันดึงข้อมูลจาก BigQuery และเขียนเป็นไฟล์ Excel บนสตรีมหน่วยความจำ (RAM)
    จากนั้นอัปโหลดไปยัง Google Cloud Storage และสร้าง Signed URL สำหรับใช้ดาวน์โหลดได้อย่างปลอดภัย
    """
    clean = clean_table_id(table_id).lower()
    if "AuthenByMenu" in clean:
        raise PermissionError("Access to the authentication table is strictly prohibited at code level.")

    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(clean, table_id)
    query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{bq_table_name}`"
    params = []
    
    # ดักกรองเงื่อนไข
    if filter_column and filter_value:
        if re.match(r"^[A-Za-z0-9_]+$", filter_column):
            query += f" WHERE {filter_column} = @filter_value"
            params.append(bigquery.ScalarQueryParameter("filter_value", "STRING", filter_value))
            
    elif condition:
        # SQL Injection Basic Firewall
        forbidden_keywords = [";", "union", "select", "insert", "update", "delete", "drop", "alter", "grant", "exec", "AuthenByMenu"]
        cond_lower = condition.lower()
        if any(keyword in cond_lower for keyword in forbidden_keywords):
            logger.warning(f"🚨 SQL Firewall Blocked Dangerous Condition: {condition}")
            raise PermissionError("ไม่อนุญาตให้ใช้คำสั่ง SQL ที่อาจเป็นอันตรายในเงื่อนไข (SQL Firewall)")
        query += f" WHERE {condition}"

    safe_limit = min(max(limit, 1), 10000) 
    query += f" LIMIT {safe_limit}"

    # ยิงคำสั่งดึงข้อมูลจาก BigQuery
    job_config = bigquery.QueryJobConfig(query_parameters=params)
    job = bq_client.query(query, job_config=job_config)
    result = job.result()
    schema = result.schema
    results = list(result)
    
    if not results: 
        return None, None

    # 📊 สร้างไฟล์ Excel ใน Memory (In-Memory Buffer)
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.append([field.name for field in schema])

    for row in results:
        row_values = []
        for field in schema:
            val = row[field.name]
            # ตัดเขตเวลา (Timezone) ออกเพื่อหลีกเลี่ยงข้อผิดพลาดของ openpyxl
            if hasattr(val, "tzinfo") and val.tzinfo is not None:
                val = val.replace(tzinfo=None)
            row_values.append(val)
        ws.append(row_values)

    # 📥 จัดเก็บไฟล์ลงในกล่องหน่วยความจำ (BytesIO Object) แทนการเซฟลงดิสก์เครื่อง
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)  # เลื่อนตัวชี้ตำแหน่งไฟล์กลับไปจุดเริ่มต้นเพื่อเริ่มอ่านส่งขึ้นคลาวด์

    # เจนเนอเรตชื่อไฟล์ที่ปลอดภัยและเป็นความลับด้วย UUID
    secure_file_id = uuid.uuid4().hex
    file_name = f"Report_{table_id}_{secure_file_id}.xlsx"
    
    try:
        # ☁️ จัดเตรียมปลายทางและอัปโหลดไฟล์สตรีมจาก RAM เข้าสู่ GCS Bucket
        bucket = gcs_client.bucket(GCS_BUCKET_NAME)
        blob = bucket.blob(file_name)
        blob.upload_from_file(
            excel_buffer, 
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        logger.info(f"☁️ Successfully uploaded {file_name} to GCS bucket: {GCS_BUCKET_NAME}")
        
        # 🔗 ออกเอกสารสิทธิ์ลิงก์ดาวน์โหลดปลอดภัย (V4 Signed URL) อายุ 15 นาที
        signed_url = blob.generate_signed_url(
            version="v4", 
            expiration=timedelta(minutes=15), 
            method="GET"
        )
        return len(results), signed_url
    except Exception as e:
        logger.error(f"GCS Operations failed: {e}")
        raise RuntimeError("เกิดข้อผิดพลาดในการบันทึกหรือออกลิงก์รายงานความปลอดภัยบน Cloud Storage")

# --------------------------------------------------------
# MCP Tools (สำหรับให้ Agent เรียกใช้งานผ่าน MCP Protocol)
# --------------------------------------------------------
@mcp.tool(
    name="generate_excel_report",
    description="สร้างไฟล์ Excel จากรายงานใน BigQuery และส่งลิงก์ดาวน์โหลดที่ปลอดภัยจาก GCS",
)
def mcp_generate_excel_report(
    table_id: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    condition: Optional[str] = None,
    username: Optional[str] = None,
    user_email: Optional[str] = None,
    **kwargs,
) -> str:
    username_to_use = resolve_username(username, user_email)
    if not username_to_use: 
        return "❌ ไม่พบข้อมูลยืนยันตัวตนในระบบ"
    if "AuthenByMenu" in table_id.lower(): 
        return "🔒 [Access Denied] ระบบไม่อนุญาตให้เข้าถึงตารางสิทธิ์ผู้ใช้งานผ่านช่องทางนี้"
    if not validate_table_id(table_id): 
        return "❌ table_id ไม่ถูกต้องตามระบบความปลอดภัย"

    report_name = check_user_permission(username_to_use, table_id)
    if not report_name: 
        return f"🙏 ขออภัยครับ คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"

    try:
        row_count, signed_download_url = fetch_and_generate_excel_gcs(
            table_id, limit, filter_column, filter_value, condition
        )
        if row_count is None: 
            return f"❌ ไม่พบข้อมูลในรายงาน {report_name}"
        
        return (
            f"📊 จัดเตรียม **รายงาน{report_name}** ({row_count} แถว) เรียบร้อยแล้วครับ\n"
            f"🔗 ลิงก์ดาวน์โหลดปลอดภัย (ใช้งานได้ 15 นาที):\n{signed_download_url}"
        )
    except Exception as e:
        return f"❌ เกิดข้อผิดพลาดในการทำงาน: {str(e)}"

# --------------------------------------------------------
# REST API Endpoints (ช่องทาง Webhook/HTTP ปกติ)
# --------------------------------------------------------
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest):
    tid = clean_table_id(request.table_id)
    username_val = resolve_username(request.username, request.user_email)

    if not username_val: 
        return {"success": False, "message": "❌ Unauthorized access. ไม่พบข้อมูลผู้ใช้"}
    if "AuthenByMenu" in tid.lower(): 
        return {"success": False, "message": "🔒 Access Denied: ตารางข้อมูลนี้เป็นความลับขั้นสูงของระบบ"}
    if not validate_table_id(tid): 
        return {"success": False, "message": f"❌ table_id '{tid}' ไม่ถูกต้อง"}

    report_name = check_user_permission(username_val, tid)
    if not report_name: 
        return {"success": False, "message": "🙏 ขออภัย คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    try:
        row_count, signed_download_url = fetch_and_generate_excel_gcs(
            tid, 
            limit=request.limit, 
            filter_column=request.filter_column, 
            filter_value=request.filter_value, 
            condition=request.condition
        )
        if row_count is None: 
            return {"success": False, "message": "❌ ไม่พบข้อมูลในระบบตารางรายงาน"}
            
        return {
            "success": True, 
            "message": f"✅ จัดเตรียม รายงาน{report_name} สำเร็จ", 
            "download_url": signed_download_url
        }
    except Exception as e:
        return {"success": False, "message": f"🚨 Error: {str(e)}"}

# 🚨 สังเกต: เราลบแอปเร้าเตอร์ @app.get("/download/{file_name}") ออกไปแล้ว 
# เพราะตอนนี้ Client สามารถยิงตรงหา GCS Gateway ด้วย Signed URL ได้อย่างปลอดภัยสูงสุด โดยไม่ต้องผ่านเซิร์ฟเวอร์เราซ้ำซ้อน

app.mount("", mcp_asgi_app)

if __name__ == "__main__":
    import uvicorn
    # Cloud Run จะส่งค่า PORT มาให้ทาง Environment Variable (ค่าเริ่มต้นคือ 8080)
    port = int(os.environ.get("PORT", 8080))
    logger.info(f"Starting server on port {port}...")
    uvicorn.run(app, host="0.0.0.0", port=port)