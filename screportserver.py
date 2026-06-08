import re
import os
from contextvars import ContextVar
from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse
from pydantic import BaseModel
from google.cloud import bigquery
from mcp.server.fastmcp import FastMCP
from fastapi.middleware.cors import CORSMiddleware

# --------------------------------------------------------
# Config & Client Setup
# --------------------------------------------------------
PROJECT_ID = "sc-ai-uat"
DATASET_ID = "SCReport"
AUTH_TABLE = "AuthenByMenu"
# ใส่ Email จริงของคุณเป็น Default เพื่อกันความผิดพลาด
DEFAULT_USER_EMAIL = "rattanachote@scasset.com" 

PUBLIC_BASE_URL = os.environ.get(
    "PUBLIC_BASE_URL",
    "https://sc-report-866803019306.asia-southeast3.run.app",
)

REPORT_DIR = Path("/tmp/reports").resolve()
REPORT_DIR.mkdir(parents=True, exist_ok=True)

bq_client = bigquery.Client(project=PROJECT_ID)

# ใช้ ContextVar เพื่อส่ง Email ข้ามไปยัง Function ต่างๆ ได้ปลอดภัย
CURRENT_USER_EMAIL: ContextVar[str] = ContextVar("current_user_email", default=DEFAULT_USER_EMAIL)

# --------------------------------------------------------
# FastAPI App
# --------------------------------------------------------
app = FastAPI(title="SC Report MCP Server (Updated)")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# --------------------------------------------------------
# Helpers
# --------------------------------------------------------
def clean_string(s: str) -> str:
    """ลบ whitespace และอักขระพิเศษที่อาจปนมา"""
    return s.strip() if s else ""

def is_valid_email(email: str) -> bool:
    """ตรวจสอบว่าไม่ใช่ template string เช่น ${user.email} และเป็น email format"""
    if not email or "${" in email:
        return False
    return bool(re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", email))

# Mapping ชื่อตาราง กับชื่อรายงานใน DB (รองรับทุก Format ที่อาจเป็นไปได้)
TABLE_TO_REPORT_NAMES = {
    "vrptexpension": [
        "รายงานตรวจสอบการจ่ายเงิน(ต้นทุน)",
        "รายงานตรวจสอบการจ่ายเงิน (ต้นทุน)",
        "ทะเบียนคุมการเบิกจ่าย (ต้นทุน)",
        "ทะเบียนคุมการเบิกจ่าย(ต้นทุน)",
    ],
    "vrptexpensionexpmodule": [
        "รายงานตรวจสอบการจ่ายเงิน (ค่าใช้จ่าย)",
        "รายงานตรวจสอบการจ่ายเงิน(ค่าใช้จ่าย)",
        "ทะเบียนคุมการเบิกจ่าย (ค่าใช้จ่าย)",
        "ทะเบียนคุมการเบิกจ่าย(ค่าใช้จ่าย)",
    ],
}

def check_user_permission(email: str, table_id: str) -> Optional[str]:
    """ตรวจสอบสิทธิ์ใน BigQuery แบบละเอียด"""
    tid_clean = table_id.lower().strip()
    report_names = TABLE_TO_REPORT_NAMES.get(tid_clean)

    if not report_names:
        print(f"🚨 [Auth] Table '{tid_clean}' not mapped in code.")
        return None

    email_clean = clean_string(email).lower()
    print(f"🔍 [Auth] Checking: {email_clean} for {tid_clean}")

    try:
        # ใช้ TRIM() ทั้งสองฝั่งเพื่อความปลอดภัยสูงสุด
        placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
        query = f"""
            SELECT TRIM(ReportName) AS ReportName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE LOWER(TRIM(Email)) = @email
            AND TRIM(ReportName) IN ({placeholders})
            LIMIT 1
        """
        params = [bigquery.ScalarQueryParameter("email", "STRING", email_clean)]
        for i, name in enumerate(report_names):
            params.append(bigquery.ScalarQueryParameter(f"name{i}", "STRING", name))

        job_config = bigquery.QueryJobConfig(query_parameters=params)
        df = bq_client.query(query, job_config=job_config).to_dataframe()

        if not df.empty:
            found_name = df["ReportName"].iloc[0]
            print(f"✅ [Auth] Permission Granted: {found_name}")
            return found_name

        print(f"❌ [Auth] Permission Denied for {email_clean}")
        return None
    except Exception as e:
        print(f"🚨 [Auth Error] {str(e)}")
        return None

# --------------------------------------------------------
# MCP & Main Tool
# --------------------------------------------------------
mcp = FastMCP("sc-report-server", stateless_http=True)

@mcp.tool()
def generate_excel_report(table_id: str, user_email: str) -> str:
    """ดึงรายงานและสร้างลิงก์ Excel (ตรวจสอบสิทธิ์จาก AuthenByMenu)"""
    
    # 1. Resolve Email: ลำดับคือ user_email ที่ส่งมา > Context > Default
    resolved_email = clean_string(user_email) if is_valid_email(user_email) else CURRENT_USER_EMAIL.get()
    tid = table_id.lower().strip().split(".")[-1].replace("`", "")

    # 2. Check Permission
    report_display_name = check_user_permission(resolved_email, tid)
    if not report_display_name:
        return (f"🙏 ขออภัยครับคุณ {resolved_email.split('@')[0]} ระบบตรวจสอบพบว่าคุณยังไม่มีสิทธิ์เข้าถึงรายงาน '{tid}'\n"
                "กรุณาตรวจสอบสิทธิ์ในระบบ SC System หรือติดต่อ Admin ครับ")

    # 3. Fetch Data & Create Excel
    try:
        query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{tid}` LIMIT 5000"
        df = bq_client.query(query).to_dataframe()
        
        if df.empty:
            return f"❌ ไม่พบข้อมูลในรายงาน {report_display_name}"

        # ลบ Timezone ออกเพื่อให้ Excel รองรับ
        for col in df.select_dtypes(include=["datetimetz"]).columns:
            df[col] = df[col].dt.tz_localize(None)

        file_name = f"Report_{tid}.xlsx"
        file_path = REPORT_DIR / file_name
        df.to_excel(file_path, index=False)

        download_url = f"{PUBLIC_BASE_URL.rstrip('/')}/download/{file_name}"

        return (
            f"✅ **สร้างรายงานสำเร็จ!**\n"
            f"📊 รายงาน: {report_display_name}\n"
            f"📥 [ดาวน์โหลดไฟล์ Excel คลิกที่นี่]({download_url})"
        )
    except Exception as e:
        return f"🚨 เกิดข้อผิดพลาดขณะสร้างไฟล์: {str(e)}"

# --------------------------------------------------------
# FastAPI Middleware & Routes
# --------------------------------------------------------
@app.middleware("http")
async def inject_user_email(request: Request, call_next):
    # ดึง email จาก Header ที่ AI ส่งมา
    email = request.headers.get("x-user-email") or DEFAULT_USER_EMAIL
    token = CURRENT_USER_EMAIL.set(email)
    try:
        return await call_next(request)
    finally:
        CURRENT_USER_EMAIL.reset(token)

@app.get("/download/{file_name}")
def download_report(file_name: str):
    file_path = (REPORT_DIR / file_name).resolve()
    if file_path.exists() and file_path.is_relative_to(REPORT_DIR):
        return FileResponse(path=file_path, filename=file_name)
    return {"error": "File not found"}

# Mount MCP Server
app.mount("", mcp.streamable_http_app())

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
