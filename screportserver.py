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
# FastAPI App
# --------------------------------------------------------
app = FastAPI(
    title="SC Report MCP API & Server",
    description="API & MCP Server สำหรับตรวจสอบสิทธิ์และสร้าง Excel Report จาก BigQuery",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --------------------------------------------------------
# Config
# --------------------------------------------------------
PROJECT_ID = "sc-ai-uat"
DATASET_ID = "SCReport"
AUTH_TABLE = "AuthenByMenu"
DEFAULT_USER_EMAIL = "test_user@scasset.com"
PUBLIC_BASE_URL = os.environ.get(
    "PUBLIC_BASE_URL",
    "https://sc-report-866803019306.asia-southeast3.run.app",
)

REPORT_DIR = Path("/tmp/reports").resolve()
REPORT_DIR.mkdir(parents=True, exist_ok=True)

bq_client = bigquery.Client(project=PROJECT_ID)

CURRENT_REQUEST_USER_EMAIL: ContextVar[Optional[str]] = ContextVar(
    "current_request_user_email",
    default=None,
)

# --------------------------------------------------------
# Models
# --------------------------------------------------------
class GenerateReportRequest(BaseModel):
    table_id: str
    user_email: Optional[str] = None

# --------------------------------------------------------
# Basic Routes
# --------------------------------------------------------
@app.get("/")
def read_root():
    return {"message": "Welcome to SC Report API. MCP server is active at /mcp"}

@app.get("/health")
def health():
    return {"status": "ok", "service": "sc-report-api-mcp"}

# --------------------------------------------------------
# Middleware — จับ x-user-email header
# --------------------------------------------------------
@app.middleware("http")
async def capture_user_email_header(request: Request, call_next):
    email = request.headers.get("x-user-email") or DEFAULT_USER_EMAIL
    token = CURRENT_REQUEST_USER_EMAIL.set(email)
    try:
        response = await call_next(request)
    finally:
        CURRENT_REQUEST_USER_EMAIL.reset(token)
    return response

# --------------------------------------------------------
# Helpers
# --------------------------------------------------------
def clean_table_id(raw: str) -> str:
    """ลบ backtick, project/dataset prefix, และ whitespace"""
    return raw.strip().strip("`").split(".")[-1].strip()

def validate_table_id(table_id: str) -> bool:
    return bool(re.match(r"^[A-Za-z0-9_]+$", table_id))

def is_valid_email(email: str) -> bool:
    """กรอง literal template string เช่น ${user.email} ออก"""
    return bool(re.match(r"^[^@${}]+@[^@${}]+\.[^@${}]+$", email))

# Mapping ชื่อตารางไปยังชื่อรายงานเพื่อเช็คสิทธิ์ใน DB
TABLE_TO_REPORT_NAMES: dict[str, list[str]] = {
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
    """
    เช็คสิทธิ์การเข้าถึงตารางจาก AuthenByMenu ครั้งเดียว
    ถ้าเจอสิทธิ์ จะคืนค่าชื่อรายงาน (ReportName) กลับไป ถ้าไม่เจอคืนค่า None
    """
    clean = table_id.lower()
    report_names = TABLE_TO_REPORT_NAMES.get(clean)

    if not report_names:
        print(f"🚨 [Auth] ไม่พบ mapping สำหรับตาราง: {clean}")
        return None

    try:
        placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
        query = f"""
            SELECT TRIM(ReportName) AS ReportName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE LOWER(TRIM(Email)) = LOWER(TRIM(@email))
            AND TRIM(ReportName) IN ({placeholders})
            LIMIT 1
        """
        params = [bigquery.ScalarQueryParameter("email", "STRING", email.strip())]
        for i, name in enumerate(report_names):
            params.append(bigquery.ScalarQueryParameter(f"name{i}", "STRING", name))

        job_config = bigquery.QueryJobConfig(query_parameters=params)
        df = bq_client.query(query, job_config=job_config).to_dataframe()

        if not df.empty:
            return df["ReportName"].iloc[0]

        return None
    except Exception as e:
        print(f"🚨 [Auth Error] {str(e)}")
        return None

def fetch_and_generate_excel(table_id: str) -> Optional[pd.DataFrame]:
    """ดึงข้อมูลจาก BigQuery และบันทึกเป็น Excel"""
    query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{table_id}` LIMIT 2000"
    df = bq_client.query(query).to_dataframe()
    if df.empty:
        return None
    # Strip timezone จาก datetime columns (Excel ไม่รองรับ timezone-aware)
    for col in df.select_dtypes(include=["datetimetz"]).columns:
        df[col] = df[col].dt.tz_localize(None)
    file_path = REPORT_DIR / f"Report_{table_id}.xlsx"
    df.to_excel(file_path, index=False)
    return df

# --------------------------------------------------------
# MCP Server
# --------------------------------------------------------
mcp = FastMCP(
    "sc-report-server",
    stateless_http=True,
    json_response=True,
    host="0.0.0.0",
)

@mcp.tool(
    name="generate_excel_report",
    description=(
        "ดึงข้อมูลรายงานจาก BigQuery และสร้างไฟล์ Excel ให้ดาวน์โหลด "
        "โดยทำการตรวจสอบสิทธิ์การดูข้อมูลก่อนสร้างไฟล์"
    ),
)
def mcp_generate_excel_report(table_id: str, user_email: str) -> str:
    tid = clean_table_id(table_id)

    if not validate_table_id(tid):
        return f"❌ table_id '{tid}' ไม่ถูกต้อง"

    email = user_email if is_valid_email(user_email) else None
    email = email or CURRENT_REQUEST_USER_EMAIL.get() or DEFAULT_USER_EMAIL

    # เช็คสิทธิ์ครั้งเดียว ถ้ามีสิทธิ์ก็ไปสร้าง Excel เลย
    report_name = check_user_permission(email, tid)
    if not report_name:
        return (
            f"🙏 ขออภัยคุณ {email.split('@')[0]}\n"
            "ระบบตรวจสอบพบว่าคุณยังไม่มีสิทธิ์เข้าถึงรายงานนี้ครับ"
        )

    try:
        df = fetch_and_generate_excel(tid)
        if df is None:
            return f"❌ ไม่พบข้อมูลในรายงาน {report_name}"

        file_name = f"Report_{tid}.xlsx"
        download_url = f"{PUBLIC_BASE_URL.rstrip('/')}/download/{file_name}"

        return (
            "✅ ตรวจสอบสิทธิ์ผ่านเรียบร้อย\n"
            f"📊 ระบบเตรียม **รายงาน{report_name}** จำนวน {len(df)} แถวเสร็จสิ้น\n\n"
            "📥 **ดาวน์โหลดไฟล์:**\n"
            f"🔗 {download_url}"
        )
    except Exception as e:
        return f"🚨 เกิดข้อผิดพลาด: {str(e)}"

# --------------------------------------------------------
# Startup / Shutdown
# --------------------------------------------------------
@app.on_event("startup")
async def startup_event():
    app.state.mcp_session_manager_ctx = mcp.session_manager.run()
    await app.state.mcp_session_manager_ctx.__aenter__()

@app.on_event("shutdown")
async def shutdown_event():
    ctx = getattr(app.state, "mcp_session_manager_ctx", None)
    if ctx is not None:
        await ctx.__aexit__(None, None, None)

# --------------------------------------------------------
# REST API
# --------------------------------------------------------
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest, http_request: Request):
    tid = clean_table_id(request.table_id)
    user_email = (
        http_request.headers.get("x-user-email")
        or request.user_email
        or DEFAULT_USER_EMAIL
    )

    if not validate_table_id(tid):
        return {"success": False, "message": f"❌ table_id '{tid}' ไม่ถูกต้อง"}

    # เช็คสิทธิ์
    report_name = check_user_permission(user_email, tid)
    if not report_name:
        return {"success": False, "message": "🙏 ขออภัย คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    # มีสิทธิ์ -> เจเนอเรตไฟล์
    try:
        df = fetch_and_generate_excel(tid)
        if df is None:
            return {"success": False, "message": "❌ ไม่พบข้อมูล"}

        file_name = f"Report_{tid}.xlsx"
        download_url = f"{PUBLIC_BASE_URL.rstrip('/')}/download/{file_name}"

        return {
            "success": True,
            "message": f"✅ จัดเตรียม **รายงาน{report_name}** สำเร็จ\n🔗 {download_url}",
        }
    except Exception as e:
        return {"success": False, "message": f"🚨 เกิดข้อผิดพลาด: {str(e)}"}

# --------------------------------------------------------
# Download Excel
# --------------------------------------------------------
@app.get("/download/{file_name}")
def download_report(file_name: str):
    if not file_name.endswith(".xlsx"):
        return {"success": False, "message": "❌ อนุญาตเฉพาะไฟล์ .xlsx"}

    file_path = (REPORT_DIR / file_name).resolve()

    if not file_path.is_relative_to(REPORT_DIR):
        return {"success": False, "message": "❌ ไม่มีสิทธิ์เข้าถึงไฟล์"}

    if not file_path.exists():
        return {"success": False, "message": "❌ ไม่พบไฟล์"}

    return FileResponse(
        path=file_path,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# --------------------------------------------------------
# Mount MCP
# --------------------------------------------------------
app.mount("", mcp.streamable_http_app())

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
