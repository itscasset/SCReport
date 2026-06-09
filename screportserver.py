import os
import re
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
DEFAULT_USER_EMAIL = "rattanachote@scasset.com"
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
    limit: int = 2000
    condition: Optional[str] = None

# --------------------------------------------------------
# Basic Routes
# --------------------------------------------------------
@app.get("/")
def read_root():
    return {"message": "Welcome to SC Report API. MCP server is active at root \"\""}

@app.get("/health")
def health():
    return {"status": "ok", "service": "sc-report-api-mcp"}

# --------------------------------------------------------
# Middleware — จับ x-user-email header
# --------------------------------------------------------
@app.middleware("http")
async def capture_user_email_header(request: Request, call_next):
    email = request.headers.get("x-user-email")
    print(f"📨 [Middleware] {request.method} {request.url.path} | x-user-email={repr(email)}")
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

def is_valid_email(email: Optional[str]) -> bool:
    """กรอง None และ literal template string เช่น ${user.email} ออก"""
    if not email:
        return False
    return bool(re.match(r"^[^@${}]+@[^@${}]+\.[^@${}]+$", email))

# Mapping table_id → ชื่อรายงานที่อาจเก็บใน DB ทั้ง 2 รูปแบบ (มีเว้นวรรค / ไม่มี)
TABLE_TO_REPORT_NAMES: dict[str, list[str]] = {
    "vrptexpension": [
        "รายงานตรวจสอบการจ่ายเงิน (ต้นทุน)",
        "รายงานตรวจสอบการจ่ายเงิน(ต้นทุน)",
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

# Mapping table_id (lowercase) -> ชื่อตารางจริงใน BigQuery (Case-sensitive)
TABLE_TO_BQ_TABLE_NAME: dict[str, str] = {
    "vrptexpension": "VRptExpension",
    "vrptexpensionexpmodule": "VRptExpensionExpModule",
}

def check_user_permission(email: str, table_id: str) -> Optional[str]:
    """ตรวจสอบสิทธิ์ผู้ใช้จากตาราง AuthenByMenu ใน BigQuery
    รองรับชื่อรายงานทั้งแบบมีและไม่มีเว้นวรรคหน้าวงเล็บ
    """
    clean = clean_table_id(table_id).lower()
    report_names = TABLE_TO_REPORT_NAMES.get(clean)

    if not report_names:
        print(f"🚨 [Auth] ไม่พบ mapping สำหรับตาราง: {clean}")
        return None

    print(f"🔍 [Auth] ตรวจสอบสิทธิ์: {email} → ตาราง: {clean}")

    try:
        # ใช้ IN clause รองรับชื่อรายงานได้หลายรูปแบบในครั้งเดียว
        placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
        query = f"""
            SELECT ReportName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE LOWER(TRIM(Email)) = LOWER(TRIM(@email))
            AND TRIM(ReportName) IN ({placeholders})
            LIMIT 1
        """
        params = [bigquery.ScalarQueryParameter("email", "STRING", email.strip())]
        for i, name in enumerate(report_names):
            params.append(bigquery.ScalarQueryParameter(f"name{i}", "STRING", name))

        job_config = bigquery.QueryJobConfig(query_parameters=params)
        job = bq_client.query(query, job_config=job_config)
        results = list(job)

        if results:
            found = results[0]["ReportName"]
            print(f"✅ [Auth] พบสิทธิ์: {found}")
            return found

        print(f"❌ [Auth] ไม่พบสิทธิ์: {email} / {clean}")
        return None
    except Exception as e:
        print(f"🚨 [Auth Error] {str(e)}")
        return None

def fetch_and_generate_excel(table_id: str, limit: int = 2000, condition: Optional[str] = None) -> Optional[pd.DataFrame]:
    """ดึงข้อมูลจาก BigQuery และบันทึกเป็น Excel"""
    # แปลง table_id เป็นชื่อตารางจริงที่มีตัวพิมพ์ใหญ่-เล็กตรงกับ BigQuery
    clean = clean_table_id(table_id).lower()
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(clean, table_id)

    where_clause = ""
    if condition:
        clean_condition = condition.replace(";", "") # ป้องกัน SQL Injection บางส่วน
        where_clause = f" WHERE {clean_condition}"
        
    query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{bq_table_name}`{where_clause} LIMIT {limit}"
    df = bq_client.query(query).to_dataframe()
    if df.empty:
        return None
    for col in df.select_dtypes(include=["datetimetz"]).columns:
        df[col] = df[col].dt.tz_localize(None)
    file_path = REPORT_DIR / f"Report_{table_id}.xlsx"
    df.to_excel(file_path, index=False)
    return df

# --------------------------------------------------------
# MCP Server — Stateless HTTP
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
        "สร้างไฟล์ Excel จากรายงานใน BigQuery และส่งลิงก์ดาวน์โหลด\n"
        "table_id ที่รองรับ:\n"
        "1. 'vrptexpension' -> สำหรับ 'รายงานตรวจสอบการจ่ายเงิน(ต้นทุน)' และ 'ทะเบียนคุมการเบิกจ่าย (ต้นทุน)'\n"
        "2. 'vrptexpensionexpmodule' -> สำหรับ 'รายงานตรวจสอบการจ่ายเงิน (ค่าใช้จ่าย)' และ 'ทะเบียนคุมการเบิกจ่าย (ค่าใช้จ่าย)'\n"
        "user_email คือ email ของผู้ใช้ที่ระบุไว้ใน system prompt โปรดส่ง parameter นี้มาด้วยทุกครั้ง\n"
        "limit: จำนวนแถว (ค่าเริ่มต้น 2000)\n"
        "condition: เงื่อนไข SQL WHERE (ไม่ต้องมีคำว่า WHERE) เช่น \"ProjectName = 'โครงการ A'\""
    ),
)
def mcp_generate_excel_report(table_id: str, user_email: Optional[str] = None, limit: int = 2000, condition: Optional[str] = None) -> str:
    email_to_use = user_email or CURRENT_REQUEST_USER_EMAIL.get() or DEFAULT_USER_EMAIL

    if not validate_table_id(table_id):
        return "❌ table_id ไม่ถูกต้อง"

    report_name = check_user_permission(email_to_use, table_id)
    if not report_name:
        return (
            f"🙏 ขออภัยในความไม่สะดวกครับคุณ {email_to_use.split('@')[0]} "
            "เนื่องจากระบบตรวจสอบพบว่าคุณยังไม่มีสิทธิ์เข้าถึงรายงานตัวนี้ในขณะนี้\n\n"
            "💡 หากต้องการตรวจสอบหรือดูข้อมูลรายงานเพิ่มเติม สามารถเข้าชมได้ที่ระบบ **sc system** ครับ"
        )

    try:
        df = fetch_and_generate_excel(table_id, limit=limit, condition=condition)
        if df is None:
            return f"❌ ไม่พบข้อมูลในรายงาน {report_name}"

        file_name = f"Report_{table_id}.xlsx"
        download_url = (
            f"{PUBLIC_BASE_URL.rstrip('/')}/download/{file_name}"
            if PUBLIC_BASE_URL
            else f"/download/{file_name}"
        )

        return (
            "✅ ตรวจสอบสิทธิ์ผ่านระบบ AuthenByMenu เรียบร้อยแล้วครับ\n"
            f"📊 ระบบได้จัดเตรียม **รายงาน{report_name}** จำนวนทั้งหมด {len(df)} แถวให้คุณเรียบร้อยแล้ว\n\n"
            "📥 **คลิกลิงก์เพื่อดาวน์โหลดไฟล์ลงเครื่อง:**\n"
            f"🔗 {download_url}"
        )
    except Exception as e:
        return f"🚨 เกิดข้อผิดพลาด: {str(e)}"

# --------------------------------------------------------
# REST API
# --------------------------------------------------------
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest, http_request: Request):
    tid = clean_table_id(request.table_id)

    raw_email = (
        http_request.headers.get("x-user-email")
        or request.user_email
        or ""
    )
    user_email = raw_email if is_valid_email(raw_email) else DEFAULT_USER_EMAIL

    if not validate_table_id(tid):
        return {"success": False, "message": f"❌ table_id '{tid}' ไม่ถูกต้อง"}

    report_name = check_user_permission(user_email, tid)
    if not report_name:
        return {"success": False, "message": "🙏 ขออภัย คุณยังไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    try:
        df = fetch_and_generate_excel(tid, limit=request.limit, condition=request.condition)
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
# Download Excel (🛡️ ป้องกัน Path Traversal)
# --------------------------------------------------------
@app.get("/download/{file_name}")
def download_report(file_name: str):
    if not file_name.endswith(".xlsx"):
        return {"success": False, "message": "❌ อนุญาตเฉพาะไฟล์ .xlsx เท่านั้น"}

    file_path = (REPORT_DIR / file_name).resolve()

    if not file_path.is_relative_to(REPORT_DIR):
        return {"success": False, "message": "❌ ไม่มีสิทธิ์เข้าถึงไฟล์นอกไดเรกทอรีที่กำหนด"}

    if not file_path.exists():
        return {"success": False, "message": "❌ ไม่พบไฟล์ที่ต้องการดาวน์โหลด"}

    return FileResponse(
        path=file_path,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# --------------------------------------------------------
# 🔧 DEBUG: ตรวจสอบ email และสิทธิ์โดยตรง (ลบออกเมื่อ deploy จริง)
# --------------------------------------------------------
@app.post("/debug/check_permission")
def debug_check_permission(request: GenerateReportRequest, http_request: Request):
    raw_email = (
        http_request.headers.get("x-user-email")
        or request.user_email
        or ""
    )
    user_email = raw_email if is_valid_email(raw_email) else DEFAULT_USER_EMAIL
    clean = clean_table_id(request.table_id)
    report_names = TABLE_TO_REPORT_NAMES.get(clean.lower())
    result = check_user_permission(user_email, request.table_id)

    return {
        "received_raw_email": raw_email,
        "resolved_email": user_email,
        "is_valid_email": is_valid_email(raw_email),
        "received_table_id": request.table_id,
        "cleaned_table_id": clean.lower(),
        "expected_report_names": report_names,
        "permission_result": result,
        "has_permission": result is not None,
    }

# --------------------------------------------------------
# Mount MCP Apps (ต้องไว้ท้ายสุดเพื่อไม่ให้ intercept routes อื่น)
# --------------------------------------------------------
app.mount("/mcp", mcp.streamable_http_app())
app.mount("", mcp.streamable_http_app())

# --------------------------------------------------------
# Entry Point
# --------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
