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

# Mapping table_id → keyword สำหรับ LIKE matching
# ใช้ keyword แทน hardcode ชื่อเต็ม เพื่อรองรับ invisible char / spacing ใน DB
TABLE_TO_KEYWORDS: dict[str, list[str]] = {
    "vrptexpension":          ["ต้นทุน"],
    "vrptexpensionexpmodule":  ["ค่าใช้จ่าย"],
}

def check_user_permission(email: str, table_id: str) -> Optional[str]:
    """ตรวจสอบสิทธิ์ผู้ใช้จากตาราง AuthenByMenu ใน BigQuery
    - ใช้ LIKE กับ ReportName เพื่อรองรับ invisible char / spacing ใน DB
    - ใช้ = กับ Email เพื่อความปลอดภัย (ห้ามใช้ LIKE)
    """
    clean = clean_table_id(table_id).lower()
    keywords = TABLE_TO_KEYWORDS.get(clean)

    if not keywords:
        print(f"🚨 [Auth] ไม่พบ mapping สำหรับตาราง: {clean}")
        return None

    print(f"🔍 [Auth] ตรวจสอบสิทธิ์: {email} → ตาราง: {clean}")
    print(f"🔍 [Auth] keywords: {keywords}")

    try:
        like_clauses = " OR ".join(
            [f"TRIM(ReportName) LIKE @kw{i}" for i in range(len(keywords))]
        )
        query = f"""
            SELECT TRIM(ReportName) AS ReportName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE LOWER(TRIM(Email)) = LOWER(TRIM(@email))
            AND ({like_clauses})
            LIMIT 1
        """
        params = [bigquery.ScalarQueryParameter("email", "STRING", email.strip())]
        for i, kw in enumerate(keywords):
            params.append(bigquery.ScalarQueryParameter(f"kw{i}", "STRING", f"%{kw}%"))

        job_config = bigquery.QueryJobConfig(query_parameters=params)
        df = bq_client.query(query, job_config=job_config).to_dataframe()

        if not df.empty:
            found = df["ReportName"].iloc[0]
            print(f"✅ [Auth] พบสิทธิ์: {found}")
            return found

        print(f"❌ [Auth] ไม่พบสิทธิ์: {email} / {clean}")
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
        "สร้างไฟล์ Excel จากรายงานใน BigQuery และส่งลิงก์ดาวน์โหลด "
        "ใช้เมื่อผู้ใช้ต้องการดาวน์โหลดข้อมูล "
        "table_id ที่รองรับ: vrptexpension, vrptexpensionexpmodule "
        "user_email คือ email ของผู้ใช้ที่ระบุไว้ใน system prompt"
    ),
)
def mcp_generate_excel_report(table_id: str, user_email: str) -> str:
    tid = clean_table_id(table_id)
    print(f"📋 [MCP] table_id raw='{table_id}' → clean='{tid}'")
    print(f"👤 [MCP] user_email arg='{user_email}'")

    if not validate_table_id(tid):
        return f"❌ table_id '{tid}' ไม่ถูกต้อง"

    # ลำดับความสำคัญ email:
    # 1. user_email จาก tool argument (Claude อ่านจาก system prompt)
    # 2. x-user-email header ผ่าน ContextVar
    # 3. DEFAULT_USER_EMAIL
    if is_valid_email(user_email):
        resolved_email = user_email
    else:
        ctx_email = CURRENT_REQUEST_USER_EMAIL.get()
        resolved_email = ctx_email if is_valid_email(ctx_email) else DEFAULT_USER_EMAIL

    print(f"👤 [MCP] resolved_email='{resolved_email}'")

    report_name = check_user_permission(resolved_email, tid)
    if not report_name:
        return (
            f"🙏 ขออภัยในความไม่สะดวกครับคุณ {resolved_email.split('@')[0]} "
            "เนื่องจากระบบตรวจสอบพบว่าคุณยังไม่มีสิทธิ์เข้าถึงรายงานตัวนี้ในขณะนี้\n\n"
            "💡 หากต้องการตรวจสอบหรือดูข้อมูลรายงานเพิ่มเติม สามารถเข้าชมได้ที่ระบบ **sc system** ครับ"
        )

    try:
        df = fetch_and_generate_excel(tid)
        if df is None:
            return f"❌ ไม่พบข้อมูลในรายงาน {report_name}"

        file_name = f"Report_{tid}.xlsx"
        download_url = f"{PUBLIC_BASE_URL.rstrip('/')}/download/{file_name}"

        return (
            "✅ ตรวจสอบสิทธิ์ผ่านระบบ AuthenByMenu เรียบร้อยแล้วครับ\n"
            f"📊 ระบบได้จัดเตรียม **รายงาน{report_name}** จำนวนทั้งหมด {len(df)} แถวให้คุณเรียบร้อยแล้ว\n\n"
            "📥 **คลิกลิงก์เพื่อดาวน์โหลดไฟล์ลงเครื่อง:**\n"
            f"🔗 {download_url}"
        )
    except Exception as e:
        return f"🚨 เกิดข้อผิดพลาด: {str(e)}"

# --------------------------------------------------------
# Mount MCP ที่ root "" ตามที่ Tantawan client เชื่อมต่อ
# --------------------------------------------------------
app.mount("", mcp.streamable_http_app())

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
    tid = clean_table_id(request.table_id)
    keywords = TABLE_TO_KEYWORDS.get(tid.lower())
    result = check_user_permission(user_email, tid)

    return {
        "received_raw_email": raw_email,
        "resolved_email": user_email,
        "is_valid_email": is_valid_email(raw_email),
        "received_table_id": request.table_id,
        "cleaned_table_id": tid.lower(),
        "keywords_used": keywords,
        "permission_result": result,
        "has_permission": result is not None,
    }

# --------------------------------------------------------
# Entry Point
# --------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
PYEOF
echo "done"
Output

done
