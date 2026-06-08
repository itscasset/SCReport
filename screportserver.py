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
    print(f"📨 [Request] {request.method} {request.url.path} | x-user-email: {email}")
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
    return raw.strip().strip("`").split(".")[-1].strip()

def validate_table_id(table_id: str) -> bool:
    return bool(re.match(r"^[A-Za-z0-9_]+$", table_id))

def is_valid_email(email: str) -> bool:
    return bool(re.match(r"^[^@${}]+@[^@${}]+\.[^@${}]+$", email))

def check_user_has_access(email: str, table_id: str) -> bool:
    """
    ตรวจสอบว่า email มีสิทธิ์ดูตารางนี้ไหม
    โดยเช็คว่ามี row ใดๆ ใน AuthenByMenu ที่ตรงกับ email และมีคำที่เกี่ยวข้องกับตารางนั้น
    """
    # Keyword mapping — ถ้า email มี ReportName ที่มีคำเหล่านี้ = มีสิทธิ์
    TABLE_KEYWORDS: dict[str, list[str]] = {
        "vrptexpension": ["ต้นทุน"],
        "vrptexpensionexpmodule": ["ค่าใช้จ่าย"],
    }

    clean = table_id.lower()
    keywords = TABLE_KEYWORDS.get(clean)

    if not keywords:
        print(f"🚨 [Auth] ไม่พบ keyword สำหรับตาราง: {clean}")
        return False

    try:
        conditions = " OR ".join([f"TRIM(ReportName) LIKE @kw{i}" for i in range(len(keywords))])
        query = f"""
            SELECT 1
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE LOWER(TRIM(Email)) = LOWER(TRIM(@email))
            AND ({conditions})
            LIMIT 1
        """
        params = [bigquery.ScalarQueryParameter("email", "STRING", email.strip())]
        for i, kw in enumerate(keywords):
            params.append(bigquery.ScalarQueryParameter(f"kw{i}", "STRING", f"%{kw}%"))

        job_config = bigquery.QueryJobConfig(query_parameters=params)
        df = bq_client.query(query, job_config=job_config).to_dataframe()

        has_access = not df.empty
        print(f"{'✅' if has_access else '❌'} [Auth] {email} → {clean}: {'มีสิทธิ์' if has_access else 'ไม่มีสิทธิ์'}")
        return has_access

    except Exception as e:
        print(f"🚨 [Auth Error] {str(e)}")
        return False


def fetch_and_generate_excel(table_id: str) -> Optional[pd.DataFrame]:
    """ดึงข้อมูลจาก BigQuery และบันทึกเป็น Excel"""
    query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{table_id}` LIMIT 2000"
    df = bq_client.query(query).to_dataframe()
    if df.empty:
        return None
    # ✅ Strip timezone จาก datetime columns (Excel ไม่รองรับ timezone-aware)
    for col in df.select_dtypes(include=["datetimetz"]).columns:
        df[col] = df[col].dt.tz_localize(None)
    file_path = REPORT_DIR / f"Report_{table_id}.xlsx"
    df.to_excel(file_path, index=False)
    return df

# --------------------------------------------------------
# ✅ MCP Server (Stateless Mode)
# --------------------------------------------------------
mcp = FastMCP(
    "sc-report-server",
    stateless_http=True,
    json_response=True,
)

@mcp.tool(
    name="generate_excel_report",
    description=(
        "ดึงข้อมูลรายงานจาก BigQuery และสร้างไฟล์ Excel ให้ดาวน์โหลด "
        "table_id ที่รองรับ: vrptexpension, vrptexpensionexpmodule "
        "user_email ต้องเป็น email จริงเท่านั้น เช่น name@scasset.com"
    ),
)
def mcp_generate_excel_report(table_id: str, user_email: str) -> str:
    # 1. Clean table_id
    tid = clean_table_id(table_id)
    print(f"📋 [MCP] table_id='{tid}' user_email='{user_email}'")

    if not validate_table_id(tid):
        return f"❌ table_id '{tid}' ไม่ถูกต้อง กรุณาระบุเช่น vrptexpension"

    # 2. Resolve email
    email = user_email if is_valid_email(user_email) else None
    email = email or CURRENT_REQUEST_USER_EMAIL.get() or DEFAULT_USER_EMAIL
    print(f"👤 [MCP] resolved email='{email}'")

    # 3. ตรวจสอบสิทธิ์ดู (ถ้ามีสิทธิ์ดู = สร้างได้เลย)
    if not check_user_has_access(email, tid):
        return (
            f"🙏 ขออภัยในความไม่สะดวกครับคุณ {email.split('@')[0]} "
            "เนื่องจากระบบตรวจสอบพบว่าคุณยังไม่มีสิทธิ์เข้าถึงรายงานตัวนี้\n\n"
            "💡 สามารถตรวจสอบสิทธิ์เพิ่มเติมได้ที่ระบบ **sc system** ครับ"
        )

    # 4. สร้าง Excel ได้เลย
    try:
        df = fetch_and_generate_excel(tid)
        if df is None:
            return f"❌ ไม่พบข้อมูลในตาราง {tid}"

        file_name = f"Report_{tid}.xlsx"
        download_url = f"{PUBLIC_BASE_URL.rstrip('/')}/download/{file_name}"

        return (
            "✅ ตรวจสอบสิทธิ์เรียบร้อยแล้วครับ\n"
            f"📊 จัดเตรียมรายงานจำนวนทั้งหมด {len(df)} แถวให้คุณเรียบร้อยแล้ว\n\n"
            "📥 **คลิกลิงก์เพื่อดาวน์โหลดไฟล์:**\n"
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
    user_email = (
        http_request.headers.get("x-user-email")
        or request.user_email
        or DEFAULT_USER_EMAIL
    )

    if not validate_table_id(tid):
        return {"success": False, "message": f"❌ table_id '{tid}' ไม่ถูกต้อง"}

    if not check_user_has_access(user_email, tid):
        return {"success": False, "message": "🙏 ขออภัย คุณยังไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    try:
        df = fetch_and_generate_excel(tid)
        if df is None:
            return {"success": False, "message": "❌ ไม่พบข้อมูล"}

        file_name = f"Report_{tid}.xlsx"
        download_url = f"{PUBLIC_BASE_URL.rstrip('/')}/download/{file_name}"

        return {
            "success": True,
            "message": f"✅ จัดเตรียมรายงานสำเร็จ\n🔗 {download_url}",
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
# ✅ Mount ที่ "/mcp" เพื่อเลี่ยงไม่ให้ชนกับ Root API และ Health Check ของ GCP
# --------------------------------------------------------
app.mount("/mcp", mcp.streamable_http_app())

# --------------------------------------------------------
# Entry Point
# --------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
