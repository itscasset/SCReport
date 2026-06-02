import re
from contextvars import ContextVar
from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse
from pydantic import BaseModel
from google.cloud import bigquery

# 📦 MCP FastMCP สำหรับ Streamable HTTP
from mcp.server.fastmcp import FastMCP

# --------------------------------------------------------
# FastAPI App
# --------------------------------------------------------
app = FastAPI(
    title="SC Report MCP API & Server",
    description="API & MCP Server สำหรับตรวจสอบสิทธิ์และสร้าง Excel Report จาก BigQuery",
    version="1.0.0",
)

# 🛠️ 1. ลงทะเบียน MCP Server แบบ Streamable HTTP
mcp = FastMCP(
    "sc-report-server",
    stateless_http=True,
    json_response=True,
)

# --------------------------------------------------------
# Config
# --------------------------------------------------------
PROJECT_ID = "sc-ai-uat"
DATASET_ID = "SCReport"
AUTH_TABLE = "AuthenByMenu"
DEFAULT_USER_EMAIL = "test_user@scasset.com"
PUBLIC_BASE_URL = ""

REPORT_DIR = Path("/tmp/reports")
REPORT_DIR.mkdir(parents=True, exist_ok=True)

CURRENT_REQUEST_USER_EMAIL: ContextVar[Optional[str]] = ContextVar(
    "current_request_user_email",
    default=None,
)

class GenerateReportRequest(BaseModel):
    table_id: str
    user_email: Optional[str] = None

@app.get("/health")
def health():
    return {"status": "ok", "service": "sc-report-api-mcp"}


@app.middleware("http")
async def capture_user_email_header(request: Request, call_next):
    # Header from LibreChat is expected as X-User-Email
    token = CURRENT_REQUEST_USER_EMAIL.set(request.headers.get("x-user-email"))
    try:
        response = await call_next(request)
    finally:
        CURRENT_REQUEST_USER_EMAIL.reset(token)
    return response

def validate_table_id(table_id: str) -> bool:
    return bool(re.match(r"^[A-Za-z0-9_]+$", table_id))

# --------------------------------------------------------
# Authen Function (โครงสร้างเดิมของพี่หนุ่ย แนะนำปรับ WHERE ตามตารางจริง)
# --------------------------------------------------------
def check_user_permission(email: str, table_id: str):
    print(f"🔍 [Authentication] ตรวจสอบสิทธิ์: {email} สำหรับตาราง: {table_id}")
    try:
        client = bigquery.Client(project=PROJECT_ID)
        
        # 🎯 ปรับแต่งเพิ่มเงื่อนไขตรวจเช็คคู่กับ TableName ในตาราง AuthenByMenu เพื่อความปลอดภัย
        auth_query = f"""
            SELECT ReportName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE Email = @email AND TableName = @table_id
            LIMIT 1
        """
        job_config = bigquery.QueryJobConfig(
            query_parameters=[
                bigquery.ScalarQueryParameter("email", "STRING", email),
                bigquery.ScalarQueryParameter("table_id", "STRING", table_id),
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
# 🌐 ส่วนเสริม: MCP Server Interface (Streamable HTTP)
# --------------------------------------------------------

@mcp.tool(
    name="generate_excel_report",
    description="ดึงข้อมูลรายงานจาก BigQuery โดยจะเช็คสิทธิ์ผู้ใช้งานจากตาราง AuthenByMenu อัตโนมัติก่อนสร้างไฟล์ Excel",
)
def mcp_generate_excel_report(table_id: str) -> str:
    user_email = CURRENT_REQUEST_USER_EMAIL.get() or DEFAULT_USER_EMAIL

    if not validate_table_id(table_id):
        return "❌ table_id ไม่ถูกต้อง"

    report_name = check_user_permission(user_email, table_id)
    if not report_name:
        return (
            f"🙏 ขออภัยในความไม่สะดวกครับคุณ {user_email.split('@')[0]} "
            "เนื่องจากระบบตรวจสอบพบว่าคุณยังไม่มีสิทธิ์เข้าถึงรายงานตัวนี้ในขณะนี้\n\n"
            "💡 หากต้องการตรวจสอบหรือดูข้อมูลรายงานเพิ่มเติม สามารถเข้าชมได้ที่ระบบ **sc system** ครับ"
        )

    try:
        client = bigquery.Client(project=PROJECT_ID)
        query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{table_id}` LIMIT 2000"
        df = client.query(query).to_dataframe()

        if df.empty:
            return f"❌ ไม่พบข้อมูลในรายงาน {report_name}"

        file_name = f"Report_{table_id}.xlsx"
        file_path = REPORT_DIR / file_name
        df.to_excel(file_path, index=False)

        base_url = PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip()}/download/{file_name}" if base_url else f"/download/{file_name}"

        return (
            "✅ ตรวจสอบสิทธิ์ผ่านระบบ AuthenByMenu เรียบร้อยแล้วครับ\n"
            f"📊 ระบบได้จัดเตรียม **รายงาน{report_name}** จำนวนทั้งหมด {len(df)} แถวให้คุณเรียบร้อยแล้ว\n\n"
            "📥 **คลิกลิงก์เพื่อดาวน์โหลดไฟล์ลงเครื่อง:**\n"
            f"🔗 {download_url}"
        )
    except Exception as e:
        return f"🚨 เกิดข้อผิดพลาด: {str(e)}"

# Mount MCP endpoint เป็น Streamable HTTP ที่ /mcp
app.mount("/mcp", mcp.streamable_http_app())


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
# REST API (คงไว้เผื่อฝั่งอื่นจะยิงเรียกใช้แบบเดิม)
# --------------------------------------------------------
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest, http_request: Request):
    # ... (Logic ตัวเดิมที่คุณส่งมาทั้งหมด ยกมาทำงานได้ตามปกติ 100% เลยครับ) ...
    table_id = request.table_id
    user_email = request.user_email or http_request.headers.get("x-user-email") or DEFAULT_USER_EMAIL
    if not validate_table_id(table_id): return {"success": False, "message": "❌ table_id ไม่ถูกต้อง"}
    report_name = check_user_permission(user_email, table_id)
    if not report_name: return {"success": False, "message": f"🙏 ขออภัย... ดูข้อมูลเพิ่มเติมที่ sc system"}
    try:
        client = bigquery.Client(project=PROJECT_ID)
        df = client.query(f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{table_id}` LIMIT 2000").to_dataframe()
        if df.empty: return {"success": False, "message": f"❌ ไม่พบข้อมูล"}
        file_name = f"Report_{table_id}.xlsx"
        df.to_excel(REPORT_DIR / file_name, index=False)
        base_url = PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip()}/download/{file_name}" if base_url else f"/download/{file_name}"
        return {"success": True, "message": f"✅ จัดเตรียม **รายงาน{report_name}** สำเร็จ\n🔗 {download_url}"}
    except Exception as e: return {"success": False, "message": f"🚨 เกิดข้อผิดพลาด: {str(e)}"}

# --------------------------------------------------------
# Download Excel
# --------------------------------------------------------
@app.get("/download/{file_name}")
def download_report(file_name: str):
    if not file_name.endswith(".xlsx"):
        return {"success": False, "message": "❌ อนุญาตเฉพาะไฟล์ .xlsx เท่านั้น"}
    file_path = REPORT_DIR / file_name
    if not file_path.exists():
        return {"success": False, "message": "❌ ไม่พบไฟล์ที่ต้องการดาวน์โหลด"}
    return FileResponse(path=file_path, filename=file_name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    import uvicorn
    # รันบนเครื่องพอร์ต 8080 เป็นหลัก
    uvicorn.run(app, host="0.0.0.0", port=8080)