import os
import re
from contextlib import asynccontextmanager
from contextvars import ContextVar
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel
from google.cloud import bigquery
from mcp.server.fastmcp import FastMCP
from fastapi.middleware.cors import CORSMiddleware

# --------------------------------------------------------
# Config
# --------------------------------------------------------
PROJECT_ID = "sc-ai-uat"
DATASET_ID = "SCReport"
AUTH_TABLE = "AuthenByMenu"
DEFAULT_USER_EMAIL = ""
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

CURRENT_REQUEST_BASE_URL: ContextVar[Optional[str]] = ContextVar(
    "current_request_base_url",
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
# Mapping
# --------------------------------------------------------
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

TABLE_TO_BQ_TABLE_NAME: dict[str, str] = {
    "vrptexpension": "VRptExpension",
    "vrptexpensionexpmodule": "VRptExpensionExpModule",
}

# --------------------------------------------------------
# MCP Server — ต้อง init ก่อน FastAPI เพื่อใช้ใน lifespan
# --------------------------------------------------------
mcp = FastMCP(
    "sc-report-server",
    stateless_http=True,
    json_response=True,
    host="0.0.0.0",
)

# --------------------------------------------------------
# ASGI sub-app (ต้องสร้างก่อนเพื่อให้ session_manager ถูก init)
# --------------------------------------------------------
mcp_asgi_app = mcp.streamable_http_app()

# --------------------------------------------------------
# FastAPI App — lifespan เรียก session_manager.run()
# --------------------------------------------------------
@asynccontextmanager
async def lifespan(app: FastAPI):
    async with mcp.session_manager.run():
        yield

app = FastAPI(
    title="SC Report MCP API & Server",
    description="API & MCP Server สำหรับตรวจสอบสิทธิ์และสร้าง Excel Report จาก BigQuery",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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
# Middleware — จับ current-user หรือ x-user-email header
# --------------------------------------------------------
@app.middleware("http")
async def capture_user_email_header(request: Request, call_next):
    email = (
        request.headers.get("current_user")
        # or request.headers.get("current_user")
        or request.headers.get("x-user-email")
    )
    print(f"📨 [Middleware] {request.method} {request.url.path} | email={repr(email)}")
    token_email = CURRENT_REQUEST_USER_EMAIL.set(email)

    base_url = str(request.base_url).rstrip("/")
    token_base = CURRENT_REQUEST_BASE_URL.set(base_url)

    try:
        response = await call_next(request)
    finally:
        CURRENT_REQUEST_USER_EMAIL.reset(token_email)
        CURRENT_REQUEST_BASE_URL.reset(token_base)
    return response

# --------------------------------------------------------
# Helpers
# --------------------------------------------------------
def clean_table_id(raw: str) -> str:
    """ลบ backtick, project/dataset prefix, และ whitespace"""
    return raw.strip().strip("`").split(".")[-1].strip()

def validate_table_id(table_id: str) -> bool:
    return bool(re.match(r"^[A-Za-z0-9_]+$", table_id))

def is_valid_user(user: Optional[str]) -> bool:
    """กรอง None, empty, และ literal template string เช่น ${user.email} หรือ {{current_user}} ออก"""
    if not user:
        return False
    user_str = user.strip()
    if not user_str:
        return False
    if "${" in user_str or "}" in user_str or "{{" in user_str:
        return False
    return True

def check_user_permission(user_val: str, table_id: str) -> Optional[str]:
    """ตรวจสอบสิทธิ์ผู้ใช้จากตาราง AuthenByMenu ใน BigQuery"""
    clean = clean_table_id(table_id).lower()
    report_names = TABLE_TO_REPORT_NAMES.get(clean)

    if not report_names:
        print(f"🚨 [Auth] ไม่พบ mapping สำหรับตาราง: {clean}")
        return None

    print(f"🔍 [Auth] ตรวจสอบสิทธิ์: {user_val} → ตาราง: {clean}")

    try:
        placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
        query = f"""
            SELECT ReportName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE (LOWER(TRIM(UserName)) = LOWER(TRIM(@user_val)) OR LOWER(TRIM(Email)) = LOWER(TRIM(@user_val)))
            AND TRIM(ReportName) IN ({placeholders})
            LIMIT 1
        """
        params = [bigquery.ScalarQueryParameter("user_val", "STRING", user_val.strip())]
        for i, name in enumerate(report_names):
            params.append(bigquery.ScalarQueryParameter(f"name{i}", "STRING", name))

        job_config = bigquery.QueryJobConfig(query_parameters=params)
        job = bq_client.query(query, job_config=job_config)
        results = list(job)

        if results:
            found = results[0]["ReportName"]
            print(f"✅ [Auth] พบสิทธิ์: {found}")
            return found

        print(f"❌ [Auth] ไม่พบสิทธิ์: {user_val} / {clean}")
        return None
    except Exception as e:
        print(f"🚨 [Auth Error] {str(e)}")
        raise e

def fetch_and_generate_excel(table_id: str, limit: int = 2000, condition: Optional[str] = None) -> Optional[int]:
    """ดึงข้อมูลจาก BigQuery และบันทึกเป็น Excel (ใช้ openpyxl ตรงๆ ไม่ใช้ pandas และ db-dtypes)"""
    clean = clean_table_id(table_id).lower()
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(clean, table_id)

    where_clause = ""
    if condition:
        clean_condition = condition.replace(";", "")
        where_clause = f" WHERE {clean_condition}"

    query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{bq_table_name}`{where_clause} LIMIT {limit}"

    job = bq_client.query(query)
    result = job.result()
    schema = result.schema
    results = list(result)
    if not results:
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    headers = [field.name for field in schema]
    ws.append(headers)

    for row in results:
        row_values = []
        for field in schema:
            val = row[field.name]
            if hasattr(val, "tzinfo") and val.tzinfo is not None:
                val = val.replace(tzinfo=None)
            row_values.append(val)
        ws.append(row_values)

    file_path = REPORT_DIR / f"Report_{table_id}.xlsx"
    wb.save(file_path)
    return len(results)

# --------------------------------------------------------
# MCP Tools
# --------------------------------------------------------
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
def mcp_generate_excel_report(
    table_id: str,
    user_email: Optional[str] = None,
    limit: int = 2000,
    condition: Optional[str] = None,
) -> str:
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
        row_count = fetch_and_generate_excel(table_id, limit=limit, condition=condition)
        if row_count is None:
            return f"❌ ไม่พบข้อมูลในรายงาน {report_name}"

        file_name = f"Report_{table_id}.xlsx"
        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"

        return (
            "✅ ตรวจสอบสิทธิ์ผ่านระบบ AuthenByMenu เรียบร้อยแล้วครับ\n"
            f"📊 ระบบได้จัดเตรียม **รายงาน{report_name}** จำนวนทั้งหมด {row_count} แถวให้คุณเรียบร้อยแล้ว\n\n"
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
        http_request.headers.get("current-user")
        or http_request.headers.get("current_user")
        or http_request.headers.get("x-user-email")
        or request.user_email
        or ""
    )
    user_email = raw_email if is_valid_user(raw_email) else DEFAULT_USER_EMAIL

    if not validate_table_id(tid):
        return {"success": False, "message": f"❌ table_id '{tid}' ไม่ถูกต้อง"}

    report_name = check_user_permission(user_email, tid)
    if not report_name:
        return {"success": False, "message": "🙏 ขออภัย คุณยังไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    try:
        row_count = fetch_and_generate_excel(tid, limit=request.limit, condition=request.condition)
        if row_count is None:
            return {"success": False, "message": "❌ ไม่พบข้อมูล"}

        file_name = f"Report_{tid}.xlsx"
        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"

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
def download_report(file_name: str, direct: bool = False):
    if not file_name.endswith(".xlsx"):
        return {"success": False, "message": "❌ อนุญาตเฉพาะไฟล์ .xlsx เท่านั้น"}

    file_path = (REPORT_DIR / file_name).resolve()

    if not file_path.is_relative_to(REPORT_DIR):
        return {"success": False, "message": "❌ ไม่มีสิทธิ์เข้าถึงไฟล์นอกไดเรกทอรีที่กำหนด"}

    if not file_path.exists():
        return {"success": False, "message": "❌ ไม่พบไฟล์ที่ต้องการดาวน์โหลด"}

    if direct:
        return FileResponse(
            path=file_path,
            filename=file_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Render a premium landing page to prevent browser/iframe sandbox from blocking direct downloads
    file_size_bytes = os.path.getsize(file_path)
    if file_size_bytes < 1024:
        file_size_str = f"{file_size_bytes} B"
    elif file_size_bytes < 1024 * 1024:
        file_size_str = f"{file_size_bytes / 1024:.1f} KB"
    else:
        file_size_str = f"{file_size_bytes / (1024 * 1024):.1f} MB"

    download_url = f"/download/{file_name}?direct=true"

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Download Report | SC Asset</title>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg-color: #0d0f12;
            --card-bg: rgba(22, 28, 36, 0.75);
            --primary: #cda34f;
            --primary-hover: #b8903e;
            --text-color: #e2e8f0;
            --text-muted: #94a3b8;
            --border-color: rgba(205, 163, 79, 0.2);
            --success: #10b981;
        }}
        
        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}
        
        body {{
            font-family: 'Outfit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: radial-gradient(circle at top, #1e2025 0%, var(--bg-color) 100%);
            color: var(--text-color);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }}

        .glow {{
            position: absolute;
            width: 400px;
            height: 400px;
            background: radial-gradient(circle, rgba(205, 163, 79, 0.08) 0%, rgba(0,0,0,0) 70%);
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            z-index: 0;
            pointer-events: none;
        }}

        .container {{
            position: relative;
            background: var(--card-bg);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid var(--border-color);
            border-radius: 24px;
            padding: 40px 30px;
            width: 100%;
            max-width: 480px;
            text-align: center;
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.4);
            z-index: 1;
            animation: fadeIn 0.8s cubic-bezier(0.16, 1, 0.3, 1);
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(20px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        .logo {{
            font-weight: 800;
            font-size: 24px;
            letter-spacing: 2px;
            color: var(--primary);
            margin-bottom: 30px;
            text-transform: uppercase;
        }}

        .icon-wrapper {{
            position: relative;
            width: 100px;
            height: 100px;
            margin: 0 auto 30px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(205, 163, 79, 0.1);
            border-radius: 50%;
            border: 1px solid rgba(205, 163, 79, 0.3);
        }}

        .icon-wrapper svg {{
            width: 48px;
            height: 48px;
            fill: var(--primary);
            animation: bounce 2s infinite;
        }}

        @keyframes bounce {{
            0%, 100% {{ transform: translateY(0); }}
            50% {{ transform: translateY(-8px); }}
        }}

        .title {{
            font-size: 22px;
            font-weight: 600;
            margin-bottom: 10px;
            color: #ffffff;
        }}

        .subtitle {{
            font-size: 14px;
            color: var(--text-muted);
            margin-bottom: 30px;
        }}

        .file-info {{
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid rgba(255, 255, 255, 0.05);
            border-radius: 16px;
            padding: 16px;
            margin-bottom: 30px;
            text-align: left;
        }}

        .file-info-row {{
            display: flex;
            justify-content: space-between;
            margin-bottom: 8px;
            font-size: 14px;
        }}

        .file-info-row:last-child {{
            margin-bottom: 0;
        }}

        .label {{
            color: var(--text-muted);
        }}

        .value {{
            font-weight: 500;
            color: #ffffff;
            word-break: break-all;
            max-width: 250px;
            text-align: right;
        }}

        .btn-download {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 100%;
            padding: 16px;
            background: var(--primary);
            color: #000000;
            border: none;
            border-radius: 14px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            text-decoration: none;
            transition: all 0.2s ease;
            box-shadow: 0 4px 14px rgba(205, 163, 79, 0.3);
        }}

        .btn-download:hover {{
            background: var(--primary-hover);
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(205, 163, 79, 0.4);
        }}

        .btn-download:active {{
            transform: translateY(0);
        }}

        .auto-start {{
            margin-top: 20px;
            font-size: 13px;
            color: var(--text-muted);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }}

        .spinner {{
            width: 16px;
            height: 16px;
            border: 2px solid rgba(255, 255, 255, 0.1);
            border-top-color: var(--primary);
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }}

        @keyframes spin {{
            to {{ transform: rotate(360deg); }}
        }}
    </style>
</head>
<body>
    <div class="glow"></div>
    <div class="container">
        <div class="logo">SC ASSET</div>
        <div class="icon-wrapper">
            <svg viewBox="0 0 24 24">
                <path d="M5,20H19V18H5M19,9H15V3H9V9H5L12,16L19,9Z"/>
            </svg>
        </div>
        <h1 class="title">Your report is ready</h1>
        <p class="subtitle">Click the button below to download your spreadsheet.</p>
        
        <div class="file-info">
            <div class="file-info-row">
                <span class="label">File Name</span>
                <span class="value">{file_name}</span>
            </div>
            <div class="file-info-row">
                <span class="label">File Size</span>
                <span class="value">{file_size_str}</span>
            </div>
            <div class="file-info-row">
                <span class="label">Format</span>
                <span class="value">Excel Spreadsheet (.xlsx)</span>
            </div>
        </div>

        <a href="{download_url}" id="downloadLink" class="btn-download">
            Download Report
        </a>

        <div class="auto-start" id="autoStart">
            <div class="spinner"></div>
            <span>Downloading will start automatically in a moment...</span>
        </div>
    </div>

    <script>
        // Auto-trigger download after 800ms
        setTimeout(() => {{
            const link = document.getElementById('downloadLink');
            const tempLink = document.createElement('a');
            tempLink.href = link.href;
            tempLink.target = '_blank';
            document.body.appendChild(tempLink);
            tempLink.click();
            document.body.removeChild(tempLink);
            
            setTimeout(() => {{
                const autoStart = document.getElementById('autoStart');
                autoStart.innerHTML = '<span>If the download did not start, please click the button above.</span>';
            }}, 1500);
        }}, 800);
    </script>
</body>
</html>"""
    return HTMLResponse(content=html_content)

# --------------------------------------------------------
# 🔧 DEBUG: ตรวจสอบ email และสิทธิ์โดยตรง (ลบออกเมื่อ deploy จริง)
# --------------------------------------------------------
@app.post("/debug/check_permission")
def debug_check_permission(request: GenerateReportRequest, http_request: Request):
    raw_email = (
        http_request.headers.get("current-user")
        or http_request.headers.get("current_user")
        or http_request.headers.get("x-user-email")
        or request.user_email
        or ""
    )
    user_email = raw_email if is_valid_user(raw_email) else DEFAULT_USER_EMAIL
    clean = clean_table_id(request.table_id)
    report_names = TABLE_TO_REPORT_NAMES.get(clean.lower())

    error_msg = None
    result = None
    try:
        result = check_user_permission(user_email, request.table_id)
    except Exception as e:
        error_msg = str(e)

    return {
        "received_raw_email": raw_email,
        "resolved_email": user_email,
        "is_valid_email": is_valid_user(raw_email),
        "received_table_id": request.table_id,
        "cleaned_table_id": clean.lower(),
        "expected_report_names": report_names,
        "permission_result": result,
        "has_permission": result is not None,
        "error": error_msg,
    }

# --------------------------------------------------------
# Mount MCP ที่ root "" เพื่อให้ /mcp path ทำงานถูกต้อง
# (FastAPI explicit routes เช่น /, /health จะถูก match ก่อน mount เสมอ)
# --------------------------------------------------------
app.mount("", mcp_asgi_app)

# --------------------------------------------------------
# Entry Point
# --------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)