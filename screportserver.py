import os
import re
import uuid
import logging
import time  # 👈 เพิ่มเข้ามาเพื่อใช้วัดอายุไฟล์
import asyncio
from datetime import datetime, timedelta
from contextlib import asynccontextmanager
from contextvars import ContextVar
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from fastapi import FastAPI, Request, HTTPException, BackgroundTasks  # 👈 เพิ่ม BackgroundTasks สำหรับรันงานเบื้องหลัง
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel
from google.cloud import bigquery
from mcp.server.fastmcp import FastMCP
from fastapi.middleware.cors import CORSMiddleware

# --------------------------------------------------------
# Configuration
# --------------------------------------------------------
PROJECT_ID = "sc-ai-uat"
DATASET_ID = "SCReport"
AUTH_TABLE = "AuthenByMenu"
DEFAULT_USERNAME = ""
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "https://sc-report-866803019306.asia-southeast3.run.app")

REPORT_DIR = Path("/tmp/reports").resolve()
REPORT_DIR.mkdir(parents=True, exist_ok=True)

bq_client = bigquery.Client(project=PROJECT_ID)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("SCReportSecurity")

CURRENT_REQUEST_USERNAME: ContextVar[Optional[str]] = ContextVar("current_request_username", default=None)
CURRENT_REQUEST_BASE_URL: ContextVar[Optional[str]] = ContextVar("current_request_base_url", default=None)

# --------------------------------------------------------
# Models (🛡️ รองรับพารามิเตอร์ทุกรูปแบบของ Agent เพื่อป้องกัน 422 Error)
# --------------------------------------------------------

class FilterCondition(BaseModel):
    """
    🔒 Structured filter — แทนที่การรับ raw SQL condition string
    ทุก field ถูก validate ก่อนนำไปสร้าง Parameterized Query เท่านั้น
    """
    column: str    # ต้องอยู่ใน whitelist ของตารางนั้น
    operator: str  # ต้องอยู่ใน ALLOWED_OPERATORS
    value: str     # ถูกส่งเป็น @param เสมอ ไม่แทรกตรง query

class GenerateReportRequest(BaseModel):
    table_id: str
    limit: int = 2000
    filter_column: Optional[str] = None   # 💾 legacy single-filter
    filter_value: Optional[str] = None    # 💾 legacy single-filter value
    filters: Optional[list[FilterCondition]] = None  # ✅ structured multi-filter (แนะนำ)
    condition: Optional[str] = None       # ⚠️ deprecated — รองรับของเดิม แต่จะถูก ignore
    username: Optional[str] = None        # 💾 รองรับการส่ง Username จาก Agent
    user_email: Optional[str] = None      # 💾 รองรับการส่ง Email จาก Agent

    class Config:
        extra = "allow"  # 🔒 ซับแรงกระแทก ยอมรับทุกฟิลด์แปลกๆ ที่ Agent อาจจะส่งเพิ่มมา

# --------------------------------------------------------
# Mapping ตารางรายงานและสิทธิ์
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
# 🛡️ SQL Security — Column Whitelist & Operator Whitelist
# --------------------------------------------------------
# เพิ่มคอลัมน์จริงของแต่ละตารางที่นี่เท่านั้น — ห้ามรับชื่อคอลัมน์จาก user โดยตรงเด็ดขาด
ALLOWED_COLUMNS_BY_TABLE: dict[str, set[str]] = {
    "vrptexpension": {
        "CompanyCode", "FiscalYear", "VendorCode", "DocumentDate",
        "PostingDate", "DocumentNo", "Amount", "Currency", "CostCenter",
    },
    "vrptexpensionexpmodule": {
        "CompanyCode", "FiscalYear", "CostCenter", "GLAccount",
        "DocumentDate", "PostingDate", "DocumentNo", "Amount", "Currency",
    },
}

# อนุญาตเฉพาะ Operator ที่ปลอดภัย — ห้ามรับ operator จาก user โดยตรง
ALLOWED_OPERATORS: set[str] = {"=", ">", "<", ">=", "<=", "LIKE", "!="}

# --------------------------------------------------------
# Application & Lifespan Setup
# --------------------------------------------------------
mcp = FastMCP("sc-report-server", stateless_http=True, json_response=True, host="0.0.0.0")
mcp_asgi_app = mcp.streamable_http_app()

async def midnight_cleanup_task():
    """
    🕛 ฟังก์ชันสแกนและลบไฟล์ Excel ทั้งหมดในโฟลเดอร์ตอนเที่ยงคืนตรง
    """
    while True:
        now = datetime.now()
        # คำนวณเวลาที่เหลือจนกว่าจะถึงเที่ยงคืนของวันถัดไป
        next_midnight = datetime(now.year, now.month, now.day) + timedelta(days=1)
        sleep_seconds = (next_midnight - now).total_seconds()
        
        logger.info(f"⏳ รอการลบไฟล์ตอนเที่ยงคืนในอีก {sleep_seconds:.0f} วินาที")
        await asyncio.sleep(sleep_seconds)
        
        logger.info("🕛 เริ่มต้นการลบไฟล์ Excel ประจำวัน (เที่ยงคืน)")
        try:
            count = 0
            for file_path in REPORT_DIR.glob("Report_*.xlsx"):
                if file_path.is_file():
                    file_path.unlink()
                    count += 1
            logger.info(f"✅ ลบไฟล์รายงานเรียบร้อยแล้วจำนวน {count} ไฟล์")
        except Exception as e:
            logger.error(f"เกิดข้อผิดพลาดในการลบไฟล์ตอนเที่ยงคืน: {e}")

@asynccontextmanager
async def lifespan(app: FastAPI):
    # เริ่มต้นการทำงานเบื้องหลังลบไฟล์ตอนเที่ยงคืน
    cleanup_task = asyncio.create_task(midnight_cleanup_task())
    async with mcp.session_manager.run():
        yield
    cleanup_task.cancel()

app = FastAPI(title="Secure SC Report API", version="1.3.0", lifespan=lifespan)
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
    token_base = CURRENT_REQUEST_BASE_URL.set(str(request.base_url).rstrip("/"))
    try:
        response = await call_next(request)
    finally:
        CURRENT_REQUEST_USERNAME.reset(token_username)
        CURRENT_REQUEST_BASE_URL.reset(token_base)
    return response

# --------------------------------------------------------
# Helpers & Security Core
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
    """
    🛡️ ดึง Username แบบ Hybrid (เช็ค Header ก่อน ถ้าไม่มีให้ใช้จากที่ Agent ส่งมา)
    """
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
    
    # 🔒 [HARD SECURITY] บล็อกการเจาะระบบตรวจสอบสิทธิ์ผ่านช่องทางตาราง AuthenByMenu
    if "AuthenByMenu" in clean:
        logger.warning(f"🚨 Security Blocked: User '{username}' tried to access permission table via check_user_permission")
        return None

    report_names = TABLE_TO_REPORT_NAMES.get(clean)
    if not report_names:
        return None

    try:
        placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
        query = f"""
            SELECT ReportName FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE LOWER(TRIM(UserName)) = LOWER(TRIM(@username))
            AND TRIM(ReportName) IN ({placeholders})
            LIMIT 1
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

def validate_filter_column(table_id: str, column: str) -> bool:
    """
    🛡️ ตรวจสอบว่าชื่อคอลัมน์อยู่ใน Whitelist ของตารางนั้นหรือไม่
    ป้องกัน Schema Probing และการดึงข้อมูลคอลัมน์ที่ไม่อนุญาต
    """
    allowed = ALLOWED_COLUMNS_BY_TABLE.get(table_id.lower(), set())
    return column in allowed


def fetch_and_generate_excel(
    table_id: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
) -> tuple[Optional[int], Optional[str]]:
    """
    🔒 ดึงข้อมูลจาก BigQuery ด้วย Parameterized Query เท่านั้น
    ไม่มีการแทรก user input ลงใน query string โดยตรงอีกต่อไป
    """
    clean = clean_table_id(table_id).lower()

    # 🔒 [HARD SECURITY] บล็อกการดึงข้อมูลจากตารางสิทธิ์เด็ดขาดในระดับล่างสุด
    if "AuthenByMenu" in clean:
        raise PermissionError("Access to the authentication table is strictly prohibited at code level.")

    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(clean, table_id)
    query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{bq_table_name}`"
    params = []
    where_clauses = []

    # ✅ Path 1: Structured filters (FilterCondition list) — 100% Parameterized
    if filters:
        for i, f in enumerate(filters):
            # Whitelist: column ต้องอยู่ใน ALLOWED_COLUMNS_BY_TABLE ของตารางนี้
            if not validate_filter_column(clean, f.column):
                logger.warning(f"🚨 Column Whitelist Blocked: '{f.column}' not allowed for table '{clean}'")
                raise PermissionError(f"คอลัมน์ '{f.column}' ไม่ได้รับอนุญาตสำหรับตารางนี้")
            # Whitelist: operator ต้องอยู่ใน ALLOWED_OPERATORS
            if f.operator not in ALLOWED_OPERATORS:
                logger.warning(f"🚨 Operator Whitelist Blocked: '{f.operator}'")
                raise PermissionError(f"Operator '{f.operator}' ไม่ได้รับอนุญาต — ใช้ได้เฉพาะ: {ALLOWED_OPERATORS}")
            # ✅ column & operator มาจาก whitelist (ปลอดภัยสำหรับ interpolation)
            # ✅ value ถูกส่งเป็น @param เสมอ — ไม่แทรกตรง query string
            param_name = f"filter_{i}"
            where_clauses.append(f"`{f.column}` {f.operator} @{param_name}")
            params.append(bigquery.ScalarQueryParameter(param_name, "STRING", f.value))

    # ✅ Path 2: Legacy single filter_column/filter_value — Column Whitelist + Parameterized
    elif filter_column and filter_value:
        if not validate_filter_column(clean, filter_column):
            logger.warning(f"🚨 Column Whitelist Blocked (legacy): '{filter_column}' for table '{clean}'")
            raise PermissionError(f"คอลัมน์ '{filter_column}' ไม่ได้รับอนุญาตสำหรับตารางนี้")
        where_clauses.append(f"`{filter_column}` = @filter_value")
        params.append(bigquery.ScalarQueryParameter("filter_value", "STRING", filter_value))

    # ⚠️ Path 3: raw condition string — REMOVED (deprecated, ถูก ignore เงียบๆ เพื่อความปลอดภัย)

    if where_clauses:
        query += " WHERE " + " AND ".join(where_clauses)

    safe_limit = min(max(limit, 1), 10000)
    query += f" LIMIT {safe_limit}"

    job_config = bigquery.QueryJobConfig(query_parameters=params)
    job = bq_client.query(query, job_config=job_config)
    result = job.result()
    schema = result.schema
    results = list(result)

    if not results:
        return None, None

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.append([field.name for field in schema])

    for row in results:
        row_values = []
        for field in schema:
            val = row[field.name]
            if hasattr(val, "tzinfo") and val.tzinfo is not None:
                val = val.replace(tzinfo=None)
            row_values.append(val)
        ws.append(row_values)

    secure_file_id = uuid.uuid4().hex
    file_name = f"Report_{table_id}_{secure_file_id}.xlsx"
    file_path = REPORT_DIR / file_name
    wb.save(file_path)
    return len(results), file_name


def cleanup_old_reports(max_age_seconds: int = 1800):
    """
    🧹 สแกนและลบไฟล์ Excel ที่มีอายุเกินเวลาที่กำหนด (Default: 30 นาที)
    """
    try:
        now = time.time()
        for file_path in REPORT_DIR.glob("Report_*.xlsx"):
            if file_path.is_file():
                file_age = now - file_path.stat().st_mtime
                if file_age > max_age_seconds:
                    file_path.unlink()
                    logger.info(f"🗑️ Cleaned up expired report file: {file_path.name}")
    except Exception as e:
        logger.error(f"Error during report cleanup: {e}")

# --------------------------------------------------------
# MCP Tools (⚙️ มี Parameter ครบตามที่ Agent ต้องการ)
# --------------------------------------------------------
@mcp.tool(
    name="generate_excel_report",
    description="สร้างไฟล์ Excel จากรายงานใน BigQuery และส่งลิงก์ดาวน์โหลด สามารถกรองข้อมูลได้ด้วย filters (column, operator, value)",
)
def mcp_generate_excel_report(
    table_id: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,   # 💾 legacy single-filter
    filter_value: Optional[str] = None,    # 💾 legacy single-filter value
    filters: Optional[list[dict]] = None,  # ✅ structured multi-filter [{"column":...,"operator":...,"value":...}]
    condition: Optional[str] = None,       # ⚠️ deprecated — รับไว้เพื่อไม่ให้ Agent พัง แต่จะถูก ignore
    username: Optional[str] = None,        # 💾 รองรับของเดิมที่ Agent ส่งมา
    user_email: Optional[str] = None,      # 💾 รองรับของเดิมที่ Agent ส่งมา
    **kwargs,                              # 🔒 ป้องกันพังหากระบบส่งฟิลด์อื่นๆ เพิ่มเติมมา
) -> str:
    cleanup_old_reports()

    username_to_use = resolve_username(username, user_email)
    if not username_to_use:
        return "❌ ไม่พบข้อมูลยืนยันตัวตนในระบบ"
    if "AuthenByMenu" in table_id.lower():
        return "🔒 [Access Denied] ระบบไม่อนุญาตให้เข้าถึงตารางสิทธิ์ผู้ใช้งานผ่านช่องทางนี้"
    if not validate_table_id(table_id):
        return "❌ table_id ไม่ถูกต้อง"

    report_name = check_user_permission(username_to_use, table_id)
    if not report_name:
        return "🙏 ขออภัยครับ คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"

    # แปลง filters dict -> FilterCondition objects
    parsed_filters = None
    if filters:
        try:
            parsed_filters = [FilterCondition(**f) for f in filters]
        except Exception:
            return "❌ รูปแบบ filters ไม่ถูกต้อง — ต้องการ: [{\"column\": str, \"operator\": str, \"value\": str}]"

    try:
        row_count, file_name = fetch_and_generate_excel(
            table_id, limit,
            filter_column=filter_column,
            filter_value=filter_value,
            filters=parsed_filters,
        )
        if row_count is None:
            return f"❌ ไม่พบข้อมูลในรายงาน {report_name}"
        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"
        return f"📊 จัดเตรียม **รายงาน{report_name}** ({row_count} แถว) เรียบร้อยแล้วครับ\n🔗 {download_url}"
    except PermissionError as pe:
        return f"❌ {str(pe)}"
    except Exception as e:
        logger.error(f"Error: {e}")
        return "🚨 เกิดข้อผิดพลาดภายในระบบ"

# --------------------------------------------------------
# REST API Endpoints
# --------------------------------------------------------
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest, background_tasks: BackgroundTasks):
    # 🧹 นำงานลบไฟล์ไปรันเป็น Background Task (เบื้องหลัง) เพื่อให้ API ตอบกลับไวขึ้น ไม่หน่วงผู้ใช้
    background_tasks.add_task(cleanup_old_reports)

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
        row_count, file_name = fetch_and_generate_excel(
            tid,
            limit=request.limit,
            filter_column=request.filter_column,
            filter_value=request.filter_value,
            filters=request.filters,  # ✅ structured + parameterized — raw condition ถูก ignore
        )
        if row_count is None:
            return {"success": False, "message": "❌ ไม่พบข้อมูล"}

        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"
        return {"success": True, "message": f"✅ จัดเตรียม **รายงาน{report_name}** สำเร็จ\n🔗 {download_url}"}
    except PermissionError as pe:
        return {"success": False, "message": f"🔒 {str(pe)}"}
    except Exception as e:
        logger.error(f"API Error: {e}")
        return {"success": False, "message": "🚨 เกิดข้อผิดพลาดภายในเซิร์ฟเวอร์"}

# --------------------------------------------------------
# Download Endpoint
# --------------------------------------------------------
@app.get("/download/{file_name}")
def download_report(file_name: str, direct: bool = False):
    if not re.match(r"^Report_[A-Za-z0-9_]+_[a-f0-9]{32}\.xlsx$", file_name):
        return {"success": False, "message": "❌ รูปแบบลิงก์ไม่ถูกต้อง"}

    file_path = (REPORT_DIR / file_name).resolve()
    if not file_path.is_relative_to(REPORT_DIR) or not file_path.exists():
        return {"success": False, "message": "❌ ไม่พบไฟล์หรือไม่มีสิทธิ์เข้าถึง"}

    if direct:
        return FileResponse(path=file_path, filename=file_name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    html_content = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Downloading...</title></head><body><script>const downloadUrl = window.location.pathname + "?direct=true";const iframe = document.createElement('iframe');iframe.style.display = 'none';iframe.src = downloadUrl;document.body.appendChild(iframe);setTimeout(() => {{ window.close(); }}, 1000);</script></body></html>"""
    return HTMLResponse(content=html_content)

app.mount("", mcp_asgi_app)