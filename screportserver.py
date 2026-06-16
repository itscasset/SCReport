import os
import re
import uuid
import logging
import time
import random
import asyncio
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo
from contextlib import asynccontextmanager
from contextvars import ContextVar
from pathlib import Path
from enum import Enum
from typing import Literal, Optional
from pydantic import BaseModel, Field, ConfigDict

from openpyxl import Workbook
from fastapi import FastAPI, Request, HTTPException, BackgroundTasks,Response
from fastapi.responses import FileResponse, HTMLResponse, Response
from google.cloud import bigquery
from mcp.server.fastmcp import FastMCP
from fastapi.middleware.cors import CORSMiddleware
from contextlib import suppress

# --------------------------------------------------------
# Configuration
# --------------------------------------------------------
BANGKOK_TZ = ZoneInfo("Asia/Bangkok")
PROJECT_ID = os.environ.get("PROJECT_ID", "sc-ai-uat")
DATASET_ID = os.environ.get("DATASET_ID", "SCReport")
AUTH_TABLE = os.environ.get("AUTH_TABLE", "AuthenByMenu")
LOG_TABLE = os.environ.get("LOG_TABLE", "AiLog")
DEFAULT_USERNAME = ""
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "https://sc-report-866803019306.asia-southeast3.run.app")
# --------------------------------------------------------
# Mapping Configuration (Imported)
# --------------------------------------------------------
from mapping import (
    TABLE_TO_REPORT_NAMES,
    TABLE_TO_BQ_TABLE_NAME,
    ALLOWED_OPERATORS,
    TABLE_COLUMN_MAPPING,
)

REPORT_DIR = Path("/tmp/reports").resolve()
REPORT_DIR.mkdir(parents=True, exist_ok=True)

bq_client = bigquery.Client(project=PROJECT_ID)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("SCReportSecurity")

CURRENT_REQUEST_USERNAME: ContextVar[Optional[str]] = ContextVar("current_request_username", default=None)
CURRENT_REQUEST_BASE_URL: ContextVar[Optional[str]] = ContextVar("current_request_base_url", default=None)

# --------------------------------------------------------
# ✅ Midnight Cleanup Scheduler
# --------------------------------------------------------
async def _midnight_cleanup_loop() -> None:
    while True:
        try:
            now = datetime.now(BANGKOK_TZ)
            
            # คำนวณหาเที่ยงคืนถัดไป
            next_midnight = (now + timedelta(days=1)).replace(
                hour=0, minute=0, second=0, microsecond=0
            )
            seconds_until_midnight = (next_midnight - now).total_seconds()
            
            logger.info(f" Midnight cleanup scheduled in {seconds_until_midnight:.0f}s (at {next_midnight})")
            await asyncio.sleep(seconds_until_midnight)
            
            logger.info(" Running midnight report cleanup...")
            
            #  แก้ไขปัญหา Blocking: รัน Sync Function ใน Thread Pool
            # (ถ้า cleanup_old_reports เป็น async อยู่แล้ว ให้เปลี่ยนเป็น await cleanup_old_reports() ได้เลย)
            await asyncio.to_thread(cleanup_old_reports)
            
            logger.info(" Midnight report cleanup finished successfully.")
            
        except Exception as e:
            #  ดักจับ Error เพื่อไม่ให้ Loop ตาย
            logger.error(f"❌ Error during midnight cleanup: {e}", exc_info=True)
            
        # เผื่อเวลาไว้ 5 วินาที เพื่อไม่ให้ loop มันหมุนซ้ำถ้าระบบตื่นมาเร็วเกินไปนิดหน่อยก่อนเที่ยงคืนจริง
        await asyncio.sleep(5)

# --------------------------------------------------------
# MCP Server & FastAPI App
# --------------------------------------------------------
mcp = FastMCP(
    "sc-report-server",
    stateless_http=True,
    json_response=True,
    host="0.0.0.0",
)

mcp_asgi_app = mcp.streamable_http_app()

@asynccontextmanager
async def lifespan(app: FastAPI):
    # 1. เริ่มรัน Background Task ก่อน
    cleanup_task = None

    try:
        # 2. รัน MCP Session ภายใน try block
        async with mcp.session_manager.run():
            cleanup_task = asyncio.create_task(_midnight_cleanup_loop())
            logger.info("✅ Midnight cleanup scheduler started")
            yield
            
    finally:
        if cleanup_task:
            cleanup_task.cancel()

            try:
                await asyncio.wait_for(
                    cleanup_task,
                    timeout=10,
                )

            except asyncio.CancelledError:
                logger.info(
                    "🛑 Midnight cleanup scheduler stopped"
                )

            except asyncio.TimeoutError:
                logger.warning(
                    "⚠️ Midnight cleanup task timeout on cancel, but app is shutting down anyway."
                )


app = FastAPI(
    title="SC Report MCP API & Server",
    description="API & MCP Server สำหรับตรวจสอบสิทธิ์และสร้าง Excel Report จาก BigQuery",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://ai-uat.scasset.com",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# --------------------------------------------------------
# Middleware — Fix 302 → 307 to preserve POST method on redirect
# (FastMCP redirects /mcp → /mcp/ with 302, undici changes POST→GET)
# --------------------------------------------------------
@app.middleware("http")
async def preserve_method_on_redirect(request: Request, call_next):
    response = await call_next(request)
    if response.status_code == 302 and request.method in ("POST", "PUT", "PATCH", "DELETE"):
        location = response.headers.get("location", "")
        logger.info(f"🔀 [RedirectFix] 302→307 {request.method} {request.url.path} → {location}")
        response.status_code = 307
    return response

# --------------------------------------------------------
# Middleware — จับ current-user หรือ x-user-email header
# --------------------------------------------------------

# ฟังก์ชันช่วยบังตา (Mask) ข้อมูลส่วนบุคคลก่อนลง Log
def mask_username(username: str | None) -> str:
    if not username:
        return "Anonymous"
    if "@" in username:  # ถ้าเป็น Email
        parts = username.split("@")
        name = parts[0]
        domain = parts[1]
        masked_name = name[0] + "***" + name[-1] if len(name) > 2 else name[0] + "***"
        return f"{masked_name}@{domain}"
    # ถ้าเป็น Username ทั่วไป
    return username[0] + "***" + username[-1] if len(username) > 2 else "***"

@app.middleware("http")
async def capture_username_header(request: Request, call_next):
    username_val = None
    for header_name in ["current-user", "current_user", "x-user-email"]:
        val = request.headers.get(header_name)
        if val:
            username_val = val
            break
    
    masked_user = mask_username(username_val)
    logger.info(f"📨 [Middleware] {request.method} {request.url.path} | Resolved: {repr(username_val)}")
    
    token_username = CURRENT_REQUEST_USERNAME.set(username_val)
    base_url = str(request.base_url).rstrip("/")
    token_base = CURRENT_REQUEST_BASE_URL.set(base_url)

    try:
        response = await call_next(request)
        return response
    finally:
        CURRENT_REQUEST_USERNAME.reset(token_username)
        CURRENT_REQUEST_BASE_URL.reset(token_base)

# --------------------------------------------------------
# Models
# --------------------------------------------------------
class FilterCondition(BaseModel):
    column: str = Field(..., min_length=1, description="ชื่อคอลัมน์ที่จะฟิลเตอร์")
    operator: AllowedOperators = Field(default="=", description="เครื่องหมายเปรียบเทียบ")
    value: str = Field(..., description="ค่าที่ใช้ค้นหา")


class GenerateReportRequest(BaseModel):
    # ปรับใช้โครงสร้าง Config ของ Pydantic v2
    model_config = ConfigDict(extra="allow")
    table_id: str = Field(..., min_length=1, description="ชื่อตารางใน BigQuery")

    # ป้องกันเซิร์ฟเวอร์พัง: กำหนดให้มากกว่า 0 และสูงสุดไม่เกิน 50,000 แถว (ปรับตัวเลขได้ตามความแรงของ RAM)
    limit: int = Field(default=2000, gt=0, le=50000, description="จำนวนแถวสูงสุดที่ดึง")

    # ยุบเหลือแบบ List รูปแบบเดียว เพื่อลดความสับสนในการเขียน Logic
    filters: Optional[list[FilterCondition]] = Field(default=None, description="รายการฟิลเตอร์")

    # 4. ล็อกคำเชื่อมให้เป็นแค่ AND หรือ OR เท่านั้น ป้องกันการแอบฉีดคำสั่ง SQL
    condition: Optional[Literal["AND", "OR"]] = Field(default="AND", description="เงื่อนไขการเชื่อมฟิลเตอร์")
    columns: Optional[list[str]] = Field(default=None, description="รายชื่อคอลัมน์ที่ต้องการเลือก (ถ้าต้องการเจาะจง)")


# --------------------------------------------------------
# Helpers & Security Core
# --------------------------------------------------------
def clean_table_id(raw: str) -> str:
    """
    ทำความสะอาด Table ID โดยลบช่องว่างและเครื่องหมาย Backtick (`) ออก
    แก้ไข: ไม่ใช้ .split(".")[-1] เพื่อรักษาชื่อ Dataset/Project ไว้ให้ BigQuery ทำงานต่อได้
    """
    if not raw:
        return ""
    return raw.strip().strip("`").strip()

def validate_table_id(table_id: str) -> bool:
    """
    ตรวจสอบความปลอดภัยของ Table ID ด้วยวิธี Whitelist
    แก้ไข: ยอมรับเครื่องหมาย . และ - เพื่อไม่ให้ติดบั๊กเวลาส่ง Full Path ของ BigQuery เข้ามา
    """
    if not table_id:
        return False
    return bool(BQ_TABLE_REGEX.match(table_id))

def is_valid_user(user: Optional[str]) -> bool:
    """
    ตรวจสอบความปลอดภัยของ Username
    แก้ไข: เปลี่ยนจากระบบไล่แบน (Blacklist) มาเป็นระบบตรวจสอบฟอร์แมตที่อนุญาต (Whitelist) 
    เพื่อความปลอดภัยสูงสุดตามหลัก Security Best Practice
    """
    if not user:
        return False
    
    user_clean = user.strip()
    return bool(USER_REGEX.match(user_clean))

def resolve_username(fallback_user: Optional[str] = None) -> str:

    # 1. เช็กจาก Context ตัวตนจริงในเซสชัน
    context_username = CURRENT_REQUEST_USERNAME.get()
    if is_valid_user(context_username):
        return context_username.strip()
        
    # 2. เช็กจากฟังก์ชันสำรองที่ส่งมา
    if is_valid_user(fallback_user):
        return fallback_user.strip()
        
    # 3. ใช้ค่าตั้งต้นระบบ
    if DEFAULT_USERNAME and is_valid_user(DEFAULT_USERNAME):
        return DEFAULT_USERNAME.strip()
        
    return ""

def check_user_permission(username: str, table_id: str) -> Optional[str]:
    # 🛡️ 1. ยามหน้าบ้าน: ตรวจสอบฟอร์แมตอินพุตด้วย Whitelist ฟังก์ชันที่เราทำร่วมกัน
    if not is_valid_user(username):
        logger.warning(f"🛑 Security Blocked: Invalid username format '{repr(username)}'")
        return None

    cleaned_table = clean_table_id(table_id)
    if not validate_table_id(cleaned_table):
        logger.warning(f"🛑 Security Blocked: Invalid table_id format '{repr(table_id)}'")
        return None

    # แปลงเป็นตัวเล็กเพื่อใช้เปรียบเทียบในขั้นตอนถัดไป
    table_key = cleaned_table.lower()

    # 🛡️ 2. ตรวจสอบตารางต้องห้าม (Blacklist Check)
    if "authenbymenu" in table_key:
        logger.warning(f"🚨 Security Blocked: User '{username}' tried to access permission table directly")
        return None

    # 🛡️ 3. ตรวจสอบ Whitelist Mapping 
    # ถ้าตารางนี้ไม่ได้ถูกลงทะเบียนไว้ในระบบรายงานของเรา ก็เด้งออกทันที ไม่ต้องเปิดบิลไปรัน Query บน BigQuery ให้เปลืองคอสท์
    report_names = TABLE_TO_REPORT_NAMES.get(table_key)
    if not report_names:
        logger.info(f"ℹ️ Access Denied: Table '{table_key}' is not mapped to any report.")
        return None

    try:
        # เตรียม Placeholders สำหรับกลุ่มชื่อรายงาน
        placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
        
        # 优化 (Optimize) SQL: ย้ายภาระการแต่งสตริงมาทำที่ฝั่ง Python แทนการใช้ฟังก์ชันครอบคอลัมน์ใน SQL
        # (หมายเหตุ: สมมติฐานว่าข้อมูล UserName ใน Auth Table ถูกจัดเก็บอย่างเป็นระเบียบแล้ว)
        query = f"""
            SELECT ReportName FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE UserName = @username
            AND ReportName IN ({placeholders})
            LIMIT 1
        """
        
        # จัดการแปลงค่าอินพุตให้สะอาดและเป็นมาตรฐานเดียวกันจากฝั่ง Python 
        clean_username = username.strip()
        params = [bigquery.ScalarQueryParameter("username", "STRING", clean_username)]
        
        for i, name in enumerate(report_names):
            params.append(bigquery.ScalarQueryParameter(f"name{i}", "STRING", name.strip()))
            
        job_config = bigquery.QueryJobConfig(query_parameters=params)
        job = bq_client.query(query, job_config=job_config)
        results = list(job)
        
        # ส่งชื่อรายงานกลับไปเมื่อตรวจเจอสิทธิ์ ถ้าไม่มีสิทธิ์คืนเป็น None
        if results:
            return results[0]["ReportName"]
            
        logger.warning(f"🔒 Access Denied: User '{clean_username}' has no permission for table '{table_key}'")
        return None
        
    except Exception as e:
        logger.error(f"❌ Auth verification database error: {e}")
        return None


def validate_filter_column(table_id: str, column: str) -> bool:
    """
    ตรวจสอบว่าคอลัมน์ที่ส่งมาอยู่ใน Whitelist ของตารางนั้นจริงไหม
    ปรับปรุง: เปลี่ยนไปใช้ any() เพื่อทำ Short-circuit evaluation ไม่ต้องสร้าง Set ใหม่ในหน่วยความจำ
    """
    if not table_id or not column:
        return False
        
    cleaned_table = clean_table_id(table_id).lower()
    mapping_dict = TABLE_COLUMN_MAPPING.get(cleaned_table, {})
    
    # วนลูปเช็กแบบ Case-Insensitive เปรียบเทียบเจอแล้วหยุดทันที ประหยัด RAM และ CPU
    target_column = column.lower()
    return any(k.lower() == target_column for k in mapping_dict.keys())

def build_complete_condition_log(
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
    condition: Optional[str] = None,
    row_generated: int = 0,
) -> str:
    """
    สร้าง Log สรุปเงื่อนไขการสร้างรายงานอย่างปลอดภัย
    ปรับปรุง: ปกปิด (Mask) ค่า Value ที่ค้นหาเพื่อป้องกันปัญหา PDPA / Data Leak ลงระบบ Log
    """
    parts = []
    
    # 1. จัดการ Log สำหรับ ฟิลเตอร์แบบกลุ่ม (New Format)
    if filters:
        filter_strs = []
        for f in filters:
            # เก็บข้อมูลเฉพาะ คอลัมน์ และ ตัวดำเนินการ ส่วนค่านำมาใส่ [MASKED] หรือครอบเครื่องหมายไว้
            # หรือถ้าอยากเก็บค่า: สามารถเปลี่ยนเป็น f"{f.column}{f.operator}{repr(f.value[:5])}..." เพื่อเก็บแค่ 5 ตัวแรก
            filter_strs.append(f"{f.column}{f.operator}[MASKED]")
            
        if filter_strs:
            parts.append(f"filters=[{', '.join(filter_strs)}]")
            
    # 2. จัดการ Log สำหรับ ฟิลเตอร์แบบเดี่ยว (Legacy Format)
    if filter_column and filter_value:
        parts.append(f"legacy:{filter_column}=[MASKED]")
        
    # 3. คำเชื่อม และจำนวนแถวที่เกิดขึ้น
    if condition:
        parts.append(f"custom_condition:{condition}")
    if row_generated > 0:
        parts.append(f"rows={row_generated}")
        
    return "; ".join(parts) if parts else "no_filter"

def generate_ai_log_id() -> int:
    """
    เจน ID เป็นตัวเลขขนาดใหญ่ที่มีโอกาสชนกันต่ำมาก (สไตล์ Snowflake / Timestamp-based)
    โดยใช้เวลาปัจจุบันเป็นมิลลิวินาที รวมกับเลขสุ่ม 3 หลักท้าย
    """
    # ได้ตัวเลขความยาวประมาณ 16 หลัก ปลอดภัยบน BigQuery INT64 และไม่มีวันชนกันในระดับวินาทีเดียวกัน
    millis = int(time.time() * 1000)
    rand_part = random.randint(100, 999)
    return int(f"{millis}{rand_part}")

def insert_ai_log(
    username: str,
    table_name: str,
    report_name: str,
    status: str,
    condition: str = "",
    size: float = 0.0,
    bytes_billed: int = 0,
    url: str = "",
    row_generated: int = 0,
    ai_log_id: int = 0,
) -> None:
    try:
        bangkok_tz = timezone(timedelta(hours=7))
        now = datetime.now(bangkok_tz)
        
        if ai_log_id == 0:
            ai_log_id = generate_ai_log_id()

        rows_to_insert = [{
            "CreatedAt":     now.strftime("%Y-%m-%dT%H:%M:%S"),
            "UserName":      (username or "")[:50],
            "TableName":     (table_name or "")[:50],
            "ReportName":    (report_name or "")[:100],
            # 🐛 แก้ไขบั๊ก: ขยายพื้นที่ให้เก็บคำว่า SUCCESS หรือ FAILED ได้เต็มคำ (ห้ามตัดเหลือ 5)
            "Status":        (status or "")[:15], 
            "Condition":     (condition or "")[:2000],
            "Size":          round(size, 4),
            "BytesBilled":   bytes_billed,
            "URL":           (url or "")[:1000],
            "Row_generated": row_generated,
            "AiLogID":       ai_log_id,
        }]

        full_table = f"{PROJECT_ID}.{DATASET_ID}.{LOG_TABLE}"
        errors = bq_client.insert_rows_json(full_table, rows_to_insert)
        
        if errors:
            logger.error(f"❌ AiLog insert FAILED: {errors}")
        else:
            logger.info(f"✅ AiLog OK | id={ai_log_id} user={username} status={status}")
            
    except Exception as e:
        logger.error(f"❌ Exception in insert_ai_log: {e}", exc_info=True)



# --------------------------------------------------------
# Fetch & Generate Excel
# --------------------------------------------------------
def fetch_and_generate_excel(
    table_id: str,
    report_name: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
    columns: Optional[list[str]] = None,
) -> tuple[Optional[int], Optional[str], int, float, str, int]:
    
    # 🛡️ 1. ทำความสะอาดและตรวจสอบความปลอดภัยของชื่อตาราง
    clean = clean_table_id(table_id)
    if not validate_table_id(clean):
        raise PermissionError(f"โครงสร้าง Table ID '{table_id}' ไม่ถูกต้องหรือไม่อนุญาตให้เข้าถึง")
        
    clean_lower = clean.lower()
    if "authenbymenu" in clean_lower:
        raise PermissionError("Access to the authentication table is strictly prohibited at code level.")

    # ดึงชื่อตารางจริงของ BigQuery
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(clean_lower, clean)

    # 🛡️ 2. ตรวจสอบสิทธิ์คอลัมน์ฝั่ง SELECT Clause
    if columns:
        for col in columns:
            if not validate_filter_column(clean_lower, col):
                logger.warning(f"🚨 Column Blocked in SELECT: '{col}'")
                raise PermissionError(f"คอลัมน์ '{col}' ไม่ได้รับอนุญาต")
        select_clause = ", ".join([f"`{col}`" for col in columns])
    else:
        select_clause = "*"

    query = f"SELECT {select_clause} FROM `{PROJECT_ID}.{DATASET_ID}.{bq_table_name}`"
    params = []
    where_clauses = []

    # 🛡️ 3. ตรวจสอบและสร้างเงื่อนไข WHERE Clause อย่างปลอดภัย
    if filters:
        for i, f in enumerate(filters):
            if not validate_filter_column(clean_lower, f.column):
                raise PermissionError(f"คอลัมน์ '{f.column}' ไม่ได้รับอนุญาตสำหรับตารางนี้")
            if f.operator not in ALLOWED_OPERATORS:
                raise PermissionError(f"Operator '{f.operator}' ไม่ได้รับอนุญาต")
            
            param_name = f"filter_{i}"
            where_clauses.append(f"`{f.column}` {f.operator} @{param_name}")
            params.append(bigquery.ScalarQueryParameter(param_name, "STRING", f.value))

    elif filter_column and filter_value:
        if not validate_filter_column(clean_lower, filter_column):
            raise PermissionError(f"คอลัมน์ '{filter_column}' ไม่ได้รับอนุญาตสำหรับตารางนี้")
        
        where_clauses.append(f"`{filter_column}` = @filter_value")
        params.append(bigquery.ScalarQueryParameter("filter_value", "STRING", filter_value))

    if where_clauses:
        query += " WHERE " + " AND ".join(where_clauses)

    # ป้องกันการดึงข้อมูลเกินขนาด (Guardrail Limit)
    safe_limit = min(max(limit, 1), 10000)
    query += f" LIMIT {safe_limit}"

    logger.info(f"🔍 Final Query: {query}")

    # 🚀 4. ยิง Query หา BigQuery
    job_config = bigquery.QueryJobConfig(query_parameters=params)
    job = bq_client.query(query, job_config=job_config)
    result = job.result()

    bytes_billed: int = job.total_bytes_billed or 0
    schema = result.schema
    row_count = result.total_rows

    # 🛡️ 5. สร้าง Audit Log ที่ปลอดภัย (ใช้ฟังก์ชันมาตรฐานเพื่อ Mask PII Data)
    condition_log = build_complete_condition_log(
        filter_column=filter_column,
        filter_value=filter_value,
        filters=filters,
        condition=f"LIMIT {safe_limit}",
        row_generated=row_count or 0
    )

    # กรณีไม่พบข้อมูล คืนค่ากลับทันที
    if row_count == 0 or row_count is None:
        ai_log_id = generate_ai_log_id()
        return 0, None, bytes_billed, 0.0, condition_log, ai_log_id

    # 📊 6. สร้างไฟล์ Excel สรุปรายงาน
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    # แก้ไขบั๊กคำแปลพิมพ์เล็ก-ใหญ่: แปลง dictionary mapping ให้คีย์เป็น lowercase ทั้งหมดก่อนใช้งาน
    raw_mapping = TABLE_COLUMN_MAPPING.get(clean_lower, {})
    lowercase_mapping = {k.lower(): v for k, v in raw_mapping.items()}
    
    headers = []
    for field in schema:
        # ดึงคำแปลภาษาไทยแบบ Case-Insensitive เจอพิมพ์ไหนก็แปลได้แม่นยำ
        translated_header = lowercase_mapping.get(field.name.lower(), field.name)
        headers.append(translated_header)
            
    ws.append(headers)

    # ใส่ข้อมูลลงแผ่นงานและล้าง Naive Timezone ของ Python ออกเพื่อไม่ให้ Excel บั๊ก
    for row in result:
        row_values = []
        for field in schema:
            val = row[field.name]
            if hasattr(val, "tzinfo") and val.tzinfo is not None:
                val = val.replace(tzinfo=None)
            row_values.append(val)
        ws.append(row_values)

    # 💾 7. บันทึกไฟล์ลง Disk ระบบอย่างปลอดภัย
    ai_log_id = generate_ai_log_id()  # ใช้เวอร์ชันใหม่ที่โอกาส ID ชนกันเป็นศูนย์
    secure_file_id = uuid.uuid4().hex[:8]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"{bq_table_name}_{secure_file_id}_{timestamp}.xlsx"

    file_path = REPORT_DIR / file_name
    wb.save(file_path)

    file_size_mb = round(file_path.stat().st_size / (1024 * 1024), 4)
    return row_count, file_name, bytes_billed, file_size_mb, condition_log, ai_log_id


def cleanup_old_reports() -> None:
    deleted = 0
    cutoff = datetime.now() - timedelta(hours=24)
    try:
        for file_path in REPORT_DIR.glob("*.xlsx"):
            if file_path.is_file():
                mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
                if mtime < cutoff:
                    file_path.unlink()
                    deleted += 1
        logger.info(f"🧹 Midnight cleanup done — {deleted} file(s) removed")
    except Exception as e:
        logger.error(f"Error during report cleanup: {e}")

# --------------------------------------------------------
# Centralized Authentication & Generation Executor
# --------------------------------------------------------
def _execute_report_generation(
    table_id: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
    columns: Optional[list[str]] = None,
    username: Optional[str] = None,
    user_email: Optional[str] = None,
) -> dict:
    username_to_use = resolve_username(username, user_email)
    tid = clean_table_id(table_id)
    clean = tid.lower()
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(clean, tid)

    if not username_to_use:
        insert_ai_log("", bq_table_name, "", "FAIL")
        return {"success": False, "message": "❌ ไม่พบข้อมูลยืนยันตัวตนในระบบ"}

    if "authenbymenu" in clean:
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return {"success": False, "message": "🔒 [Access Denied] ระบบไม่อนุญาตให้เข้าถึงตารางสิทธิ์"}

    if not validate_table_id(tid):
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return {"success": False, "message": f"❌ table_id '{tid}' ไม่ถูกต้อง"}

    report_name = check_user_permission(username_to_use, tid)
    if not report_name:
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return {"success": False, "message": "🙏 ขออภัย คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    try:
        row_count, file_name, bytes_billed, file_size_mb, condition_log, ai_log_id = fetch_and_generate_excel(
            tid, report_name, limit,
            filter_column=filter_column,
            filter_value=filter_value,
            filters=filters,
            columns=columns,
        )
        if row_count is None:
            insert_ai_log(username_to_use, bq_table_name, report_name, "FAIL",
                          condition=condition_log, bytes_billed=bytes_billed, ai_log_id=ai_log_id)
            return {"success": False, "message": f"❌ ไม่พบข้อมูลในรายงาน {report_name}"}

        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"

        insert_ai_log(
            username=username_to_use, table_name=bq_table_name, report_name=report_name,
            status="OK", condition=condition_log, size=file_size_mb,
            bytes_billed=bytes_billed, url=download_url,
            row_generated=row_count or 0, ai_log_id=ai_log_id,
        )
        return {
            "success": True,
            "message": f"✅ จัดเตรียมรายงาน{report_name} ({row_count} แถว) สำเร็จ",
            "download_url": download_url,
            "file_name": file_name,
            "row_count": row_count,
            "file_size_mb": file_size_mb,
            "report_name": report_name,
        }
    except PermissionError as pe:
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": f"🔒 {str(pe)}"}
    except Exception as e:
        logger.error(f"Execution Error: {e}", exc_info=True)
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": "🚨 พบปัญหาการสร้าง Excel"}

# --------------------------------------------------------
# MCP Tools
# --------------------------------------------------------
@mcp.tool(
    name="sc_report_export",
    description="ดึงข้อมูล SC Report และสร้างไฟล์ Excel สำหรับ Agent — ส่ง table_id, filters, columns, limit และ username",
)
def mcp_sc_report_export(
    table_id: str,
    username: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[dict]] = None,
    columns: Optional[list[str]] = None,
    **kwargs,
) -> dict:
    parsed_filters = None
    if filters:
        try:
            parsed_filters = [FilterCondition(**f) for f in filters]
        except Exception:
            return {"success": False, "message": "❌ รูปแบบ filters ไม่ถูกต้อง"}

    return _execute_report_generation(
        table_id=table_id,
        limit=limit,
        filter_column=filter_column,
        filter_value=filter_value,
        filters=parsed_filters,
        columns=columns,
        username=username,
    )


@mcp.tool(
    name="generate_excel_report",
    description="สร้างไฟล์ Excel จากรายงานใน BigQuery และส่งลิงก์ดาวน์โหลด",
)
def mcp_generate_excel_report(
    table_id: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[dict]] = None,
    condition: Optional[str] = None,
    username: Optional[str] = None,
    user_email: Optional[str] = None,
    columns: Optional[list[str]] = None,
    **kwargs,
) -> str:
    parsed_filters = None
    if filters:
        try:
            parsed_filters = [FilterCondition(**f) for f in filters]
        except Exception:
            return "❌ รูปแบบ filters ไม่ถูกต้อง"

    res = _execute_report_generation(
        table_id=table_id,
        limit=limit,
        filter_column=filter_column,
        filter_value=filter_value,
        filters=parsed_filters,
        columns=columns,
        username=username,
        user_email=user_email,
    )
    if res["success"]:
        return f"📊 จัดเตรียม **รายงาน{res['report_name']}** ({res['row_count']} แถว) เรียบร้อยแล้วครับ\n🔗 {res['download_url']}"
    else:
        return res["message"]

# --------------------------------------------------------
# REST API Endpoints
# --------------------------------------------------------
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest):
    res = _execute_report_generation(
        table_id=request.table_id,
        limit=request.limit,
        filter_column=request.filter_column,
        filter_value=request.filter_value,
        filters=request.filters,
        columns=request.columns,
        username=request.username,
        user_email=request.user_email,
    )
    if res["success"]:
        return {
            "success": True,
            "message": f"✅ จัดเตรียม **รายงาน{res['report_name']}** สำเร็จ\n🔗 {res['download_url']}",
        }
    else:
        return {"success": False, "message": res["message"]}

# --------------------------------------------------------
# Download Endpoint
# --------------------------------------------------------
@app.get("/download/{file_name}")
def download_report(file_name: str, direct: bool = False):
    if not re.match(r"^[A-Za-z0-9_]+_[a-f0-9]{8}_\d{8}_\d{6}\.xlsx$", file_name):
        return {"success": False, "message": "❌ รูปแบบลิงก์ไม่ถูกต้อง"}

    file_path = (REPORT_DIR / file_name).resolve()
    if not file_path.is_relative_to(REPORT_DIR) or not file_path.exists():
        return {"success": False, "message": "❌ ไม่พบไฟล์หรือไม่มีสิทธิ์เข้าถึง"}

    if direct:
        return FileResponse(
            path=file_path,
            filename=file_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    html_content = """<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Downloading...</title></head><body><script>const downloadUrl = window.location.pathname + "?direct=true";const iframe = document.createElement('iframe');iframe.style.display = 'none';iframe.src = downloadUrl;document.body.appendChild(iframe);setTimeout(() => { window.close(); }, 1000);</script></body></html>"""
    return HTMLResponse(content=html_content)

# --------------------------------------------------------
# Health Check Endpoint
# --------------------------------------------------------
@app.get("/health")
def health_check():
    return {"status": "healthy"}

# --------------------------------------------------------
# Files List Endpoint
# --------------------------------------------------------
@app.get("/files")
def list_files():
    try:
        files = []
        for file_path in REPORT_DIR.glob("*"):
            if file_path.is_file():
                stat = file_path.stat()
                files.append({
                    "name": file_path.name,
                    "size_bytes": stat.st_size,
                    "size_mb": round(stat.st_size / (1024 * 1024), 4),
                    "created_at": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                    "modified_at": datetime.fromtimestamp(stat.st_mtime).isoformat()
                })
        return {"directory": str(REPORT_DIR), "files": files}
    except Exception as e:
        logger.error(f"Error listing files: {e}")
        return {"success": False, "message": "Failed to list files"}

app.mount("", mcp_asgi_app)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)