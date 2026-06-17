import os
import re
import uuid
import logging
import time
import asyncio
import math
from dataclasses import dataclass
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo
from contextlib import asynccontextmanager
from contextvars import ContextVar
from pathlib import Path
from threading import Lock
from typing import Literal, Optional

from pydantic import BaseModel, Field, ConfigDict
from openpyxl import Workbook
from fastapi import FastAPI, Request, Header, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from google.cloud import bigquery
from mcp.server.fastmcp import FastMCP
from fastapi.middleware.cors import CORSMiddleware

# ============================================================
# SECTION 1 — COST CALCULATOR (inline, no separate file needed)
# ============================================================
BQ_PRICE_PER_TB_USD: float = 6.25
CLOUD_RUN_REQUEST_PRICE_PER_M: float = 0.40
CLOUD_RUN_CPU_PRICE_PER_VCPU_SEC: float = 0.00002400
CLOUD_RUN_MEM_PRICE_PER_GIB_SEC: float = 0.00000250

@dataclass
class BigQueryCost:
    bytes_billed: int
    tb_billed: float
    cost_usd: float
    cost_thb: float

@dataclass
class CloudRunCost:
    duration_seconds: float
    vcpu: float
    memory_gib: float
    request_cost_usd: float
    cpu_cost_usd: float
    memory_cost_usd: float
    total_cost_usd: float
    total_cost_thb: float

@dataclass
class TotalCost:
    bigquery: BigQueryCost
    cloud_run: CloudRunCost
    grand_total_usd: float
    grand_total_thb: float
    usd_to_thb_rate: float
    summary: str

def calculate_bigquery_cost(
    bytes_billed: int,
    price_per_tb_usd: float = BQ_PRICE_PER_TB_USD,
    usd_to_thb: float = 35.0,
) -> BigQueryCost:
    safe_bytes = max(int(bytes_billed), 0)
    tb_billed = safe_bytes / (1024 ** 4)
    cost_usd = tb_billed * price_per_tb_usd
    cost_thb = cost_usd * usd_to_thb
    return BigQueryCost(
        bytes_billed=safe_bytes,
        tb_billed=round(tb_billed, 10),
        cost_usd=round(cost_usd, 8),
        cost_thb=round(cost_thb, 6),
    )

def calculate_cloud_run_cost(
    duration_seconds: float,
    vcpu: float = 1.0,
    memory_gib: float = 0.5,
    request_price_per_million: float = CLOUD_RUN_REQUEST_PRICE_PER_M,
    cpu_price_per_vcpu_sec: float = CLOUD_RUN_CPU_PRICE_PER_VCPU_SEC,
    mem_price_per_gib_sec: float = CLOUD_RUN_MEM_PRICE_PER_GIB_SEC,
    usd_to_thb: float = 35.0,
) -> CloudRunCost:
    billable_seconds = max(math.ceil(duration_seconds * 10) / 10, 0.1)
    request_cost_usd = request_price_per_million / 1_000_000
    cpu_cost_usd = vcpu * billable_seconds * cpu_price_per_vcpu_sec
    memory_cost_usd = memory_gib * billable_seconds * mem_price_per_gib_sec
    total_usd = request_cost_usd + cpu_cost_usd + memory_cost_usd
    total_thb = total_usd * usd_to_thb
    return CloudRunCost(
        duration_seconds=round(billable_seconds, 2),
        vcpu=vcpu,
        memory_gib=memory_gib,
        request_cost_usd=round(request_cost_usd, 10),
        cpu_cost_usd=round(cpu_cost_usd, 8),
        memory_cost_usd=round(memory_cost_usd, 8),
        total_cost_usd=round(total_usd, 8),
        total_cost_thb=round(total_thb, 6),
    )

def calculate_total_cost(
    bytes_billed: int,
    duration_seconds: float,
    vcpu: float = 1.0,
    memory_gib: float = 0.5,
    usd_to_thb: float = 35.0,
    bq_price_per_tb_usd: float = BQ_PRICE_PER_TB_USD,
) -> TotalCost:
    bq = calculate_bigquery_cost(bytes_billed, bq_price_per_tb_usd, usd_to_thb)
    cr = calculate_cloud_run_cost(duration_seconds, vcpu, memory_gib, usd_to_thb=usd_to_thb)
    grand_usd = bq.cost_usd + cr.total_cost_usd
    grand_thb = grand_usd * usd_to_thb
    summary = (
        f"BQ: ${bq.cost_usd:.6f} ({bq.tb_billed:.6f} TB) | "
        f"Cloud Run: ${cr.total_cost_usd:.6f} ({cr.duration_seconds}s) | "
        f"Total: ${grand_usd:.6f} / ฿{grand_thb:.4f}"
    )
    return TotalCost(
        bigquery=bq,
        cloud_run=cr,
        grand_total_usd=round(grand_usd, 8),
        grand_total_thb=round(grand_thb, 6),
        usd_to_thb_rate=usd_to_thb,
        summary=summary,
    )

def format_cost_breakdown(cost: TotalCost) -> dict:
    return {
        "bigquery": {
            "bytes_billed": cost.bigquery.bytes_billed,
            "tb_billed": cost.bigquery.tb_billed,
            "cost_usd": cost.bigquery.cost_usd,
            "cost_thb": cost.bigquery.cost_thb,
        },
        "cloud_run": {
            "duration_seconds": cost.cloud_run.duration_seconds,
            "vcpu": cost.cloud_run.vcpu,
            "memory_gib": cost.cloud_run.memory_gib,
            "request_cost_usd": cost.cloud_run.request_cost_usd,
            "cpu_cost_usd": cost.cloud_run.cpu_cost_usd,
            "memory_cost_usd": cost.cloud_run.memory_cost_usd,
            "total_cost_usd": cost.cloud_run.total_cost_usd,
            "total_cost_thb": cost.cloud_run.total_cost_thb,
        },
        "grand_total_usd": cost.grand_total_usd,
        "grand_total_thb": cost.grand_total_thb,
        "usd_to_thb_rate": cost.usd_to_thb_rate,
        "summary": cost.summary,
    }

# ============================================================
# SECTION 2 — CONFIGURATION
# ============================================================
BANGKOK_TZ = ZoneInfo("Asia/Bangkok")
PROJECT_ID = os.environ.get("PROJECT_ID", "sc-ai-uat")
DATASET_ID = os.environ.get("DATASET_ID", "SCReport")
AUTH_TABLE = os.environ.get("AUTH_TABLE", "AuthenByMenu")
LOG_TABLE = os.environ.get("LOG_TABLE", "AiLog")
DEFAULT_USERNAME = ""
PUBLIC_BASE_URL = os.environ.get(
    "PUBLIC_BASE_URL",
    "https://sc-report-866803019306.asia-southeast3.run.app",
)

COST_USD_TO_THB = float(os.environ.get("COST_USD_TO_THB", "35.0"))
COST_CLOUD_RUN_VCPU = float(os.environ.get("COST_CLOUD_RUN_VCPU", "1.0"))
COST_CLOUD_RUN_MEM_GIB = float(os.environ.get("COST_CLOUD_RUN_MEM_GIB", "0.5"))

CLEANUP_SECRET = os.environ.get("CLEANUP_SECRET", "super-secret-sc-cleanup-2025")

# ============================================================
# SECTION 3 — MAPPING
# ============================================================
from mapping import (
    TABLE_TO_REPORT_NAMES,
    TABLE_TO_BQ_TABLE_NAME,
    ALLOWED_OPERATORS,
    TABLE_COLUMN_MAPPING,
    USER_REGEX,
    BQ_TABLE_REGEX,
    SENSITIVE_COLUMNS,
)

REPORT_DIR = Path("/tmp/reports").resolve()
REPORT_DIR.mkdir(parents=True, exist_ok=True)

bq_client = bigquery.Client(project=PROJECT_ID)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("SCReportSecurity")

CURRENT_REQUEST_USERNAME: ContextVar[Optional[str]] = ContextVar(
    "current_request_username", default=None
)
CURRENT_REQUEST_BASE_URL: ContextVar[Optional[str]] = ContextVar(
    "current_request_base_url", default=None
)

# ============================================================
# SECTION 4 — LOG ID GENERATOR
# ============================================================
class _LogIDGenerator:
    _counter: int = 0
    _last_millis: int = 0
    _lock: Lock = Lock()

    @classmethod
    def generate(cls) -> int:
        with cls._lock:
            millis = int(time.time() * 1000)
            if millis != cls._last_millis:
                cls._counter = 0
                cls._last_millis = millis
            else:
                cls._counter += 1
            return int(f"{millis}{cls._counter:03d}")

def generate_ai_log_id() -> int:
    return _LogIDGenerator.generate()

# ============================================================
# SECTION 6 — MCP SERVER & FASTAPI APP
# ============================================================
mcp = FastMCP(
    "sc-report-server",
    stateless_http=True,
    json_response=True,
    host="0.0.0.0",
)

mcp_asgi_app = mcp.streamable_http_app()

@asynccontextmanager
async def lifespan(app: FastAPI):
    try:
        async with mcp.session_manager.run():
            logger.info("✅ MCP Session started")
            yield
    finally:
        logger.info("🛑 App shutting down")

app = FastAPI(
    title="SC Report MCP API & Server",
    description="API & MCP Server สำหรับตรวจสอบสิทธิ์และสร้าง Excel Report จาก BigQuery",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://ai-uat.scasset.com"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============================================================
# SECTION 7 — MIDDLEWARE
# ============================================================
@app.middleware("http")
async def preserve_method_on_redirect(request: Request, call_next):
    response = await call_next(request)
    if response.status_code == 302 and request.method in ("POST", "PUT", "PATCH", "DELETE"):
        location = response.headers.get("location", "")
        logger.info(
            f"🔀 [RedirectFix] 302→307 {request.method} {request.url.path} → {location}"
        )
        response.status_code = 307
    return response

@app.middleware("http")
async def capture_username_header(request: Request, call_next):
    username_val = None
    for header_name in ["current-user", "current_user", "x-user-email"]:
        val = request.headers.get(header_name)
        if val:
            username_val = val
            break

    masked_user = mask_username(username_val)
    logger.info(
        f"📨 [Middleware] {request.method} {request.url.path} | User: {masked_user}"
    )

    token_username = CURRENT_REQUEST_USERNAME.set(username_val)
    base_url = str(request.base_url).rstrip("/")
    token_base = CURRENT_REQUEST_BASE_URL.set(base_url)

    try:
        response = await call_next(request)
        return response
    finally:
        CURRENT_REQUEST_USERNAME.reset(token_username)
        CURRENT_REQUEST_BASE_URL.reset(token_base)

# ============================================================
# SECTION 8 — PYDANTIC MODELS
# ============================================================
AllowedOperators = Literal["=", ">", "<", ">=", "<=", "LIKE", "!="]

class FilterCondition(BaseModel):
    column: str = Field(..., min_length=1, description="ชื่อคอลัมน์ที่จะฟิลเตอร์")
    operator: AllowedOperators = Field(default="=", description="เครื่องหมายเปรียบเทียบ")
    value: str = Field(..., description="ค่าที่ใช้ค้นหา")

class GenerateReportRequest(BaseModel):
    model_config = ConfigDict(extra="allow")
    table_id: str = Field(..., min_length=1, description="ชื่อตารางใน BigQuery")
    limit: int = Field(default=2000, gt=0, le=100000, description="จำนวนแถวสูงสุดที่ดึง")
    filters: Optional[list[FilterCondition]] = Field(default=None, description="รายการฟิลเตอร์")
    condition: Optional[Literal["AND", "OR"]] = Field(default="AND", description="เงื่อนไขการเชื่อมฟิลเตอร์")
    columns: Optional[list[str]] = Field(default=None, description="รายชื่อคอลัมน์ที่ต้องการเลือก")
    username: Optional[str] = Field(default=None, description="ชื่อผู้ใช้งานระบบ")
    user_email: Optional[str] = Field(default=None, description="อีเมลผู้ใช้งานระบบ")
    filter_column: Optional[str] = Field(default=None, description="คอลัมน์กรองข้อมูลแบบเดิม")
    filter_value: Optional[str] = Field(default=None, description="ค่ากรองข้อมูลแบบเดิม")

class CostEstimateRequest(BaseModel):
    bytes_billed: int = Field(..., ge=0, description="bytes_billed จาก BigQuery job")
    duration_seconds: float = Field(..., gt=0, description="เวลา request รวม (วินาที)")
    vcpu: float = Field(default=1.0, gt=0, description="vCPU ที่ Cloud Run ใช้")
    memory_gib: float = Field(default=0.5, gt=0, description="Memory (GiB) ที่ Cloud Run ใช้")
    usd_to_thb: float = Field(default=35.0, gt=0, description="อัตราแลกเปลี่ยน USD→THB")

# ============================================================
# SECTION 9 — HELPERS & SECURITY (UPDATED)
# ============================================================
def mask_username(username: str | None) -> str:
    if not username:
        return "Anonymous"
    if "@" in username:
        parts = username.split("@", 1)
        name, domain = parts[0], parts[1]
        if len(name) == 0:
            masked_name = "***"
        elif len(name) <= 2:
            masked_name = name[0] + "***"
        else:
            masked_name = name[0] + "***" + name[-1]
        return f"{masked_name}@{domain}"
    if len(username) == 0:
        return "***"
    if len(username) <= 2:
        return username[0] + "***"
    return username[0] + "***" + username[-1]

def clean_table_id(raw: str) -> str:
    if not raw:
        return ""
    return raw.strip().strip("`").strip()

def validate_table_id(table_id: str) -> bool:
    if not table_id:
        return False
    full_bq_regex = re.compile(r"^[A-Za-z0-9_-]+\.[A-Za-z0-9_]+\.[A-Za-z0-9_]+$")
    return bool(BQ_TABLE_REGEX.match(table_id)) or bool(full_bq_regex.match(table_id))

def extract_short_table_name(table_id: str) -> str:
    clean = table_id.strip().strip("`").strip()
    return clean.split(".")[-1].lower()

def is_valid_user(user: Optional[str]) -> bool:
    if not user:
        return False
    return bool(USER_REGEX.match(user.strip()))

def is_sensitive_column(column: str) -> bool:
    return bool(column and column.lower() in SENSITIVE_COLUMNS)

def resolve_username(
    fallback_user: Optional[str] = None,
    fallback_email: Optional[str] = None,
) -> str:
    context_username = CURRENT_REQUEST_USERNAME.get()
    if is_valid_user(context_username):
        return context_username.strip()
    if is_valid_user(fallback_user):
        return fallback_user.strip()
    if fallback_email and "@" in fallback_email:
        email_prefix = fallback_email.split("@")[0]
        if is_valid_user(email_prefix):
            return email_prefix.strip()
    if DEFAULT_USERNAME and is_valid_user(DEFAULT_USERNAME):
        return DEFAULT_USERNAME.strip()
    return ""

def check_user_permission(username: str, table_id: str) -> Optional[str]:
    if not is_valid_user(username):
        logger.warning(f"🛑 Security Blocked: Invalid username format '{repr(username)}'")
        return None

    cleaned_table = clean_table_id(table_id)
    if not validate_table_id(cleaned_table):
        logger.warning(f"🛑 Security Blocked: Invalid table_id format '{repr(table_id)}'")
        return None

    short_name = extract_short_table_name(cleaned_table)
    logger.info(f"🔍 [Auth] table_id='{table_id}' -> short_name='{short_name}' | user='{username}'")

    if "authenbymenu" in short_name:
        logger.warning(f"🚨 Security Blocked: User '{username}' tried to access permission table directly")
        return None

    report_names = TABLE_TO_REPORT_NAMES.get(short_name)
    if not report_names:
        logger.info(f"ℹ️ Access Denied: Table '{short_name}' is not mapped to any report.")
        return None

    try:
        clean_username = username.strip()
        placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
        query = f"""
            SELECT ReportName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE UserName = @username
              AND ReportName IN ({placeholders})
            LIMIT 1
        """
        params = [bigquery.ScalarQueryParameter("username", "STRING", clean_username)]
        for i, name in enumerate(report_names):
            params.append(bigquery.ScalarQueryParameter(f"name{i}", "STRING", name.strip()))

        job_config = bigquery.QueryJobConfig(query_parameters=params)
        job = bq_client.query(query, job_config=job_config)
        results = list(job)

        if results:
            matched_report = results[0]["ReportName"]
            return matched_report

        logger.warning(f"🔒 Access Denied: User '{clean_username}' has no permission for table '{short_name}'")
        return None
    except Exception as e:
        logger.error(f"❌ Auth verification database error: {e}")
        return None

def validate_filter_column(table_id: str, column: str) -> bool:
    if not table_id or not column:
        return False
    short_name = extract_short_table_name(table_id)
    mapping_dict = TABLE_COLUMN_MAPPING.get(short_name, {})
    target_column = column.lower()
    return any(k.lower() == target_column for k in mapping_dict.keys())

# เพิ่มตัวแปรสำหรับรับค่าข้อมูลเงินแยกตาม Schema ใหม่ของตาราง AiLog
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
    cost_thb: float = 0.0,
    cost_usd: float = 0.0,
    cost_summary: str = "",
) -> None:
    try:
        bangkok_tz = timezone(timedelta(hours=7))
        now = datetime.now(bangkok_tz)

        if ai_log_id == 0:
            ai_log_id = generate_ai_log_id()

        rows_to_insert = [
            {
                "CreatedAt": now.strftime("%Y-%m-%dT%H:%M:%S"),
                "UserName": (username or "")[:50],
                "TableName": (table_name or "")[:50],
                "ReportName": (report_name or "")[:100],
                "Status": (status or "")[:15],
                "Condition": (condition or "")[:2000], # จุดนี้จะใช้เก็บ SQL Statement จริง
                "Size": round(size, 4),
                "BytesBilled": bytes_billed,
                "URL": (url or "")[:1000],
                "Row_generated": row_generated,
                "AiLogID": ai_log_id,
                "CostTHB": round(cost_thb, 6),        # บันทึกยอดรวมเงินบาทลงตาราง
                "CostUSD": round(cost_usd, 8),        # บันทึกยอดรวมดอลลาร์ลงตาราง
                "CostSummary": (cost_summary or "")[:500], # เก็บบันทึกตัวหนังสือย่อยของ Cloud Run และ BQ
            }
        ]
        full_table = f"{PROJECT_ID}.{DATASET_ID}.{LOG_TABLE}"
        errors = bq_client.insert_rows_json(full_table, rows_to_insert)

        if errors:
            logger.error(f"❌ AiLog insert FAILED: {errors}")
    except Exception as e:
        logger.error(f"❌ Exception in insert_ai_log: {e}", exc_info=True)

def map_param_type_and_value(column_name: str, str_val: str, schema_dict: dict) -> tuple[str, any]:
    col_type = schema_dict.get(column_name.lower())
    if not col_type:
        return "STRING", str_val
    if col_type in ("INTEGER", "INT64"):
        try:
            return "INT64", int(float(str_val))
        except ValueError:
            pass
    elif col_type in ("FLOAT", "FLOAT64", "NUMERIC", "BIGNUMERIC"):
        try:
            return "FLOAT64", float(str_val)
        except ValueError:
            pass
    elif col_type in ("BOOLEAN", "BOOL"):
        return "BOOL", str_val.lower() in ("true", "1", "yes")
    return "STRING", str_val

# ============================================================
# SECTION 10 — FETCH & GENERATE EXCEL (UPDATED TO LOG RAW SQL)
# ============================================================
def fetch_and_generate_excel(
    table_id: str,
    report_name: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
    columns: Optional[list[str]] = None,
) -> tuple[Optional[int], Optional[str], int, float, str, int]:

    clean = clean_table_id(table_id)
    if not validate_table_id(clean):
        raise PermissionError(f"โครงสร้าง Table ID '{table_id}' ไม่ถูกต้องหรือไม่อนุญาตให้เข้าถึง")

    short_name = extract_short_table_name(clean)
    if "authenbymenu" in short_name:
        raise PermissionError("Access to the authentication table is strictly prohibited at code level.")

    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(short_name, clean)
    table_ref = f"{PROJECT_ID}.{DATASET_ID}.{bq_table_name}"
    
    try:
        table = bq_client.get_table(table_ref)
        schema_dict = {field.name.lower(): field.field_type for field in table.schema}
    except Exception as e:
        logger.error(f"⚠️ [Schema] Failed to fetch schema: {e}")
        schema_dict = {}

    if columns:
        for col in columns:
            if not validate_filter_column(short_name, col):
                raise PermissionError(f"คอลัมน์ '{col}' ไม่ได้รับอนุญาต")
        select_clause = ", ".join([f"`{col}`" for col in columns])
    else:
        select_clause = "*"

    query = f"SELECT {select_clause} FROM `{PROJECT_ID}.{DATASET_ID}.{bq_table_name}`"
    params = []
    where_clauses = []

    if filters:
        for i, f in enumerate(filters):
            if not validate_filter_column(short_name, f.column):
                raise PermissionError(f"คอลัมน์ '{f.column}' ไม่ได้รับอนุญาตสำหรับตารางนี้")
            if f.operator not in ALLOWED_OPERATORS:
                raise PermissionError(f"Operator '{f.operator}' ไม่ได้รับอนุญาต")
            param_name = f"filter_{i}"
            if f.operator == "LIKE":
                where_clauses.append(f"CAST(`{f.column}` AS STRING) LIKE @{param_name}")
                params.append(bigquery.ScalarQueryParameter(param_name, "STRING", f.value))
            else:
                p_type, p_val = map_param_type_and_value(f.column, f.value, schema_dict)
                where_clauses.append(f"`{f.column}` {f.operator} @{param_name}")
                params.append(bigquery.ScalarQueryParameter(param_name, p_type, p_val))
    elif filter_column and filter_value:
        if not validate_filter_column(short_name, filter_column):
            raise PermissionError(f"คอลัมน์ '{filter_column}' ไม่ได้รับอนุญาตสำหรับตารางนี้")
        p_type, p_val = map_param_type_and_value(filter_column, filter_value, schema_dict)
        where_clauses.append(f"`{filter_column}` = @filter_value")
        params.append(bigquery.ScalarQueryParameter("filter_value", p_type, p_val))

    if where_clauses:
        query += " WHERE " + " AND ".join(where_clauses)

    safe_limit = min(max(limit, 1), 50000)
    query += f" LIMIT {safe_limit}"

    # บันทึกตัวแปรภาษา SQL แท้เพื่อนำไปใช้งานต่อในช่อง Condition ของ Log
    actual_sql_executed = query

    job_config = bigquery.QueryJobConfig(query_parameters=params)
    job = bq_client.query(query, job_config=job_config)
    result = job.result()

    bytes_billed: int = job.total_bytes_billed or 0
    schema = result.schema
    row_count = result.total_rows

    if not row_count:
        ai_log_id = generate_ai_log_id()
        return 0, None, bytes_billed, 0.0, actual_sql_executed, ai_log_id

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    raw_mapping = TABLE_COLUMN_MAPPING.get(short_name, {})
    lowercase_mapping = {k.lower(): v for k, v in raw_mapping.items()}

    headers = []
    for field in schema:
        translated_header = lowercase_mapping.get(field.name.lower(), field.name)
        headers.append(translated_header)
    ws.append(headers)

    for row in result:
        row_values = []
        for field in schema:
            val = row[field.name]
            if hasattr(val, "tzinfo") and val.tzinfo is not None:
                val = val.replace(tzinfo=None)
            row_values.append(val)
        ws.append(row_values)

    ai_log_id = generate_ai_log_id()
    secure_file_id = uuid.uuid4().hex[:8]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"{bq_table_name}_{secure_file_id}_{timestamp}.xlsx"

    file_path = REPORT_DIR / file_name
    wb.save(file_path)

    file_size_mb = round(file_path.stat().st_size / (1024 * 1024), 4)
    return row_count, file_name, bytes_billed, file_size_mb, actual_sql_executed, ai_log_id

# ============================================================
# SECTION 11 — CLEANUP OLD REPORTS
# ============================================================
def cleanup_old_reports() -> None:
    cleanup_strategy = os.environ.get("CLEANUP_STRATEGY", "age").lower()
    cleanup_age_seconds = int(os.environ.get("CLEANUP_AGE_SECONDS", "86400"))
    cleanup_max_files = int(os.environ.get("CLEANUP_MAX_FILES", "20"))
    
    logger.info(f"🧹 Starting cleanup: strategy={cleanup_strategy}, age_threshold={cleanup_age_seconds}s, max_files={cleanup_max_files}")
    
    if not REPORT_DIR.exists():
        logger.info(f"🧹 Report directory does not exist: {REPORT_DIR}")
        return
    
    deleted = 0
    current_time = time.time()
    
    try:
        if cleanup_strategy == "age":
            cutoff_timestamp = current_time - cleanup_age_seconds
            for file_path in REPORT_DIR.glob("*.xlsx"):
                try:
                    if file_path.is_file():
                        file_mtime = file_path.stat().st_mtime
                        file_age_seconds = current_time - file_mtime
                        if file_mtime < cutoff_timestamp:
                            file_path.unlink()
                            deleted += 1
                            logger.info(f"✓ Deleted: {file_path.name} (age: {file_age_seconds:.1f}s)")
                except FileNotFoundError:
                    pass
                except Exception as file_error:
                    logger.warning(f"⚠️ Skipped '{file_path.name}' due to error: {file_error}")
        
        elif cleanup_strategy == "keep_latest":
            files = sorted(REPORT_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
            for file_path in files[cleanup_max_files:]:
                try:
                    if file_path.is_file():
                        file_path.unlink()
                        deleted += 1
                        logger.info(f"✓ Deleted old file: {file_path.name}")
                except FileNotFoundError:
                    pass
                except Exception as file_error:
                    logger.warning(f"⚠️ Skipped '{file_path.name}': {file_error}")
        
        elif cleanup_strategy == "aggressive":
            for file_path in REPORT_DIR.glob("*.xlsx"):
                try:
                    if file_path.is_file():
                        file_path.unlink()
                        deleted += 1
                        logger.info(f"✓ Deleted: {file_path.name}")
                except FileNotFoundError:
                    pass
                except Exception as file_error:
                    logger.warning(f"⚠️ Skipped '{file_path.name}': {file_error}")
        
        if deleted > 0:
            logger.info(f"🧹 Cleanup completed — {deleted} file(s) removed (strategy: {cleanup_strategy})")
        else:
            logger.info(f"🧹 Cleanup completed but no files were deleted")
    except Exception as e:
        logger.error(f"❌ Cleanup failed with error: {e}", exc_info=True)

# ============================================================
# SECTION 12 — CENTRALIZED REPORT GENERATION (UPDATED WITH ADVANCED COST LOGGING)
# ============================================================
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
    _start_time = time.perf_counter()
    username_to_use = resolve_username(username, user_email)
    tid = clean_table_id(table_id)
    short_name = extract_short_table_name(tid)
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(short_name, tid)

    if not username_to_use:
        insert_ai_log("", bq_table_name, "", "FAIL")
        return {"success": False, "message": "❌ ไม่พบข้อมูลยืนยันตัวตนในระบบ"}

    if "authenbymenu" in short_name:
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
        # รับค่า actual_sql ออกมาจากตัวแปรลำดับที่ 5
        row_count, file_name, bytes_billed, file_size_mb, actual_sql, ai_log_id = (
            fetch_and_generate_excel(
                tid, report_name, limit,
                filter_column=filter_column,
                filter_value=filter_value,
                filters=filters,
                columns=columns,
            )
        )

        _elapsed = time.perf_counter() - _start_time
        cost = calculate_total_cost(
            bytes_billed=bytes_billed,
            duration_seconds=_elapsed,
            vcpu=COST_CLOUD_RUN_VCPU,
            memory_gib=COST_CLOUD_RUN_MEM_GIB,
            usd_to_thb=COST_USD_TO_THB,
            bq_price_per_tb_usd=float(os.environ.get("COST_BQ_PRICE_PER_TB", "6.25")),
        )
        cost_breakdown = format_cost_breakdown(cost)

        if not row_count:
            insert_ai_log(
                username_to_use, bq_table_name, report_name, "FAIL",
                condition=actual_sql, bytes_billed=bytes_billed, ai_log_id=ai_log_id,
                cost_thb=cost.grand_total_thb, cost_usd=cost.grand_total_usd, cost_summary=cost.summary
            )
            return {"success": False, "message": f"❌ ไม่พบข้อมูลในรายงาน {report_name}"}

        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"

        # ส่งข้อมูลลงคอลัมน์ใหม่อย่างครบถ้วน และใช้ Condition เก็บ SQL แท้แทน
        insert_ai_log(
            username=username_to_use,
            table_name=bq_table_name,
            report_name=report_name,
            status="OK",
            condition=actual_sql, 
            size=file_size_mb,
            bytes_billed=bytes_billed,
            url=download_url,
            row_generated=row_count or 0,
            ai_log_id=ai_log_id,
            cost_thb=cost.grand_total_thb,
            cost_usd=cost.grand_total_usd,
            cost_summary=cost.summary
        )
        return {
            "success": True,
            "message": f"✅ จัดเตรียมรายงาน{report_name} ({row_count} แถว) สำเร็จ",
            "download_url": download_url,
            "file_name": file_name,
            "row_count": row_count,
            "file_size_mb": file_size_mb,
            "report_name": report_name,
            "cost": cost_breakdown,
        }

    except PermissionError as pe:
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": f"🔒 {str(pe)}"}
    except Exception as e:
        logger.error(f"Execution Error: {e}", exc_info=True)
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": "🚨 พบปัญหาการสร้าง Excel"}

# ============================================================
# SECTION 13 — MCP TOOLS
# ============================================================
@mcp.tool(name="sc_report_export", description="ดึงข้อมูล SC Report และสร้างไฟล์ Excel")
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

    res = _execute_report_generation(
        table_id=table_id, limit=limit, filter_column=filter_column,
        filter_value=filter_value, filters=parsed_filters, columns=columns, username=username,
    )
    if res.get("success") and "cost" in res:
        res["cost_summary"] = res["cost"].get("summary", "")
    return res

@mcp.tool(name="generate_excel_report", description="สร้างไฟล์ Excel จากรายงานใน BigQuery")
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
        table_id=table_id, limit=limit, filter_column=filter_column,
        filter_value=filter_value, filters=parsed_filters, columns=columns,
        username=username, user_email=user_email,
    )
    if res["success"]:
        cost_line = ""
        if "cost" in res:
            c = res["cost"]
            cost_line = (
                f"\n💰 ค่าใช้จ่าย: BQ ${c['bigquery']['cost_usd']:.6f} | "
                f"Cloud Run ${c['cloud_run']['total_cost_usd']:.6f} | "
                f"รวม ฿{c['grand_total_thb']:.4f}"
            )
        return (
            f"📊 จัดเตรียม **รายงาน{res['report_name']}** ({res['row_count']} แถว) เรียบร้อยแล้วครับ\n"
            f"🔗 {res['download_url']}{cost_line}"
        )
    return res["message"]

# ============================================================
# SECTION 14 — REST API ENDPOINTS
# ============================================================
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest):
    res = _execute_report_generation(
        table_id=request.table_id, limit=request.limit, filter_column=request.filter_column,
        filter_value=request.filter_value, filters=request.filters, columns=request.columns,
        username=request.username, user_email=request.user_email,
    )
    if res["success"]:
        return {
            "success": True,
            "message": f"✅ จัดเตรียม **รายงาน{res['report_name']}** สำเร็จ\n🔗 {res['download_url']}",
            "cost": res.get("cost"),
        }
    return {"success": False, "message": res["message"]}

@app.post("/cost_estimate")
def cost_estimate(req: CostEstimateRequest):
    cost = calculate_total_cost(
        bytes_billed=req.bytes_billed, duration_seconds=req.duration_seconds,
        vcpu=req.vcpu, memory_gib=req.memory_gib, usd_to_thb=req.usd_to_thb,
        bq_price_per_tb_usd=float(os.environ.get("COST_BQ_PRICE_PER_TB", "6.25")),
    )
    return format_cost_breakdown(cost)

@app.get("/download/{file_name}")
def download_report(file_name: str, direct: bool = False):
    if not re.match(r"^[A-Za-z0-9_]+_[a-f0-9]{8}_\d{8}_\d{6}\.xlsx$", file_name):
        return {"success": False, "message": "❌ รูปแบบลิงก์ไม่ถูกต้อง"}

    file_path = (REPORT_DIR / file_name).resolve()
    if not file_path.is_relative_to(REPORT_DIR) or not file_path.exists():
        return {"success": False, "message": "❌ ไม่พบไฟล์หรือไม่มีสิทธิ์เข้าถึง"}

    if direct:
        return FileResponse(
            path=file_path, filename=file_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    html_content = (
        '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Downloading...</title></head><body><script>'
        'const downloadUrl = window.location.pathname + "?direct=true";'
        "const iframe = document.createElement('iframe');iframe.style.display = 'none';iframe.src = downloadUrl;"
        "document.body.appendChild(iframe);setTimeout(() => { window.close(); }, 1000);</script></body></html>"
    )
    return HTMLResponse(content=html_content)

@app.get("/health")
def health_check():
    return {"status": "healthy"}

@app.get("/files")
def list_files():
    try:
        files = []
        for file_path in REPORT_DIR.glob("*"):
            if file_path.is_file():
                stat = file_path.stat()
                files.append({
                    "name": file_path.name, "size_bytes": stat.st_size,
                    "size_mb": round(stat.st_size / (1024 * 1024), 4),
                    "created_at": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                    "modified_at": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                })
        return {"directory": str(REPORT_DIR), "files": files}
    except Exception as e:
        logger.error(f"Error listing files: {e}")
        return {"success": False, "message": "Failed to list files"}

@app.post("/internal/cleanup")
async def trigger_cleanup(x_cleanup_token: str = Header(None)):
    if x_cleanup_token != CLEANUP_SECRET:
        logger.warning(f"🚨 Unauthorized cleanup attempt with token: {x_cleanup_token}")
        raise HTTPException(status_code=401, detail="Unauthorized")
    logger.info("🧹 Triggering manual report cleanup via API...")
    await asyncio.to_thread(cleanup_old_reports)
    return {"success": True, "message": "🧹 Cleanup completed successfully"}

# ============================================================
# SECTION 15 — MOUNT MCP & ENTRYPOINT
# ============================================================
app.mount("", mcp_asgi_app)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)