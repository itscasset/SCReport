import os
import re
import uuid
import logging
import time
import asyncio
import math
from dataclasses import dataclass
import pandas as pd
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
from graph_mcp import run_visualization_code

# ============================================================
# SECTION 1 — COST CALCULATOR
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
    usd_to_thb: float = 32.57,
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
    usd_to_thb: float = 32.57,
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
    usd_to_thb: float = 32.57,
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

COST_USD_TO_THB = float(os.environ.get("COST_USD_TO_THB", "32.57"))
COST_CLOUD_RUN_VCPU = float(os.environ.get("COST_CLOUD_RUN_VCPU", "1.0"))
COST_CLOUD_RUN_MEM_GIB = float(os.environ.get("COST_CLOUD_RUN_MEM_GIB", "0.5"))

CLEANUP_SECRET = os.environ.get("CLEANUP_SECRET")

LIBRECHAT_JWT_SECRET = os.environ.get("LIBRECHAT_JWT_SECRET", "")

# ============================================================
# SECTION 3 — MAPPING & FIRESTORE
# ============================================================
from google.cloud import firestore
from mapping import (
    TABLE_TO_REPORT_NAMES,
    TABLE_TO_BQ_TABLE_NAME,
    ALLOWED_OPERATORS,
    USER_REGEX,
    BQ_TABLE_REGEX,
    SENSITIVE_COLUMNS,
)

REPORT_DIR = Path("/tmp/reports").resolve()
REPORT_DIR.mkdir(parents=True, exist_ok=True)

bq_client = bigquery.Client(project=PROJECT_ID)
firestore_client = firestore.Client(project=PROJECT_ID)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("SCReportSecurity")

_MAPPING_CACHE: dict[str, tuple[dict[str, str], datetime]] = {}
CACHE_TTL = timedelta(minutes=5)

_USERNAME_CACHE: dict[str, tuple[Optional[str], datetime]] = {}
_USERNAME_CACHE_LOCK = Lock()
USERNAME_CACHE_TTL = timedelta(minutes=10)

UNRESOLVED_TEMPLATE_REGEX = re.compile(
    r'(\$\{[^}]+\}|\{\{[^}]+\}\}|\{[^}]+\}|\[.*?MASKED.*?\])',
    re.IGNORECASE
)

def is_valid_email(value: str) -> bool:
    """ตรวจว่าค่าเป็น email จริงๆ — ไม่รับ template variable หรือ placeholder ใดๆ"""
    if not value:
        return False
    val_strip = value.strip()
    # กรอง template variable ที่ยังไม่ถูก resolve
    if UNRESOLVED_TEMPLATE_REGEX.search(val_strip):
        return False
    return bool(re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', val_strip))

def get_table_column_mapping(table_id: str) -> dict[str, str]:
    cleaned_table = clean_table_id(table_id)
    short_name = extract_short_table_name(cleaned_table)
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(short_name, cleaned_table)

    now = datetime.now(timezone.utc)
    if bq_table_name in _MAPPING_CACHE:
        mapping, expiry = _MAPPING_CACHE[bq_table_name]
        if now < expiry:
            return mapping

    try:
        doc_ref = (
            firestore_client.collection("data_dictionary")
            .document(DATASET_ID)
            .collection("tables")
            .document(bq_table_name)
        )
        doc = doc_ref.get()
        mapping = {}
        if doc.exists:
            data = doc.to_dict() or {}
            columns_list = data.get("columns", [])
            for col in columns_list:
                name = col.get("name")
                desc = col.get("description")
                if name:
                    mapping[name] = desc if desc else name
            logger.info(f"[Firestore] Loaded mapping for '{bq_table_name}' with {len(mapping)} columns.")
        else:
            logger.warning(f"[Firestore] Document data_dictionary/{DATASET_ID}/tables/{bq_table_name} does not exist.")

        _MAPPING_CACHE[bq_table_name] = (mapping, now + CACHE_TTL)
        return mapping
    except Exception as e:
        logger.error(f"Error fetching column mapping from Firestore: {e}", exc_info=True)
        return {}

CURRENT_REQUEST_USERNAME: ContextVar[Optional[str]] = ContextVar(
    "current_request_username", default=None
)
CURRENT_REQUEST_BQ_USERNAME: ContextVar[Optional[str]] = ContextVar(
    "current_request_bq_username", default=None
)
CURRENT_REQUEST_BASE_URL: ContextVar[Optional[str]] = ContextVar(
    "current_request_base_url", default=None
)

# ============================================================
# SECTION 4 — LOG ID GENERATOR
# ============================================================
SAFE_COLUMN_REGEX = re.compile(r"^[A-Za-z0-9_]{1,128}$")

def sanitize_column_name(column: str) -> str:
    if not column or not SAFE_COLUMN_REGEX.match(column):
        raise PermissionError(f"คอลัมน์ '{column}' มีรูปแบบไม่ถูกต้องหรือไม่ปลอดภัย")
    return column

def generate_ai_log_id() -> int:
    return uuid.uuid4().int & 0x7FFFFFFFFFFFFFFF

# ============================================================
# SECTION 5 — JWT EMAIL EXTRACTOR
# ============================================================
def extract_email_from_jwt(token: str) -> Optional[str]:
    if not token:
        return None
    try:
        import jwt as pyjwt

        if LIBRECHAT_JWT_SECRET:
            payload = pyjwt.decode(
                token,
                LIBRECHAT_JWT_SECRET,
                algorithms=["HS256", "HS384", "HS512"],
            )
        else:
            payload = pyjwt.decode(
                token,
                options={"verify_signature": False},
                algorithms=["HS256", "HS384", "HS512", "RS256"],
            )

        email = (
            payload.get("email")
            or payload.get("user_email")
            or payload.get("preferred_username")
        )

        if not email:
            sub = payload.get("sub", "")
            if sub and "@" in sub:
                email = sub

        if email and is_valid_email(email):
            logger.info(f"[JWT] Extracted email from token: {mask_username(email)}")
            return email.strip().lower()
        else:
            logger.warning(f"[JWT] Token decoded but no valid email found. Keys: {list(payload.keys())}")
            return None

    except Exception as e:
        logger.warning(f"[JWT] Failed to decode token: {type(e).__name__}: {e}")
        return None


def lookup_username_from_email(email: str) -> Optional[str]:
    if not is_valid_email(email):
        logger.warning(f"[lookup_username_from_email] Invalid/unresolved email skipped: '{email}'")
        return None

    email_lower = email.strip().lower()
    now = datetime.now(timezone.utc)

    with _USERNAME_CACHE_LOCK:
        if email_lower in _USERNAME_CACHE:
            cached_username, expiry = _USERNAME_CACHE[email_lower]
            if now < expiry:
                logger.info(f"[UsernameCache] HIT for {mask_username(email)}: {cached_username}")
                return cached_username

    try:
        query = f"""
            SELECT UserName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE LOWER(Email) = @email
            LIMIT 1
        """
        job_config = bigquery.QueryJobConfig(
            query_parameters=[
                bigquery.ScalarQueryParameter("email", "STRING", email_lower)
            ]
        )
        results = list(bq_client.query(query, job_config=job_config).result())
        username = results[0]["UserName"] if results else None

        with _USERNAME_CACHE_LOCK:
            _USERNAME_CACHE[email_lower] = (username, now + USERNAME_CACHE_TTL)

        if username:
            logger.info(f"[UsernameCache] MISS → resolved '{mask_username(email)}' to '{username}'")
        else:
            logger.warning(f"[UsernameCache] MISS → no UserName found for '{mask_username(email)}'")
        return username

    except Exception as e:
        logger.error(f"[lookup_username_from_email] BQ error: {e}", exc_info=True)
        return None

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
            logger.info("MCP Session started")
            yield
    finally:
        logger.info("App shutting down")

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
async def capture_username_header(request: Request, call_next):
    """
    ลำดับการหา email:
      1. Header x-user-email / current-user / current_user
         → ตรวจว่าเป็น email จริง (ไม่ใช่ template variable)
      2. JWT จาก Authorization: Bearer <token>
         → decode เพื่อดึง email field
      3. ไม่พบ → username_val = None (anonymous)
    """
    logger.info(f"🔍 ALL HEADERS: {dict(request.headers)}")

    body_bytes = await request.body()
    logger.info(f"🔍 BODY PREVIEW: {body_bytes[:500]}")

    async def receive():
        return {"type": "http.request", "body": body_bytes}
    request = Request(request.scope, receive)

    username_val: Optional[str] = None
    email_source = "none"

    # ── 1. ลอง header ที่ส่งมาตรงๆ ──────────────────────────────────────────
    for header_name in ["current-user", "current_user", "x-user-email"]:
        val = request.headers.get(header_name, "").strip()
        if val:
            if is_valid_email(val):
                username_val = val.lower()
                email_source = f"header:{header_name}"
                break
            else:
                logger.warning(
                    f"⚠️ [Middleware] Header '{header_name}' has unresolved/invalid value: '{val}' — skipping"
                )

    # ── 2. Fallback: decode JWT จาก Authorization header ────────────────────
    if not username_val:
        auth_header = request.headers.get("authorization", "").strip()
        if auth_header.lower().startswith("bearer "):
            token = auth_header[7:].strip()
            jwt_email = await asyncio.to_thread(extract_email_from_jwt, token)
            if jwt_email:
                username_val = jwt_email
                email_source = "jwt"

    show_raw = os.getenv("SHOW_RAW_EMAIL", "false").lower() == "true"
    masked_user = username_val if show_raw else mask_username(username_val)
    logger.info(
        f"📨 [Middleware] {request.method} {request.url.path} | "
        f"User: {masked_user} | Source: {email_source}"
    )

    token_username = CURRENT_REQUEST_USERNAME.set(username_val)

    # ── 3. Resolve BQ username จาก email ────────────────────────────────────
    bq_username: Optional[str] = None
    if username_val:
        bq_username = await asyncio.to_thread(lookup_username_from_email, username_val)
        if bq_username:
            logger.info(f"✅ [Middleware] Resolved BQ username: '{bq_username}' for {masked_user}")
        else:
            logger.warning(f"⚠️ [Middleware] No BQ username found for: {masked_user}")

    token_bq_username = CURRENT_REQUEST_BQ_USERNAME.set(bq_username)
    base_url = str(request.base_url).rstrip("/")
    token_base = CURRENT_REQUEST_BASE_URL.set(base_url)

    try:
        response = await call_next(request)
        return response
    finally:
        CURRENT_REQUEST_USERNAME.reset(token_username)
        CURRENT_REQUEST_BQ_USERNAME.reset(token_bq_username)
        CURRENT_REQUEST_BASE_URL.reset(token_base)


@app.middleware("http")
async def block_sse_stream(request: Request, call_next):
    if request.method == "GET" and request.url.path.rstrip("/") in ("/mcp", "/mcp/sse"):
        from starlette.responses import JSONResponse
        return JSONResponse(
            status_code=405,
            content={"detail": "SSE transport is disabled. Use POST /mcp for Streamable HTTP."},
            headers={"Retry-After": "3600"},
        )
    return await call_next(request)


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
    usd_to_thb: float = Field(default=32.57, gt=0, description="อัตราแลกเปลี่ยน USD→THB")

# ============================================================
# SECTION 9 — HELPERS & SECURITY
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
    """
    Priority order (PATCHED — ปิด fallback จาก AI เพื่อป้องกัน identity spoofing):
      1. BQ username ที่ middleware resolve ไว้แล้ว (เท่านั้นสำหรับ MCP path)
      2. lookup จาก fallback_email (REST API โดยตรง — ไม่ผ่าน MCP)
      3. "" — caller ต้อง reject เอง
    หมายเหตุ: fallback_user จาก AI ถูกตัดออกแล้ว
              ใครส่ง username มาใน tool call จะถูกเพิกเฉย
    """
    # 1. BQ username จาก middleware (แม่นยำและปลอดภัยที่สุด)
    bq_username = CURRENT_REQUEST_BQ_USERNAME.get()
    if bq_username and is_valid_user(bq_username):
        return bq_username.strip()

    # 2. lookup จาก fallback_email (REST API path เท่านั้น)
    if fallback_email and is_valid_email(fallback_email):
        db_username = lookup_username_from_email(fallback_email)
        if db_username and is_valid_user(db_username):
            return db_username.strip()

    # 3. ไม่พบ — return "" ให้ caller reject
    logger.warning(
        f"[resolve_username] Could not resolve identity — "
        f"middleware_email={repr(CURRENT_REQUEST_USERNAME.get())}, "
        f"bq_username={repr(CURRENT_REQUEST_BQ_USERNAME.get())}, "
        f"fallback_email={repr(fallback_email)}"
    )
    return ""

def check_user_permission(username: str, table_id: str) -> Optional[str]:
    if not is_valid_user(username):
        logger.warning(f"Security Blocked: Invalid username format '{repr(username)}'")
        return None

    cleaned_table = clean_table_id(table_id)
    if not validate_table_id(cleaned_table):
        logger.warning(f"Security Blocked: Invalid table_id format '{repr(table_id)}'")
        return None

    short_name = extract_short_table_name(cleaned_table)
    masked_user = mask_username(username)
    logger.info(f"[Auth] table_id='{table_id}' -> short_name='{short_name}' | user='{masked_user}'")

    if "authenbymenu" in short_name:
        logger.warning(f"Security Blocked: User '{masked_user}' tried to access permission table directly")
        return None

    report_names = TABLE_TO_REPORT_NAMES.get(short_name)
    if not report_names:
        logger.info(f"Access Denied: Table '{short_name}' is not mapped to any report.")
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

        logger.warning(f"Access Denied: User '{mask_username(clean_username)}' has no permission for table '{short_name}'")
        return None
    except Exception as e:
        logger.error(f"Auth verification database error: {e}")
        return None

def validate_filter_column(table_id: str, column: str) -> bool:
    if not table_id or not column:
        return False
    mapping_dict = get_table_column_mapping(table_id)
    target_column = column.lower()
    return any(k.lower() == target_column for k in mapping_dict.keys())

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
                "Condition": (condition or "")[:2000],
                "Size": round(size, 4),
                "BytesBilled": bytes_billed,
                "URL": (url or "")[:1000],
                "Row_generated": row_generated,
                "AiLogID": ai_log_id,
                "CostTHB": round(cost_thb, 6),
                "CostUSD": round(cost_usd, 8),
                "CostSummary": (cost_summary or "")[:500],
            }
        ]
        full_table = f"{PROJECT_ID}.{DATASET_ID}.{LOG_TABLE}"
        errors = bq_client.insert_rows_json(full_table, rows_to_insert)

        if errors:
            logger.error(f"AiLog insert FAILED: {errors}")
    except Exception as e:
        logger.error(f"Exception in insert_ai_log: {e}", exc_info=True)

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
# SECTION 10 — FETCH & GENERATE EXCEL
# ============================================================
def fetch_and_generate_excel(
    table_id: str,
    report_name: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
    columns: Optional[list[str]] = None,
    condition: Optional[str] = "AND",
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
        logger.error(f"[Schema] Failed to fetch schema: {e}")
        schema_dict = {}

    if columns:
        for col in columns:
            if not validate_filter_column(short_name, col):
                raise PermissionError(f"คอลัมน์ '{col}' ไม่ได้รับอนุญาต")
            sanitize_column_name(col)
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
            sanitize_column_name(f.column)
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
        sanitize_column_name(filter_column)
        p_type, p_val = map_param_type_and_value(filter_column, filter_value, schema_dict)
        where_clauses.append(f"`{filter_column}` = @filter_value")
        params.append(bigquery.ScalarQueryParameter("filter_value", p_type, p_val))

    if where_clauses:
        joiner = " AND "
        if condition and condition.upper() in ("AND", "OR"):
            joiner = f" {condition.upper()} "
        query += " WHERE " + joiner.join(where_clauses)

    safe_limit = min(max(limit, 1), 100000)
    query += f" LIMIT {safe_limit}"

    actual_sql_executed = query

    job_config = bigquery.QueryJobConfig(query_parameters=params)
    job = bq_client.query(query, job_config=job_config)
    result = job.result()

    bytes_billed: int = job.total_bytes_billed
    schema = result.schema
    row_count = result.total_rows

    if not row_count:
        ai_log_id = generate_ai_log_id()
        return 0, None, bytes_billed, 0.0, actual_sql_executed, ai_log_id

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Report")

    raw_mapping = get_table_column_mapping(table_id)
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

    logger.info(f"Cleanup: strategy={cleanup_strategy}")
    if not REPORT_DIR.exists():
        return

    current_time = time.time()
    deleted = 0

    if cleanup_strategy not in ["age", "keep_latest", "aggressive"]:
        logger.warning(f"Unknown strategy '{cleanup_strategy}'. Fallback to 'age'.")
        cleanup_strategy = "age"

    if cleanup_strategy == "age":
        cutoff = current_time - cleanup_age_seconds
        files_list = list(REPORT_DIR.glob("*.xlsx")) + list(REPORT_DIR.glob("*.png"))
        for file_path in files_list:
            if file_path.is_file() and file_path.stat().st_mtime < cutoff:
                file_path.unlink()
                deleted += 1

    elif cleanup_strategy == "keep_latest":
        files = list(REPORT_DIR.glob("*.xlsx")) + list(REPORT_DIR.glob("*.png"))
        sorted_files = sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)
        for file_path in sorted_files[cleanup_max_files:]:
            if file_path.is_file():
                file_path.unlink()
                deleted += 1

    elif cleanup_strategy == "aggressive":
        files_list = list(REPORT_DIR.glob("*.xlsx")) + list(REPORT_DIR.glob("*.png"))
        for file_path in files_list:
            if file_path.is_file():
                file_path.unlink()
                deleted += 1

    logger.info(f"Cleanup completed — {deleted} file(s) removed.")

# ============================================================
# SECTION 12 — CENTRALIZED REPORT GENERATION (PATCHED)
# ============================================================
def _execute_report_generation(
    table_id: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
    columns: Optional[list[str]] = None,
    username: Optional[str] = None,      # รับไว้แต่ไม่ใช้ — ป้องกัน breaking change
    user_email: Optional[str] = None,
    condition: Optional[str] = "AND",
) -> dict:
    _start_time = time.perf_counter()

    tid = clean_table_id(table_id)
    short_name = extract_short_table_name(tid)
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(short_name, tid)

    # ── Identity: ต้องมาจาก middleware เท่านั้น ──────────────────────────────
    # fallback_user=None — ไม่รับ username จาก AI เด็ดขาด
    # fallback_email=user_email — รับเฉพาะ REST API path ที่ส่ง email จริงมา
    username_to_use = resolve_username(
        fallback_user=None,
        fallback_email=user_email,
    )

    if not username_to_use:
        logger.warning(
            f"[Auth] REJECTED — identity unresolved | "
            f"middleware_email={repr(CURRENT_REQUEST_USERNAME.get())} | "
            f"bq_username={repr(CURRENT_REQUEST_BQ_USERNAME.get())} | "
            f"ai_username={repr(username)}"  # log ค่าที่ AI พยายามส่งมา
        )
        insert_ai_log("UNAUTH", bq_table_name, "", "FAIL_UNAUTH")
        return {
            "success": False,
            "message": "ไม่สามารถยืนยันตัวตนได้ กรุณา Login ใหม่อีกครั้ง"
        }
    # ─────────────────────────────────────────────────────────────────────────

    if "authenbymenu" in short_name:
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return {"success": False, "message": "[Access Denied] ระบบไม่อนุญาตให้เข้าถึงตารางสิทธิ์"}

    if not validate_table_id(tid):
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return {"success": False, "message": f"table_id '{tid}' ไม่ถูกต้อง"}

    report_name = check_user_permission(username_to_use, tid)
    if not report_name:
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return {"success": False, "message": "ขออภัย คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    try:
        row_count, file_name, bytes_billed, file_size_mb, actual_sql, ai_log_id = (
            fetch_and_generate_excel(
                tid, report_name, limit,
                filter_column=filter_column,
                filter_value=filter_value,
                filters=filters,
                columns=columns,
                condition=condition,
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
                cost_thb=cost.grand_total_thb, cost_usd=cost.grand_total_usd,
                cost_summary=cost.summary,
            )
            return {"success": False, "message": f"ไม่พบข้อมูลในรายงาน {report_name}"}

        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"

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
            cost_summary=cost.summary,
        )
        return {
            "success": True,
            "message": f"จัดเตรียมรายงาน{report_name} ({row_count} แถว) สำเร็จ",
            "download_url": download_url,
            "file_name": file_name,
            "row_count": row_count,
            "file_size_mb": file_size_mb,
            "report_name": report_name,
            "cost": cost_breakdown,
        }

    except PermissionError as pe:
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": str(pe)}
    except Exception as e:
        logger.error(f"Execution Error: {e}", exc_info=True)
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": "พบปัญหาการสร้าง Excel"}

def fetch_dataframe_for_chart(
    table_id: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
    columns: Optional[list[str]] = None,
    condition: Optional[str] = "AND",
) -> tuple[pd.DataFrame, int, str]:
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
        logger.error(f"[Schema] Failed to fetch schema: {e}")
        schema_dict = {}

    if columns:
        for col in columns:
            if not validate_filter_column(short_name, col):
                raise PermissionError(f"คอลัมน์ '{col}' ไม่ได้รับอนุญาต")
            sanitize_column_name(col)
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
            sanitize_column_name(f.column)
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
        sanitize_column_name(filter_column)
        p_type, p_val = map_param_type_and_value(filter_column, filter_value, schema_dict)
        where_clauses.append(f"`{filter_column}` = @filter_value")
        params.append(bigquery.ScalarQueryParameter("filter_value", p_type, p_val))

    if where_clauses:
        joiner = " AND "
        if condition and condition.upper() in ("AND", "OR"):
            joiner = f" {condition.upper()} "
        query += " WHERE " + joiner.join(where_clauses)

    safe_limit = min(max(limit, 1), 100000)
    query += f" LIMIT {safe_limit}"

    actual_sql_executed = query

    job_config = bigquery.QueryJobConfig(query_parameters=params)
    job = bq_client.query(query, job_config=job_config)
    df = job.to_dataframe()

    bytes_billed: int = job.total_bytes_billed
    return df, bytes_billed, actual_sql_executed

def _execute_chart_generation(
    table_id: str,
    code: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
    columns: Optional[list[str]] = None,
    username: Optional[str] = None,      # รับไว้แต่ไม่ใช้ — ป้องกัน breaking change
    user_email: Optional[str] = None,
    condition: Optional[str] = "AND",
) -> dict:
    _start_time = time.perf_counter()

    tid = clean_table_id(table_id)
    short_name = extract_short_table_name(tid)
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(short_name, tid)

    # ── Identity: ต้องมาจาก middleware เท่านั้น ──────────────────────────────
    username_to_use = resolve_username(
        fallback_user=None,
        fallback_email=user_email,
    )

    if not username_to_use:
        logger.warning(
            f"[Auth] REJECTED — identity unresolved | "
            f"middleware_email={repr(CURRENT_REQUEST_USERNAME.get())} | "
            f"bq_username={repr(CURRENT_REQUEST_BQ_USERNAME.get())} | "
            f"ai_username={repr(username)}"
        )
        insert_ai_log("UNAUTH", bq_table_name, "", "FAIL_UNAUTH")
        return {
            "success": False,
            "message": "ไม่สามารถยืนยันตัวตนได้ กรุณา Login ใหม่อีกครั้ง"
        }
    # ─────────────────────────────────────────────────────────────────────────

    if "authenbymenu" in short_name:
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL_CHART")
        return {"success": False, "message": "[Access Denied] ระบบไม่อนุญาตให้เข้าถึงตารางสิทธิ์"}

    if not validate_table_id(tid):
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL_CHART")
        return {"success": False, "message": f"table_id '{tid}' ไม่ถูกต้อง"}

    report_name = check_user_permission(username_to_use, tid)
    if not report_name:
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL_CHART")
        return {"success": False, "message": "ขออภัย คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    try:
        df, bytes_billed, actual_sql = fetch_dataframe_for_chart(
            tid, limit, filter_column, filter_value, filters, columns, condition
        )

        if df.empty:
            ai_log_id = generate_ai_log_id()
            _elapsed = time.perf_counter() - _start_time
            cost = calculate_total_cost(bytes_billed, _elapsed, COST_CLOUD_RUN_VCPU, COST_CLOUD_RUN_MEM_GIB, COST_USD_TO_THB)
            insert_ai_log(
                username_to_use, bq_table_name, report_name, "FAIL_CHART",
                condition=actual_sql, bytes_billed=bytes_billed, ai_log_id=ai_log_id,
                cost_thb=cost.grand_total_thb, cost_usd=cost.grand_total_usd, cost_summary=cost.summary
            )
            return {"success": False, "message": f"ไม่พบข้อมูลในรายงาน {report_name} เพื่อใช้วาดกราฟ"}

        raw_mapping = get_table_column_mapping(tid)
        lowercase_mapping = {k.lower(): v for k, v in raw_mapping.items()}
        rename_dict = {}
        for col in df.columns:
            translated = lowercase_mapping.get(col.lower(), col)
            rename_dict[col] = translated
        df = df.rename(columns=rename_dict)

        ai_log_id = generate_ai_log_id()
        secure_file_id = uuid.uuid4().hex[:8]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"chart_{bq_table_name}_{secure_file_id}_{timestamp}.png"
        file_path = REPORT_DIR / file_name

        res = run_visualization_code(df, code, str(file_path))

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

        if not res["success"]:
            insert_ai_log(
                username_to_use, bq_table_name, report_name, "FAIL_CHART",
                condition=actual_sql, bytes_billed=bytes_billed, ai_log_id=ai_log_id,
                cost_thb=cost.grand_total_thb, cost_usd=cost.grand_total_usd, cost_summary=cost.summary
            )
            return {"success": False, "message": f"วาดกราฟไม่สำเร็จ: {res['message']}"}

        file_size_mb = round(file_path.stat().st_size / (1024 * 1024), 4)
        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"

        insert_ai_log(
            username=username_to_use,
            table_name=bq_table_name,
            report_name=report_name,
            status="OK_CHART",
            condition=actual_sql,
            size=file_size_mb,
            bytes_billed=bytes_billed,
            url=download_url,
            row_generated=len(df),
            ai_log_id=ai_log_id,
            cost_thb=cost.grand_total_thb,
            cost_usd=cost.grand_total_usd,
            cost_summary=cost.summary
        )

        return {
            "success": True,
            "message": f"จัดเตรียมรายงานแผนภูมิของ{report_name} สำเร็จ",
            "download_url": download_url,
            "file_name": file_name,
            "row_count": len(df),
            "file_size_mb": file_size_mb,
            "report_name": report_name,
            "cost": cost_breakdown,
        }

    except Exception as e:
        logger.error(f"Chart Execution Error: {e}", exc_info=True)
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL_CHART")
        return {"success": False, "message": f"พบปัญหาการดึงข้อมูลเพื่อวาดกราฟ: {str(e)}"}

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
            return {"success": False, "message": "รูปแบบ filters ไม่ถูกต้อง"}

    res = _execute_report_generation(
        table_id=table_id, limit=limit, filter_column=filter_column,
        filter_value=filter_value, filters=parsed_filters, columns=columns,
        username=username,  # ส่งผ่านเพื่อ log แต่ไม่ได้ใช้ใน resolve_username
        condition="AND",
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
            return "รูปแบบ filters ไม่ถูกต้อง"

    res = _execute_report_generation(
        table_id=table_id, limit=limit, filter_column=filter_column,
        filter_value=filter_value, filters=parsed_filters, columns=columns,
        username=username, user_email=user_email, condition=condition or "AND",
    )
    if res["success"]:
        cost_line = ""
        if "cost" in res:
            c = res["cost"]
            cost_line = (
                f"ค่าใช้จ่าย: BQ ${c['bigquery']['cost_usd']:.6f} | "
                f"Cloud Run ${c['cloud_run']['total_cost_usd']:.6f} | "
                f"รวม ฿{c['grand_total_thb']:.4f}"
            )
        return (
            f"จัดเตรียม **รายงาน{res['report_name']}** ({res['row_count']} แถว) เรียบร้อยแล้วครับ\n"
            f"{res['download_url']}\n{cost_line}"
        )
    return res["message"]

@mcp.tool(name="generate_chart_report", description="สร้างกราฟหรือแผนภูมิจากข้อมูลรายงานใน BigQuery ด้วยโค้ด Python (Seaborn/Matplotlib)")
def mcp_generate_chart_report(
    table_id: str,
    code: str,
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
            return "รูปแบบ filters ไม่ถูกต้อง"

    res = _execute_chart_generation(
        table_id=table_id, code=code, limit=limit, filter_column=filter_column,
        filter_value=filter_value, filters=parsed_filters, columns=columns,
        username=username, user_email=user_email, condition=condition or "AND",
    )
    if res["success"]:
        cost_line = ""
        if "cost" in res:
            c = res["cost"]
            cost_line = (
                f"ค่าใช้จ่าย: BQ ${c['bigquery']['cost_usd']:.6f} | "
                f"Cloud Run ${c['cloud_run']['total_cost_usd']:.6f} | "
                f"รวม ฿{c['grand_total_thb']:.4f}"
            )
        return (
            f"จัดเตรียม **แผนภูมิของรายงาน{res['report_name']}** ({res['row_count']} แถว) เรียบร้อยแล้วครับ\n"
            f"{res['download_url']}\n{cost_line}"
        )
    return res["message"]

@mcp.tool(name="check_accessible_reports", description="ตรวจสอบว่าผู้ใช้งานคนปัจจุบันมีสิทธิ์เข้าถึงรายงาน (ReportName / Table) ใดบ้าง ไม่ต้องใส่พารามิเตอร์ใดๆ ระบบจะดึงจากอีเมลผู้ใช้อัตโนมัติจาก Header")
def mcp_check_accessible_reports() -> dict:
    username = resolve_username(None)
    if not username or username == "Anonymous" or username == "Unknown":
        return {"success": False, "message": "ไม่พบข้อมูลสิทธิ์ของผู้ใช้งาน หรือ ไม่สามารถระบุตัวตนผู้ใช้ได้"}
        
    try:
        query = f"""
            SELECT ReportName 
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE UserName = @username
        """
        job_config = bigquery.QueryJobConfig(
            query_parameters=[
                bigquery.ScalarQueryParameter("username", "STRING", username)
            ]
        )
        results = list(bq_client.query(query, job_config=job_config).result())
        reports = [row["ReportName"] for row in results]
        
        return {
            "success": True,
            "username": username,
            "accessible_reports": reports,
            "total": len(reports)
        }
    except Exception as e:
        logger.error(f"Error fetching accessible reports: {e}")
        return {"success": False, "message": f"เกิดข้อผิดพลาดในการดึงข้อมูลสิทธิ์: {str(e)}"}

# ============================================================
# SECTION 14 — REST API ENDPOINTS
# ============================================================
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest):
    res = _execute_report_generation(
        table_id=request.table_id, limit=request.limit, filter_column=request.filter_column,
        filter_value=request.filter_value, filters=request.filters, columns=request.columns,
        username=request.username, user_email=request.user_email, condition=request.condition or "AND",
    )
    if res["success"]:
        return {
            "success": True,
            "message": f"จัดเตรียม **รายงาน{res['report_name']}** สำเร็จ\n{res['download_url']}",
            "cost": res.get("cost"),
        }
    return {"success": False, "message": res["message"]}

class GenerateChartRequest(BaseModel):
    model_config = ConfigDict(extra="allow")
    table_id: str = Field(..., min_length=1, description="ชื่อตารางใน BigQuery")
    code: str = Field(..., min_length=1, description="โค้ด Python สำหรับพลอตกราฟ")
    limit: int = Field(default=2000, gt=0, le=100000, description="จำนวนแถวสูงสุดที่ดึง")
    filters: Optional[list[FilterCondition]] = Field(default=None, description="รายการฟิลเตอร์")
    condition: Optional[Literal["AND", "OR"]] = Field(default="AND", description="เงื่อนไขการเชื่อมฟิลเตอร์")
    columns: Optional[list[str]] = Field(default=None, description="รายชื่อคอลัมน์ที่ต้องการเลือก")
    username: Optional[str] = Field(default=None, description="ชื่อผู้ใช้งานระบบ")
    user_email: Optional[str] = Field(default=None, description="อีเมลผู้ใช้งานระบบ")
    filter_column: Optional[str] = Field(default=None, description="คอลัมน์กรองข้อมูลแบบเดิม")
    filter_value: Optional[str] = Field(default=None, description="ค่ากรองข้อมูลแบบเดิม")

@app.post("/generate_chart_report")
def generate_chart_report(request: GenerateChartRequest):
    res = _execute_chart_generation(
        table_id=request.table_id, code=request.code, limit=request.limit, filter_column=request.filter_column,
        filter_value=request.filter_value, filters=request.filters, columns=request.columns,
        username=request.username, user_email=request.user_email, condition=request.condition or "AND",
    )
    if res["success"]:
        return {
            "success": True,
            "message": f"จัดเตรียม **แผนภูมิของ{res['report_name']}** สำเร็จ\n{res['download_url']}",
            "download_url": res["download_url"],
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
    if not re.match(r"^(chart_)?[A-Za-z0-9_]+_[a-f0-9]{8}_\d{8}_\d{6}\.(xlsx|png)$", file_name):
        raise HTTPException(status_code=400, detail="รูปแบบลิงก์ไม่ถูกต้อง")

    file_path = (REPORT_DIR / file_name).resolve()
    if not file_path.is_relative_to(REPORT_DIR) or not file_path.exists():
        raise HTTPException(status_code=404, detail="ไม่พบไฟล์หรือไม่มีสิทธิ์เข้าถึง")

    if direct:
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if file_name.endswith(".png"):
            media_type = "image/png"
        return FileResponse(path=file_path, filename=file_name, media_type=media_type)

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
@app.get("/file")
def list_files(
    request: Request,
    x_cleanup_token: Optional[str] = Header(None),
    token: Optional[str] = None,
    format: Optional[str] = None
):
    effective_token = x_cleanup_token or token
    if CLEANUP_SECRET and effective_token != CLEANUP_SECRET:
        raise HTTPException(status_code=401, detail="Unauthorized")

    try:
        files = []
        total_size_bytes = 0
        for file_path in REPORT_DIR.glob("*"):
            if file_path.is_file():
                stat = file_path.stat()
                total_size_bytes += stat.st_size
                files.append({
                    "name": file_path.name,
                    "size_bytes": stat.st_size,
                    "size_mb": round(stat.st_size / (1024 * 1024), 4),
                    "created_at": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                    "modified_at": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                })

        files = sorted(files, key=lambda x: x["modified_at"], reverse=True)

        if format == "json" or "text/html" not in request.headers.get("accept", ""):
            return {"directory": str(REPORT_DIR), "files": files}

        token_query = f"&token={token}" if token else ""

        html_template = f"""<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SC Report - File Manager</title>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg-color: #0d1117;
            --card-bg: rgba(22, 27, 34, 0.7);
            --border-color: rgba(48, 54, 61, 0.8);
            --text-main: #f0f6fc;
            --text-muted: #8b949e;
            --accent-blue: #58a6ff;
            --accent-green: #3fb950;
            --accent-orange: #f0883e;
            --accent-red: #ff7b72;
            --glass-blur: blur(12px);
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            background-color: var(--bg-color);
            color: var(--text-main);
            font-family: 'Outfit', 'Sarabun', -apple-system, BlinkMacSystemFont, sans-serif;
            min-height: 100vh;
            background-image: radial-gradient(circle at 10% 20%, rgba(90, 120, 250, 0.05) 0%, transparent 40%),
                              radial-gradient(circle at 90% 80%, rgba(250, 100, 100, 0.03) 0%, transparent 40%);
            background-attachment: fixed;
            padding: 2rem 1.5rem;
        }}
        .container {{ max-width: 1200px; margin: 0 auto; }}
        header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 2rem; flex-wrap: wrap; gap: 1rem; }}
        .title-area h1 {{ font-size: 2rem; font-weight: 700; background: linear-gradient(135deg, #fff 0%, #8b949e 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; display: flex; align-items: center; gap: 0.5rem; }}
        .title-area p {{ color: var(--text-muted); margin-top: 0.25rem; font-size: 0.9rem; }}
        .stats-bar {{ display: flex; gap: 1.5rem; }}
        .stat-card {{ background: var(--card-bg); border: 1px solid var(--border-color); backdrop-filter: var(--glass-blur); border-radius: 12px; padding: 0.75rem 1.25rem; display: flex; flex-direction: column; }}
        .stat-card .value {{ font-size: 1.4rem; font-weight: 600; color: var(--accent-blue); }}
        .stat-card .label {{ font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--text-muted); margin-top: 0.15rem; }}
        .toolbar {{ background: var(--card-bg); border: 1px solid var(--border-color); backdrop-filter: var(--glass-blur); border-radius: 16px; padding: 1rem; margin-bottom: 1.5rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem; }}
        .search-wrapper {{ position: relative; flex: 1; max-width: 400px; }}
        .search-input {{ width: 100%; background: rgba(13, 17, 23, 0.8); border: 1px solid var(--border-color); border-radius: 8px; padding: 0.6rem 1rem 0.6rem 2.5rem; color: var(--text-main); font-family: inherit; font-size: 0.95rem; transition: all 0.2s ease; }}
        .search-input:focus {{ outline: none; border-color: var(--accent-blue); box-shadow: 0 0 0 3px rgba(88, 166, 255, 0.15); }}
        .search-icon {{ position: absolute; left: 0.8rem; top: 50%; transform: translateY(-50%); fill: var(--text-muted); width: 18px; height: 18px; pointer-events: none; }}
        .btn {{ background: var(--accent-blue); color: #0d1117; border: none; border-radius: 8px; padding: 0.6rem 1.2rem; font-weight: 600; font-size: 0.9rem; cursor: pointer; transition: all 0.2s ease; display: inline-flex; align-items: center; gap: 0.5rem; font-family: inherit; }}
        .btn:hover {{ opacity: 0.9; transform: translateY(-1px); }}
        .btn-outline {{ background: transparent; border: 1px solid var(--border-color); color: var(--text-main); }}
        .btn-outline:hover {{ background: rgba(255, 255, 255, 0.05); border-color: var(--text-muted); }}
        .btn-danger {{ background: rgba(255, 123, 114, 0.15); border: 1px solid rgba(255, 123, 114, 0.4); color: var(--accent-red); }}
        .btn-danger:hover {{ background: var(--accent-red); color: #0d1117; border-color: var(--accent-red); }}
        .file-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 1.25rem; }}
        .file-card {{ background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 16px; padding: 1.25rem; display: flex; flex-direction: column; justify-content: space-between; transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1); position: relative; overflow: hidden; }}
        .file-card:hover {{ transform: translateY(-4px); border-color: var(--accent-blue); box-shadow: 0 8px 24px rgba(0, 0, 0, 0.3); }}
        .file-header {{ display: flex; align-items: flex-start; gap: 0.75rem; margin-bottom: 1rem; }}
        .file-icon {{ width: 44px; height: 44px; background: rgba(255, 255, 255, 0.03); border-radius: 10px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; border: 1px solid rgba(255, 255, 255, 0.05); }}
        .icon-xlsx {{ color: var(--accent-green); background: rgba(63, 185, 80, 0.08); border-color: rgba(63, 185, 80, 0.15); }}
        .icon-png {{ color: var(--accent-blue); background: rgba(88, 166, 255, 0.08); border-color: rgba(88, 166, 255, 0.15); }}
        .file-info {{ min-width: 0; }}
        .file-name {{ font-size: 0.95rem; font-weight: 600; margin-bottom: 0.25rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--text-main); }}
        .file-meta {{ font-size: 0.75rem; color: var(--text-muted); display: flex; flex-direction: column; gap: 0.15rem; }}
        .thumbnail-container {{ width: 100%; height: 140px; border-radius: 10px; background: #07090e; overflow: hidden; margin-bottom: 1rem; border: 1px solid rgba(255, 255, 255, 0.05); display: flex; align-items: center; justify-content: center; cursor: zoom-in; }}
        .thumbnail-img {{ max-width: 100%; max-height: 100%; object-fit: contain; transition: transform 0.2s; }}
        .thumbnail-container:hover .thumbnail-img {{ transform: scale(1.05); }}
        .file-actions {{ display: flex; gap: 0.5rem; margin-top: auto; }}
        .action-btn {{ flex: 1; padding: 0.5rem; border-radius: 8px; font-size: 0.85rem; font-weight: 500; text-decoration: none; text-align: center; transition: all 0.2s; cursor: pointer; border: 1px solid transparent; }}
        .btn-view {{ background: rgba(88, 166, 255, 0.1); color: var(--accent-blue); border-color: rgba(88, 166, 255, 0.2); }}
        .btn-view:hover {{ background: var(--accent-blue); color: #0d1117; }}
        .btn-download {{ background: var(--accent-blue); color: #0d1117; font-weight: 600; }}
        .btn-download:hover {{ opacity: 0.9; }}
        .btn-delete {{ background: rgba(255, 123, 114, 0.1); color: var(--accent-red); border-color: rgba(255, 123, 114, 0.2); }}
        .btn-delete:hover {{ background: var(--accent-red); color: #0d1117; border-color: var(--accent-red); }}
        .empty-state {{ background: var(--card-bg); border: 2px dashed var(--border-color); border-radius: 20px; padding: 4rem 2rem; text-align: center; color: var(--text-muted); grid-column: 1 / -1; }}
        .empty-state h3 {{ color: var(--text-main); font-size: 1.3rem; margin-bottom: 0.5rem; }}
        .modal {{ display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(13, 17, 23, 0.9); backdrop-filter: blur(8px); align-items: center; justify-content: center; }}
        .modal-content {{ max-width: 90%; max-height: 85%; box-shadow: 0 24px 48px rgba(0, 0, 0, 0.5); border-radius: 12px; border: 1px solid var(--border-color); background: #0d1117; overflow: hidden; display: flex; flex-direction: column; }}
        .modal-body {{ padding: 1rem; display: flex; justify-content: center; align-items: center; }}
        .modal-body img {{ max-width: 100%; max-height: 70vh; object-fit: contain; }}
        .modal-header {{ display: flex; justify-content: space-between; align-items: center; padding: 1rem 1.5rem; border-bottom: 1px solid var(--border-color); }}
        .modal-title {{ font-size: 1.1rem; font-weight: 600; }}
        .close-btn {{ color: var(--text-muted); font-size: 28px; font-weight: bold; cursor: pointer; transition: color 0.2s; }}
        .close-btn:hover {{ color: var(--text-main); }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="title-area">
                <h1>📁 File Manager</h1>
                <p>จัดการไฟล์เอกสารรายงานและกราฟแดชบอร์ดที่ถูกสร้างขึ้น</p>
            </div>
            <div class="stats-bar">
                <div class="stat-card">
                    <span class="value">{len(files)}</span>
                    <span class="label">จำนวนไฟล์ทั้งหมด</span>
                </div>
                <div class="stat-card">
                    <span class="value">{round(total_size_bytes / (1024 * 1024), 2)} MB</span>
                    <span class="label">ขนาดรวมทั้งหมด</span>
                </div>
            </div>
        </header>
        <div class="toolbar">
            <div class="search-wrapper">
                <svg class="search-icon" viewBox="0 0 16 16"><path d="M10.68 11.74a6 6 0 0 1-8.322-8.322 6 6 0 0 1 8.322 8.322Zm1.06-1.06A7.5 7.5 0 1 0 2.22 2.22a7.5 7.5 0 0 0 10.56 10.56L15 15.28a.749.749 0 0 0 1.06-1.06l-4.32-4.32Z"></path></svg>
                <input type="text" id="searchInput" class="search-input" placeholder="ค้นหาชื่อไฟล์..." onkeyup="filterFiles()">
            </div>
            <div style="display: flex; gap: 0.5rem;">
                <button class="btn btn-outline" onclick="window.location.reload()">🔄 รีเฟรช</button>
                <button class="btn btn-danger" onclick="triggerCleanup()">🧹 เคลียร์ไฟล์ขยะ</button>
            </div>
        </div>
        <div class="file-grid" id="fileGrid">
        """

        if not files:
            html_template += """
            <div class="empty-state">
                <h3>ไม่พบไฟล์ในระบบ</h3>
                <p>ยังไม่มีไฟล์รายงานหรือกราฟบันทึกไว้ในขณะนี้</p>
            </div>
            """
        else:
            for f in files:
                name = f["name"]
                size_mb = f["size_mb"]
                created = datetime.fromisoformat(f["created_at"]).strftime("%Y-%m-%d %H:%M:%S")
                is_img = name.endswith(".png")
                icon_cls = "icon-png" if is_img else "icon-xlsx"
                icon_svg = '<svg viewBox="0 0 16 16" width="20" height="20" fill="currentColor"><path d="M2.002 1a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V3a2 2 0 0 0-2-2h-12zm12 1a1 1 0 0 1 1 1v6.5l-3.777-1.947a.5.5 0 0 0-.577.093l-3.71 3.71-2.66-1.772a.5.5 0 0 0-.63.062L1.002 12V3a1 1 0 0 1 1-1h12z"/><path d="M10.5 7a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3z"/></svg>' if is_img else '<svg viewBox="0 0 16 16" width="20" height="20" fill="currentColor"><path d="M14 14V4.5L9.5 0H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2zM9.5 1.5 13 5h-3.5V1.5zM3 9h10v1H3V9zm0 2h10v1H3v-1zm0 2h7v1H3v-1z"/></svg>'

                html_template += f"""
                <div class="file-card" data-name="{name.lower()}">
                    <div class="file-header">
                        <div class="file-icon {icon_cls}">{icon_svg}</div>
                        <div class="file-info">
                            <div class="file-name" title="{name}">{name}</div>
                            <div class="file-meta">
                                <span>ขนาด: {size_mb:.4f} MB</span>
                                <span>แก้ไขล่าสุด: {created}</span>
                            </div>
                        </div>
                    </div>
                """
                if is_img:
                    html_template += f"""
                    <div class="thumbnail-container" onclick="openPreview('/download/{name}?direct=true{token_query}', '{name}')">
                        <img class="thumbnail-img" src="/download/{name}?direct=true{token_query}" alt="{name}">
                    </div>
                    """
                html_template += f"""
                    <div class="file-actions">
                """
                if is_img:
                    html_template += f"""<button class="action-btn btn-view" onclick="openPreview('/download/{name}?direct=true{token_query}', '{name}')">👁️ ดูภาพ</button>"""
                else:
                    html_template += f"""<a class="action-btn btn-view" href="/download/{name}?direct=true{token_query}" target="_blank">👁️ ดูข้อมูล</a>"""
                html_template += f"""
                        <a class="action-btn btn-download" href="/download/{name}?direct=true{token_query}" download>📥 ดาวน์โหลด</a>
                        <button class="action-btn btn-delete" onclick="deleteFile('{name}')">🗑️ ลบ</button>
                    </div>
                </div>
                """

        html_template += f"""
        </div>
    </div>
    <div id="previewModal" class="modal" onclick="closePreview()">
        <div class="modal-content" onclick="event.stopPropagation()">
            <div class="modal-header">
                <span class="modal-title" id="modalTitle">ดูตัวอย่างภาพ</span>
                <span class="close-btn" onclick="closePreview()">&times;</span>
            </div>
            <div class="modal-body"><img id="modalImg" src="" alt="Preview"></div>
        </div>
    </div>
    <script>
        function filterFiles() {{
            const query = document.getElementById('searchInput').value.toLowerCase();
            const cards = document.getElementsByClassName('file-card');
            for (let card of cards) {{
                card.style.display = card.getAttribute('data-name').includes(query) ? 'flex' : 'none';
            }}
        }}
        function openPreview(src, title) {{
            document.getElementById('modalImg').src = src;
            document.getElementById('modalTitle').textContent = title;
            document.getElementById('previewModal').style.display = 'flex';
        }}
        async function deleteFile(fileName) {{
            if (!confirm(`คุณแน่ใจหรือไม่ว่าต้องการลบไฟล์ "${{fileName}}"?`)) return;
            try {{
                const res = await fetch(`/files/${{encodeURIComponent(fileName)}}`, {{
                    method: 'DELETE',
                    headers: {{ 'x-cleanup-token': '{CLEANUP_SECRET or ""}' }}
                }});
                const data = await res.json();
                if (res.ok && data.success) {{
                    alert('ลบไฟล์เรียบร้อยแล้ว!');
                    window.location.reload();
                }} else {{
                    alert('เกิดข้อผิดพลาด: ' + (data.detail || data.message));
                }}
            }} catch (err) {{
                alert('เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์: ' + err.message);
            }}
        }}
        function closePreview() {{
            document.getElementById('previewModal').style.display = 'none';
        }}
        async function triggerCleanup() {{
            if (!confirm('คุณแน่ใจหรือไม่ว่าต้องการสั่งล้างไฟล์ขยะและไฟล์เก่าทั้งหมดในระบบ?')) return;
            try {{
                const res = await fetch('/internal/cleanup', {{
                    method: 'POST',
                    headers: {{ 'x-cleanup-token': '{CLEANUP_SECRET or ""}' }}
                }});
                const data = await res.json();
                if (data.success) {{ alert('ล้างไฟล์ขยะเรียบร้อยแล้ว!'); window.location.reload(); }}
                else {{ alert('เกิดข้อผิดพลาด: ' + data.message); }}
            }} catch (err) {{
                alert('เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์: ' + err.message);
            }}
        }}
    </script>
</body>
</html>
"""
        return HTMLResponse(content=html_template)
    except Exception as e:
        logger.error(f"Error listing files: {e}", exc_info=True)
        return {"success": False, "message": f"Failed to list files: {str(e)}"}

@app.delete("/files/{file_name}")
async def delete_file(
    file_name: str,
    x_cleanup_token: Optional[str] = Header(None),
    token: Optional[str] = None,
):
    effective_token = x_cleanup_token or token
    if CLEANUP_SECRET and effective_token != CLEANUP_SECRET:
        raise HTTPException(status_code=401, detail="Unauthorized")

    if not re.match(r"^(chart_)?[A-Za-z0-9_]+_[a-f0-9]{8}_\d{8}_\d{6}\.(xlsx|png)$", file_name):
        raise HTTPException(status_code=400, detail="รูปแบบชื่อไฟล์ไม่ถูกต้อง")

    file_path = (REPORT_DIR / file_name).resolve()
    if not file_path.is_relative_to(REPORT_DIR):
        raise HTTPException(status_code=403, detail="ไม่มีสิทธิ์เข้าถึง")
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="ไม่พบไฟล์")

    try:
        file_path.unlink()
        logger.info(f"Deleted file: {file_name}")
        return {"success": True, "message": f"ลบไฟล์ {file_name} เรียบร้อยแล้ว"}
    except Exception as e:
        logger.error(f"Failed to delete {file_name}: {e}")
        raise HTTPException(status_code=500, detail=f"ลบไฟล์ไม่สำเร็จ: {str(e)}")

@app.post("/internal/cleanup")
async def trigger_cleanup(x_cleanup_token: Optional[str] = Header(None)):  # Add Optional
    if CLEANUP_SECRET and x_cleanup_token != CLEANUP_SECRET:  # Remove `not CLEANUP_SECRET` — allow if no secret configured
        logger.warning("Unauthorized cleanup attempt")
        raise HTTPException(status_code=401, detail="Unauthorized")
    logger.info("Triggering manual report cleanup via API...")
    await asyncio.to_thread(cleanup_old_reports)
    return {"success": True, "message": "Cleanup completed successfully"}

# ============================================================
# SECTION 15 — MOUNT MCP & ENTRYPOINT
# ============================================================
app.mount("", mcp_asgi_app)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)