import os
import re
import uuid
import logging
import time
import asyncio
from datetime import datetime, timedelta
from contextlib import asynccontextmanager
from contextvars import ContextVar
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from fastapi import FastAPI, Request, HTTPException, BackgroundTasks
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
LOG_TABLE = "AiLog"
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
# ✅ Midnight Cleanup Scheduler
# --------------------------------------------------------
async def _midnight_cleanup_loop() -> None:
    while True:
        now = datetime.now()
        next_midnight = (now + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        seconds_until_midnight = (next_midnight - now).total_seconds()
        logger.info(f"🕛 Midnight cleanup scheduled in {seconds_until_midnight:.0f}s (at {next_midnight})")
        await asyncio.sleep(seconds_until_midnight)
        logger.info("🕛 Running midnight report cleanup...")
        cleanup_old_reports()
        await asyncio.sleep(1)

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
    cleanup_task = asyncio.create_task(_midnight_cleanup_loop())
    logger.info("✅ Midnight cleanup scheduler started")
    async with mcp.session_manager.run():
        yield
    cleanup_task.cancel()
    try:
        await cleanup_task
    except asyncio.CancelledError:
        logger.info("🛑 Midnight cleanup scheduler stopped")

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
# Middleware — จับ current-user หรือ x-user-email header
# --------------------------------------------------------
@app.middleware("http")
async def capture_username_header(request: Request, call_next):
    username_val = None
    for header_name in ["current-user", "current_user", "x-user-email"]:
        val = request.headers.get(header_name)
        if val:
            username_val = val
            break
    
    logger.info(f"📨 [Middleware] {request.method} {request.url.path} | Resolved: {repr(username_val)}")
    
    token_username = CURRENT_REQUEST_USERNAME.set(username_val)
    base_url = str(request.base_url).rstrip("/")
    token_base = CURRENT_REQUEST_BASE_URL.set(base_url)

    try:
        response = await call_next(request)
    finally:
        CURRENT_REQUEST_USERNAME.reset(token_username)
        CURRENT_REQUEST_BASE_URL.reset(token_base)
    return response

# --------------------------------------------------------
# Models
# --------------------------------------------------------
class FilterCondition(BaseModel):
    column: str
    operator: str
    value: str

class GenerateReportRequest(BaseModel):
    table_id: str
    limit: int = 2000
    filter_column: Optional[str] = None
    filter_value: Optional[str] = None
    filters: Optional[list[FilterCondition]] = None
    condition: Optional[str] = None
    username: Optional[str] = None
    user_email: Optional[str] = None
    columns: Optional[list[str]] = None

    class Config:
        extra = "allow"

# --------------------------------------------------------
# Mapping Configuration
# --------------------------------------------------------
TABLE_TO_REPORT_NAMES: dict[str, list[str]] = {
    "vrptexpension": [
        "รายงานตรวจสอบการจ่ายเงิน (ต้นทุน)", "รายงานตรวจสอบการจ่ายเงิน(ต้นทุน)",
        "ทะเบียนคุมการเบิกจ่าย (ต้นทุน)", "ทะเบียนคุมการเบิกจ่าย(ต้นทุน)"
    ],
    "vrptexpensionexpmodule": [
        "รายงานตรวจสอบการจ่ายเงิน (ค่าใช้จ่าย)", "รายงานตรวจสอบการจ่ายเงิน(ค่าใช้จ่าย)",
        "ทะเบียนคุมการเบิกจ่าย (ค่าใช้จ่าย)", "ทะเบียนคุมการเบิกจ่าย(ค่าใช้จ่าย)"
    ],
}

TABLE_TO_BQ_TABLE_NAME: dict[str, str] = {
    "vrptexpension": "VRptExpension",
    "vrptexpensionexpmodule": "VRptExpensionExpModule",
}

ALLOWED_OPERATORS: set[str] = {"=", ">", "<", ">=", "<=", "LIKE", "!="}

# --------------------------------------------------------
# 🔥 Full Hardcoded Mappingแยกตารางอย่างสมบูรณ์แบบ (รวม 118 คอลัมน์)
# --------------------------------------------------------
TABLE_COLUMN_MAPPING: dict[str, dict[str, str]] = {
    # --- 1. รายงาน: VRptExpensionExpModule (ค่าใช้จ่าย) - 61 คอลัมน์ ---
    "vrptexpensionexpmodule": {
        "DocumentType": "ประเภทเอกสาร",
        "SubDocumentType": "ประเภทเอกสารย่อย",
        "DocumentID": "ไอดีเอกสาร (Internal ID)",
        "DocumentNo": "เลขที่เอกสาร",
        "DocumentDate": "วันที่เอกสาร",
        "ApprovedDate": "วันที่อนุมัติเอกสาร",
        "VendorName": "ชื่อผู้รับเหมา/Supplier",
        "WorkGroupCode": "รหัสกลุ่มงาน",
        "WorkGroupName": "ชื่อกลุ่มงาน",
        "TotalPriceAfterVAT": "มูลค่างานหลัง Vat",
        "IsApprovedCancel": "สถานะอนุมัติการยกเลิก (0/1)",
        "IsApproved": "สถานะการอนุมัติ (0/1)",
        "PlotCode": "รหัสแปลง",
        "TotalReduceAfterVAT": "มูลค่าส่วนลดหลัง VAT",
        "TotalCanceledPriceAfterVat": "มูลค่ายกเลิกหลัง VAT",
        "TotalNetPrice": "มูลค่าสุทธิ",
        "SAPInvDoc": "เลขที่ใบแจ้งหนี้ SAP",
        "SAPInvDocDate": "วันที่ใบแจ้งหนี้ SAP",
        "SAPPayDoc": "เลขที่เอกสารการจ่าย SAP",
        "SAPPayDocDate": "วันที่เอกสารการจ่าย SAP",
        "Remark": "หมายเหตุ",
        "ApprovedCancelDate": "วันที่อนุมัติยกเลิก",
        "CancelReason": "เหตุผลที่ยกเลิก",
        "StatusDesc": "คำอธิบายสถานะ",
        "Status": "รหัสสถานะ",
        "IsEarnest": "สถานะเงินมัดจำ",
        "ProjectID": "รหัสโครงการ",
        "WorkDescription": "รายละเอียดงาน",
        "VendorID": "รหัสผู้ขาย/ผู้รับเหมา",
        "WorkGroupID": "ไอดีกลุ่มงาน",
        "WorkTypeID": "ไอดีประเภทงาน",
        "WorkSubTypeID": "ไอดีประเภทงานย่อย",
        "SummaryPayAfterVat": "สรุปยอดจ่ายหลัง VAT",
        "CntApprovedPoReceive": "จำนวนใบรับของที่อนุมัติแล้ว",
        "IsPay": "สถานะการจ่ายเงิน (0/1)",
        "PayDate": "วันที่จ่ายเงิน",
        "TotalRetentionAmount": "จำนวนเงินประกันผลงานรวม",
        "SubDocumentNo": "เลขที่เอกสารย่อย",
        "SubDocumentDate": "วันที่เอกสารย่อย",
        "SubIsApproved": "สถานะอนุมัติเอกสารย่อย",
        "SubApprovedDate": "วันที่อนุมัติเอกสารย่อย",
        "PeriodNoShow": "เลขงวดงาน",
        "SubTotalRetentionAmount": "เงินประกันผลงาน (ย่อย)",
        "SubTotalPriceAfterVat": "มูลค่างานหลัง Vat (ย่อย)",
        "IsPayApproved": "สถานะอนุมัติการจ่าย",
        "PayApprovedDate": "วันที่อนุมัติการจ่าย",
        "PayApprovedDocumentNo": "เลขที่เอกสารอนุมัติการจ่าย",
        "PayApprovedDocumentDate": "วันที่เอกสารอนุมัติการจ่าย",
        "TotalWithDrawalPriceAfterVat": "ยอดเบิกสะสมหลัง Vat",
        "TotatWorkPrice": "ราคางานทั้งหมด",
        "ApprovedGRCnt": "จำนวนการตรวจรับที่อนุมัติ (GR)",
        "AgreementTypeID": "รหัสประเภทสัญญา",
        "SubTotalDeducAmount": "ยอดหักคืน (ย่อย)",
        "TotalDeducAmount": "ยอดหักคืนรวม",
        "TotalCN": "ยอดลดหนี้รวม (CN)",
        "RefDocumentNo": "เลขที่เอกสารอ้างอิง",
        "MainDocNo": "เลขที่เอกสารหลัก",
        "ContractDocNo": "เลขที่สัญญา",
        "DepartBudgetCode": "รหัสงบประมาณแผนก",
        "PCItemID": "รหัสรายการ PC",
        "TotalPayment": "มูลค่าเบิก-จ่ายรวม (Net Payment)"
    },

    # --- 2. รายงาน: VRptExpension (ต้นทุน) - 57 คอลัมน์ ---
    "vrptexpension": {
        "DocumentType": "ประเภทเอกสาร",
        "SubDocumentType": "ประเภทเอกสารย่อย",
        "DocumentID": "ไอดีเอกสาร (Internal ID)",
        "DocumentNo": "เลขที่เอกสาร",
        "DocumentDate": "วันที่เอกสาร",
        "ApprovedDate": "วันที่อนุมัติเอกสาร",
        "VENDorName": "ชื่อผู้รับเหมา/Supplier",
        "WorkGroupCode": "รหัสกลุ่มงาน",
        "WorkGroupName": "ชื่อกลุ่มงาน",
        "TotalPriceAfterVAT": "มูลค่างานหลัง Vat",
        "IsApprovedCancel": "สถานะอนุมัติการยกเลิก",
        "IsApproved": "สถานะการอนุมัติ",
        "PlotCode": "รหัสแปลงที่ดิน",
        "TotalReduceAfterVAT": "มูลค่าส่วนลดหลัง VAT",
        "TotalCanceledPriceAfterVat": "มูลค่ายกเลิกหลัง VAT",
        "TotalNetPrice": "มูลค่าสุทธิ",
        "SAPInvDoc": "เลขที่ใบแจ้งหนี้ SAP",
        "SAPInvDocDate": "วันที่ใบแจ้งหนี้ SAP",
        "SAPPayDoc": "เลขที่เอกสารการจ่าย SAP",
        "SAPPayDocDate": "วันที่เอกสารการจ่าย SAP",
        "Remark": "หมายเหตุ",
        "ApprovedCancelDate": "วันที่อนุมัติยกเลิก",
        "CancelReASon": "เหตุผลที่ยกเลิก",
        "StatusDesc": "คำอธิบายสถานะ",
        "Status": "รหัสสถานะ",
        "IsEarnest": "สถานะเงินมัดจำ",
        "ProjectID": "รหัสโครงการ",
        "WorkDescription": "รายละเอียดงาน",
        "VENDorID": "รหัสผู้รับเหมา",
        "WorkGroupID": "ไอดีกลุ่มงาน",
        "SummaryPayAfterVat": "สรุปยอดจ่ายหลัง VAT",
        "CntApprovedPoReceive": "จำนวนใบรับของที่อนุมัติแล้ว",
        "IsPay": "สถานะการจ่ายเงิน",
        "PayDate": "วันที่จ่ายเงิน",
        "TotalRetentionAmount": "จำนวนเงินประกันผลงานรวม",
        "SubDocumentNo": "เลขที่เอกสารย่อย",
        "SubDocumentDate": "วันที่เอกสารย่อย",
        "SubIsApproved": "สถานะอนุมัติเอกสารย่อย",
        "SubApprovedDate": "วันที่อนุมัติเอกสารย่อย",
        "PeriodNoShow": "เลขงวดงาน",
        "SubTotalRetentionAmount": "เงินประกันผลงาน (ย่อย)",
        "SubTotalPriceAfterVat": "มูลค่างานหลัง Vat (ย่อย)",
        "SubAdjustPOPriceAfterVat": "ยอดปรับปรุง PO หลัง Vat (ย่อย)",
        "IsPayApproved": "สถานะอนุมัติการจ่าย",
        "PayApprovedDate": "วันที่อนุมัติการจ่าย",
        "PayApprovedDocumentNo": "เลขที่เอกสารอนุมัติการจ่าย",
        "PayApprovedDocumentDate": "วันที่เอกสารอนุมัติการจ่าย",
        "TotalWithDrawalPriceAfterVat": "ยอดเบิกสะสมหลัง Vat",
        "TotatWorkPrice": "ราคางานทั้งหมด",
        "ApprovedGRCnt": "จำนวนการตรวจรับที่อนุมัติ",
        "AgreementTypeID": "รหัสประเภทสัญญา",
        "WorkTypeName": "ประเภทงาน",
        "WorkTypeID": "รหัสประเภทงาน",
        "IsInvoice": "สถานะใบกำกับภาษี",
        "TotalPayment": "มูลค่าเบิก-จ่ายรวม (Net Payment)",
        "ProcurementTypeID": "รหัสประเภทการจัดซื้อ",
        "TotalDeducAmount": "ยอดหักคืนรวม"
    }
}

# --------------------------------------------------------
# Helpers & Security Core
# --------------------------------------------------------
def clean_table_id(raw: str) -> str:
    return raw.strip().strip("`").split(".")[-1].strip()

def validate_table_id(table_id: str) -> bool:
    return bool(re.match(r"^[A-Za-z0-9_]+$", table_id))

def is_valid_user(user: Optional[str]) -> bool:
    if not user:
        return False
    user_str = user.strip()
    if not user_str or "${" in user_str or "}" in user_str or "{{" in user_str:
        return False
    return True

def resolve_username(fallback_user: Optional[str] = None, fallback_email: Optional[str] = None) -> str:
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
    if "AuthenByMenu" in clean:
        logger.warning(f"🚨 Security Blocked: User '{username}' tried to access permission table")
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
    return bool(re.match(r"^[A-Za-z0-9_]+$", column))

def build_complete_condition_log(
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    filters: Optional[list[FilterCondition]] = None,
    condition: Optional[str] = None,
    row_generated: int = 0,
) -> str:
    parts = []
    if filters:
        filter_strs = [f"{f.column}{f.operator}{f.value}" for f in filters]
        if filter_strs:
            parts.append(f"filters=[{', '.join(filter_strs)}]")
    if filter_column and filter_value:
        parts.append(f"legacy:{filter_column}={filter_value}")
    if condition:
        parts.append(f"custom:{condition}")
    if row_generated > 0:
        parts.append(f"rows={row_generated}")
    return "; ".join(parts) if parts else "no_filter"

def generate_ai_log_id() -> int:
    return int(uuid.uuid4().hex[:8], 16) % 999_999_999

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
        now = datetime.utcnow()
        if ai_log_id == 0:
            ai_log_id = generate_ai_log_id()

        rows_to_insert = [{
            "CreatedAt":     now.strftime("%Y-%m-%dT%H:%M:%S"),
            "UserName":      (username or "")[:50],
            "TableName":     (table_name or "")[:50],
            "ReportName":    (report_name or "")[:100],
            "Status":        (status or "")[:5],
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
            logger.info(f"✅ AiLog OK | id={ai_log_id} user={username} table={table_name} status={status}")
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
    clean = clean_table_id(table_id).lower()
    if "AuthenByMenu" in clean:
        raise PermissionError("Access to the authentication table is strictly prohibited at code level.")

    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(clean, table_id)

    if columns:
        for col in columns:
            if not validate_filter_column(clean, col):
                logger.warning(f"🚨 Column Blocked in SELECT: '{col}'")
                raise PermissionError(f"คอลัมน์ '{col}' ไม่ได้รับอนุญาต")
        select_clause = ", ".join([f"`{col}`" for col in columns])
    else:
        select_clause = "*"

    query = f"SELECT {select_clause} FROM `{PROJECT_ID}.{DATASET_ID}.{bq_table_name}`"
    params = []
    where_clauses = []
    condition_str_parts = []

    if filters:
        for i, f in enumerate(filters):
            if not validate_filter_column(clean, f.column):
                raise PermissionError(f"คอลัมน์ '{f.column}' ไม่ได้รับอนุญาตสำหรับตารางนี้")
            if f.operator not in ALLOWED_OPERATORS:
                raise PermissionError(f"Operator '{f.operator}' ไม่ได้รับอนุญาต")
            param_name = f"filter_{i}"
            where_clauses.append(f"`{f.column}` {f.operator} @{param_name}")
            params.append(bigquery.ScalarQueryParameter(param_name, "STRING", f.value))
            condition_str_parts.append(f"{f.column} {f.operator} ?")

    elif filter_column and filter_value:
        if not validate_filter_column(clean, filter_column):
            raise PermissionError(f"คอลัมน์ '{filter_column}' ไม่ได้รับอนุญาตสำหรับตารางนี้")
        where_clauses.append(f"`{filter_column}` = @filter_value")
        params.append(bigquery.ScalarQueryParameter("filter_value", "STRING", filter_value))
        condition_str_parts.append(f"{filter_column} = ?")

    if where_clauses:
        query += " WHERE " + " AND ".join(where_clauses)

    safe_limit = min(max(limit, 1), 10000)
    query += f" LIMIT {safe_limit}"

    logger.info(f"🔍 Final Query: {query}")

    job_config = bigquery.QueryJobConfig(query_parameters=params)
    job = bq_client.query(query, job_config=job_config)
    result = job.result()

    bytes_billed: int = job.total_bytes_billed or 0
    schema = result.schema
    results = list(result)

    condition_log = build_complete_condition_log(
        filter_column=filter_column,
        filter_value=filter_value,
        filters=filters,
        condition=" AND ".join(condition_str_parts) if condition_str_parts else None,
        row_generated=len(results),
    )

    if not results:
        return None, None, bytes_billed, 0.0, condition_log, 0

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    # ดึงชุดพจนานุกรมคำแปลภาษาไทยแยกตาม table_id (clean) เพื่อความแม่นยำสูงสุด
    mapping_dict = TABLE_COLUMN_MAPPING.get(clean, {})
    
    headers = []
    for field in schema:
        # หากเจอนามสกุลคอลัมน์ที่ถูกแมปไว้จะใช้คำแปลภาษาไทยทันที ถ้าไม่เจอจะใช้ชื่อคอลัมน์เดิม
        headers.append(mapping_dict.get(field.name, field.name))
            
    ws.append(headers)

    for row in results:
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
    return len(results), file_name, bytes_billed, file_size_mb, condition_log, ai_log_id


def cleanup_old_reports() -> None:
    deleted = 0
    try:
        for file_path in REPORT_DIR.glob("*.xlsx"):
            if file_path.is_file():
                file_path.unlink()
                deleted += 1
        logger.info(f"🧹 Midnight cleanup done — {deleted} file(s) removed")
    except Exception as e:
        logger.error(f"Error during report cleanup: {e}")

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
    username_to_use = resolve_username(username, None)
    clean = clean_table_id(table_id).lower()
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(clean, table_id)

    if not username_to_use:
        insert_ai_log("", bq_table_name, "", "FAIL")
        return {"success": False, "message": "❌ ไม่พบข้อมูลยืนยันตัวตนในระบบ"}

    if "AuthenByMenu" in table_id.lower():
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return {"success": False, "message": "🔒 [Access Denied] ระบบไม่อนุญาตให้เข้าถึงตารางสิทธิ์"}

    if not validate_table_id(table_id):
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return {"success": False, "message": f"❌ table_id '{table_id}' ไม่ถูกต้อง"}

    report_name = check_user_permission(username_to_use, table_id)
    if not report_name:
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return {"success": False, "message": "🙏 ขออภัย คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    parsed_filters = None
    if filters:
        try:
            parsed_filters = [FilterCondition(**f) for f in filters]
        except Exception:
            return {"success": False, "message": "❌ รูปแบบ filters ไม่ถูกต้อง"}

    try:
        row_count, file_name, bytes_billed, file_size_mb, condition_log, ai_log_id = fetch_and_generate_excel(
            table_id, report_name, limit,
            filter_column=filter_column,
            filter_value=filter_value,
            filters=parsed_filters,
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
        }
    except PermissionError as pe:
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": f"🔒 {str(pe)}"}
    except Exception as e:
        logger.error(f"MCP Tool Error: {e}", exc_info=True)
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": "🚨 เกิดข้อผิดพลาดภายในระบบ"}


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
    username_to_use = resolve_username(username, user_email)
    clean = clean_table_id(table_id).lower()
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(clean, table_id)

    if not username_to_use:
        insert_ai_log("", bq_table_name, "", "FAIL")
        return "❌ ไม่พบข้อมูลยืนยันตัวตนในระบบ"

    if "AuthenByMenu" in table_id.lower():
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return "🔒 [Access Denied] ระบบไม่อนุญาตให้เข้าถึงตารางสิทธิ์"

    if not validate_table_id(table_id):
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return "❌ table_id ไม่ถูกต้อง"

    report_name = check_user_permission(username_to_use, table_id)
    if not report_name:
        insert_ai_log(username_to_use, bq_table_name, "", "FAIL")
        return "🙏 ขออภัยครับ คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"

    parsed_filters = None
    if filters:
        try:
            parsed_filters = [FilterCondition(**f) for f in filters]
        except Exception:
            return "❌ รูปแบบ filters ไม่ถูกต้อง"

    try:
        row_count, file_name, bytes_billed, file_size_mb, condition_log, ai_log_id = fetch_and_generate_excel(
            table_id, report_name, limit,
            filter_column=filter_column,
            filter_value=filter_value,
            filters=parsed_filters,
            columns=columns,
        )
        if row_count is None:
            insert_ai_log(username_to_use, bq_table_name, report_name, "FAIL",
                          condition=condition_log, bytes_billed=bytes_billed, ai_log_id=ai_log_id)
            return f"❌ ไม่พบข้อมูลในรายงาน {report_name}"

        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"

        insert_ai_log(
            username=username_to_use, table_name=bq_table_name, report_name=report_name,
            status="OK", condition=condition_log, size=file_size_mb,
            bytes_billed=bytes_billed, url=download_url,
            row_generated=row_count or 0, ai_log_id=ai_log_id,
        )
        return f"📊 จัดเตรียม **รายงาน{report_name}** ({row_count} แถว) เรียบร้อยแล้วครับ\n🔗 {download_url}"
    except PermissionError as pe:
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return f"❌ {str(pe)}"
    except Exception as e:
        logger.error(f"Error: {e}")
        insert_ai_log(username_to_use, bq_table_name, report_name or "", "FAIL")
        return "🚨 เกิดข้อผิดพลาดภายในระบบ"

# --------------------------------------------------------
# REST API Endpoints
# --------------------------------------------------------
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest):
    tid = clean_table_id(request.table_id)
    bq_table_name = TABLE_TO_BQ_TABLE_NAME.get(tid.lower(), tid)
    username_val = resolve_username(request.username, request.user_email)

    if not username_val:
        insert_ai_log("", bq_table_name, "", "FAIL")
        return {"success": False, "message": "❌ Unauthorized access."}

    if "AuthenByMenu" in tid.lower():
        insert_ai_log(username_val, bq_table_name, "", "FAIL")
        return {"success": False, "message": "🔒 Access Denied"}

    if not validate_table_id(tid):
        insert_ai_log(username_val, bq_table_name, "", "FAIL")
        return {"success": False, "message": f"❌ table_id '{tid}' ไม่ถูกต้อง"}

    report_name = check_user_permission(username_val, tid)
    if not report_name:
        insert_ai_log(username_val, bq_table_name, "", "FAIL")
        return {"success": False, "message": "🙏 ขออภัย คุณไม่มีสิทธิ์เข้าถึงรายงานนี้"}

    try:
        row_count, file_name, bytes_billed, file_size_mb, condition_log, ai_log_id = fetch_and_generate_excel(
            tid, report_name,
            limit=request.limit,
            filter_column=request.filter_column,
            filter_value=request.filter_value,
            filters=request.filters,
            columns=request.columns,
        )
        if row_count is None:
            insert_ai_log(username_val, bq_table_name, report_name, "FAIL",
                          condition=condition_log, bytes_billed=bytes_billed, ai_log_id=ai_log_id)
            return {"success": False, "message": "❌ ไม่พบข้อมูล"}

        base_url = CURRENT_REQUEST_BASE_URL.get() or PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}"

        insert_ai_log(
            username=username_val, table_name=bq_table_name, report_name=report_name,
            status="OK", condition=condition_log, size=file_size_mb,
            bytes_billed=bytes_billed, url=download_url,
            row_generated=row_count or 0, ai_log_id=ai_log_id,
        )
        return {
            "success": True,
            "message": f"✅ จัดเตรียม **รายงาน{report_name}** สำเร็จ\n🔗 {download_url}",
        }
    except PermissionError as pe:
        insert_ai_log(username_val, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": f"🔒 {str(pe)}"}
    except Exception as e:
        logger.error(f"API Error: {e}")
        insert_ai_log(username_val, bq_table_name, report_name or "", "FAIL")
        return {"success": False, "message": "🚨 เกิดข้อผิดพลาดภายในเซิร์ฟเวอร์"}

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

app.mount("", mcp_asgi_app)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)