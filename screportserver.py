"""
SC Report Server - Minimal Working Version
Avoids google.cloud namespace pollution by importing submodules directly
"""

import os
import re
import sys
import uuid
import logging
import io
from datetime import timedelta
from contextlib import asynccontextmanager
from contextvars import ContextVar
from typing import Optional, Tuple

# ============================================================================
# CRITICAL: Import google cloud submodules DIRECTLY (avoids namespace issues)
# ============================================================================
try:
    import google.cloud.bigquery as bigquery
    import google.cloud.storage as storage
except ImportError as e:
    print(f"FATAL: Google Cloud libraries not installed")
    print(f"Error: {e}")
    print(f"Fix: pip install google-cloud-bigquery google-cloud-storage")
    sys.exit(1)

# Import FastAPI after google.cloud is safe
try:
    from fastapi import FastAPI, Request
    from fastapi.middleware.cors import CORSMiddleware
    from pydantic import BaseModel
    from openpyxl import Workbook
    from mcp.server.fastmcp import FastMCP
except ImportError as e:
    print(f"FATAL: Required library missing: {e}")
    sys.exit(1)

# ============================================================================
# Configuration
# ============================================================================
PROJECT_ID = os.environ.get("PROJECT_ID", "sc-ai-uat")
DATASET_ID = os.environ.get("DATASET_ID", "SCReport")
AUTH_TABLE = os.environ.get("AUTH_TABLE", "AuthenByMenu")
GCS_BUCKET = os.environ.get("GCS_REPORT_BUCKET", "sc-report-files-uat")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("sc-report")

# Global clients
bq_client = None
gcs_client = None

# Request context
CURRENT_USERNAME = ContextVar("username", default=None)


# ============================================================================
# Models
# ============================================================================
class GenerateReportRequest(BaseModel):
    table_id: str
    limit: int = 2000
    filter_column: Optional[str] = None
    filter_value: Optional[str] = None
    condition: Optional[str] = None
    username: Optional[str] = None
    user_email: Optional[str] = None

    class Config:
        extra = "allow"


# ============================================================================
# Report Mapping
# ============================================================================
TABLE_MAPPING = {
    "vrptexpension": {
        "report_names": [
            "รายงานตรวจสอบการจ่ายเงิน (ต้นทุน)",
            "รายงานตรวจสอบการจ่ายเงิน(ต้นทุน)",
        ],
        "bq_table": "VRptExpension",
    },
    "vrptexpensionexpmodule": {
        "report_names": [
            "รายงานตรวจสอบการจ่ายเงิน (ค่าใช้จ่าย)",
            "รายงานตรวจสอบการจ่ายเงิน(ค่าใช้จ่าย)",
        ],
        "bq_table": "VRptExpensionExpModule",
    },
}


# ============================================================================
# Initialization
# ============================================================================
def init_clients():
    """Initialize BigQuery and GCS clients"""
    global bq_client, gcs_client

    try:
        logger.info(f"Initializing BigQuery (project: {PROJECT_ID})")
        bq_client = bigquery.Client(project=PROJECT_ID)
        logger.info("✅ BigQuery client ready")
    except Exception as e:
        logger.error(f"FATAL: BigQuery init failed: {e}")
        sys.exit(1)

    try:
        logger.info(f"Initializing GCS (bucket: {GCS_BUCKET})")
        gcs_client = storage.Client(project=PROJECT_ID)
        bucket = gcs_client.bucket(GCS_BUCKET)
        bucket.reload()
        logger.info("✅ GCS client ready")
    except Exception as e:
        logger.error(f"FATAL: GCS init failed: {e}")
        sys.exit(1)


# ============================================================================
# Security & Validation
# ============================================================================
def clean_table_id(raw: str) -> str:
    return raw.strip().strip("`").split(".")[-1].strip().lower()


def is_valid_user(user: Optional[str]) -> bool:
    if not user:
        return False
    s = user.strip()
    return s and "${" not in s and "}}" not in s


def get_username(user: Optional[str], email: Optional[str]) -> str:
    ctx = CURRENT_USERNAME.get()
    if is_valid_user(ctx):
        return ctx
    if is_valid_user(user):
        return user
    if is_valid_user(email):
        return email
    return ""


def check_permission(username: str, table_id: str) -> Optional[str]:
    """Check if user can access table. Returns report name if allowed."""
    table_clean = clean_table_id(table_id)

    if "authenby" in table_clean:
        logger.warning(f"🚨 Blocked: {username} tried to access auth table")
        return None

    mapping = TABLE_MAPPING.get(table_clean)
    if not mapping:
        return None

    try:
        report_names = mapping["report_names"]
        placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
        query = f"""
            SELECT ReportName FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE LOWER(TRIM(UserName)) = LOWER(TRIM(@username))
            AND TRIM(ReportName) IN ({placeholders})
            LIMIT 1
        """

        params = [
            bigquery.ScalarQueryParameter("username", "STRING", username.strip())
        ]
        for i, name in enumerate(report_names):
            params.append(bigquery.ScalarQueryParameter(f"name{i}", "STRING", name))

        job_config = bigquery.QueryJobConfig(query_parameters=params)
        results = list(bq_client.query(query, job_config=job_config))
        
        if results:
            return results[0]["ReportName"]
        return None
    except Exception as e:
        logger.error(f"Permission check failed: {e}")
        return None


# ============================================================================
# Report Generation
# ============================================================================
def generate_report(
    table_id: str,
    limit: int = 2000,
    filter_col: Optional[str] = None,
    filter_val: Optional[str] = None,
    condition: Optional[str] = None,
) -> Tuple[Optional[int], Optional[str]]:
    """Generate Excel and upload to GCS. Returns (row_count, signed_url)"""
    
    table_clean = clean_table_id(table_id)
    if "authenby" in table_clean:
        raise PermissionError("Auth table access denied")

    mapping = TABLE_MAPPING.get(table_clean)
    if not mapping:
        raise ValueError(f"Unknown table: {table_id}")

    bq_table = mapping["bq_table"]
    query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{bq_table}`"
    params = []

    # Apply filters
    if filter_col and filter_val and re.match(r"^[A-Za-z0-9_]+$", filter_col):
        query += f" WHERE {filter_col} = @val"
        params.append(bigquery.ScalarQueryParameter("val", "STRING", filter_val))
    elif condition:
        # Block dangerous SQL
        dangerous = [";", "union", "select", "insert", "delete", "drop"]
        if any(x in condition.lower() for x in dangerous):
            raise PermissionError("Unsafe condition")
        query += f" WHERE {condition}"

    query += f" LIMIT {min(max(limit, 1), 10000)}"

    # Execute query
    logger.info(f"Query: {bq_table} (limit: {limit})")
    job_config = bigquery.QueryJobConfig(query_parameters=params)
    results = list(bq_client.query(query, job_config=job_config).result())

    if not results:
        return None, None

    # Create Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    schema = bq_client.query(query, job_config=job_config).result().schema
    ws.append([f.name for f in schema])

    for row in results:
        values = []
        for field in schema:
            val = row[field.name]
            if hasattr(val, "tzinfo") and val.tzinfo:
                val = val.replace(tzinfo=None)
            values.append(val)
        ws.append(values)

    # Upload to GCS
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    file_id = uuid.uuid4().hex
    filename = f"Report_{table_id}_{file_id}.xlsx"

    try:
        bucket = gcs_client.bucket(GCS_BUCKET)
        blob = bucket.blob(filename)
        blob.upload_from_file(
            buffer,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        url = blob.generate_signed_url(
            version="v4", expiration=timedelta(minutes=15), method="GET"
        )
        logger.info(f"✅ Uploaded: {filename}")
        return len(results), url
    except Exception as e:
        logger.error(f"GCS upload failed: {e}")
        raise


# ============================================================================
# FastAPI Setup
# ============================================================================
mcp = FastMCP("sc-report", stateless_http=True, json_response=True, host="0.0.0.0")
mcp_app = mcp.streamable_http_app()


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("=" * 70)
    logger.info("🚀 SC Report Server Starting")
    logger.info(f"   Project: {PROJECT_ID} | Dataset: {DATASET_ID} | Bucket: {GCS_BUCKET}")
    init_clients()
    async with mcp.session_manager.run():
        logger.info("✅ All systems ready")
        logger.info("=" * 70)
        yield


app = FastAPI(title="SC Report", version="2.0.0", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def extract_user(request: Request, call_next):
    user = None
    for header in ["current-user", "current_user", "x-user-email"]:
        if val := request.headers.get(header):
            user = val
            break
    token = CURRENT_USERNAME.set(user)
    try:
        response = await call_next(request)
    finally:
        CURRENT_USERNAME.reset(token)
    return response


@app.get("/health")
async def health():
    """Health check endpoint"""
    if not bq_client or not gcs_client:
        return {"status": "initializing"}
    return {"status": "healthy"}


@app.post("/generate_excel_report")
async def api_generate_report(req: GenerateReportRequest):
    """REST endpoint for report generation"""
    table = clean_table_id(req.table_id)
    user = get_username(req.username, req.user_email)

    if not user:
        return {"success": False, "message": "No user"}
    if "authenby" in table:
        return {"success": False, "message": "Access denied"}
    if not re.match(r"^[a-z0-9_]+$", table):
        return {"success": False, "message": "Invalid table"}

    report_name = check_permission(user, table)
    if not report_name:
        return {"success": False, "message": "No permission"}

    try:
        rows, url = generate_report(
            table,
            req.limit,
            req.filter_column,
            req.filter_value,
            req.condition,
        )
        if not rows:
            return {"success": False, "message": "No data"}
        return {
            "success": True,
            "message": f"Report: {report_name}",
            "rows": rows,
            "url": url,
        }
    except Exception as e:
        logger.error(f"Error: {e}")
        return {"success": False, "message": str(e)}


@mcp.tool(name="generate_excel_report", description="Generate and download report")
def mcp_generate_report(
    table_id: str,
    limit: int = 2000,
    filter_column: Optional[str] = None,
    filter_value: Optional[str] = None,
    condition: Optional[str] = None,
    username: Optional[str] = None,
    user_email: Optional[str] = None,
    **kwargs,
) -> str:
    """MCP tool endpoint"""
    table = clean_table_id(table_id)
    user = get_username(username, user_email)

    if not user:
        return "No user"
    if "authenby" in table:
        return "Access denied"

    report_name = check_permission(user, table)
    if not report_name:
        return "No permission"

    try:
        rows, url = generate_report(table, limit, filter_column, filter_value, condition)
        if not rows:
            return f"No data: {report_name}"
        return f"✅ {report_name} ({rows} rows)\n🔗 {url}"
    except Exception as e:
        return f"Error: {e}"


app.mount("", mcp_app)


# ============================================================================
# Entry Point
# ============================================================================
if __name__ == "__main__":
    import uvicorn

    logger.info("Starting uvicorn on 0.0.0.0:8080")
    uvicorn.run(
        "screportserver:app",
        host="0.0.0.0",
        port=8080,
        log_level="info",
    )