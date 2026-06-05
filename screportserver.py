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

# 📦 MCP FastMCP สำหรับ Streamable HTTP
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

# เพิ่ม CORS Middleware เพื่อให้เว็บ Tantawan เรียกใช้งานได้
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
DEFAULT_USER_EMAIL = "kachatharn@scasset.com"
PUBLIC_BASE_URL = ""

# กำหนด Path และทำให้แน่ใจว่าเป็  น absolute path
REPORT_DIR = Path("/tmp/reports").resolve()
REPORT_DIR.mkdir(parents=True, exist_ok=True)

# ⚡ ประกาศสร้าง Client เป็น Global เพื่อแชร์การเชื่อมต่อร่วมกัน
bq_client = bigquery.Client(project=PROJECT_ID)

CURRENT_REQUEST_USER_EMAIL: ContextVar[Optional[str]] = ContextVar(
    "current_request_user_email",
    default=None,
)

class GenerateReportRequest(BaseModel):
    table_id: str
    user_email: Optional[str] = None

@app.get("/")
def read_root():
    return {"message": "Welcome to SC Report API. MCP server is active at /mcp"}

@app.get("/health")
def health():
    return {"status": "ok", "service": "sc-report-api-mcp"}

@app.middleware("http")
async def capture_user_email_header(request: Request, call_next):
    # ดึงค่าจาก Header ตรงๆ
    email = request.headers.get("x-user-email") or DEFAULT_USER_EMAIL
    token = CURRENT_REQUEST_USER_EMAIL.set(email) # เก็บค่าลง ContextVar
    try:
        response = await call_next(request)
    finally:
        CURRENT_REQUEST_USER_EMAIL.reset(token)
    return response

def validate_table_id(table_id: str) -> bool:
    return bool(re.match(r"^[A-Za-z0-9_]+$", table_id))

# 🛡️ ตรวจสอบสิทธิ์โดยใช้ Global Client
def check_user_permission(email: str, table_id: str):
    print(f"🔍 [Authentication] ตรวจสอบสิทธิ์: {email} สำหรับตาราง: {table_id}")
    
    # 1. Mapping table_id (ชื่อจริงใน DB) ให้เป็น ReportName (ชื่อในตารางสิทธิ์)
    # ใช้ตัวเล็ก (lower) เพื่อลดความผิดพลาดในการเปรียบเทียบ
    mapping = {
        "vrptexpension": "รายงานตรวจสอบการจ่ายเงิน(ต้นทุน)",
        "vrptexpensionexpmodule": "รายงานตรวจสอบการจ่ายเงิน (ค่าใช้จ่าย)",
        "authenbymenu": "ข้อมูลสิทธิ์"
    }
    
    # Clean table_id: เอาแค่ชื่อตารางท้ายสุด และทำเป็นตัวเล็ก
    clean_table_id = table_id.split('.')[-1].lower()
    report_name = mapping.get(clean_table_id)
    
    # ถ้าไม่มีใน Mapping แสดงว่าไม่ใช่ตารางที่เราอนุญาตให้ Gen Excel
    if not report_name:
        print(f"🚨 [Error] ไม่พบ Mapping สำหรับตาราง: {clean_table_id}")
        return None

    try:
        # 2. ใช้ Query Parameters เหมือนของเดิมเพื่อความปลอดภัย (SQL Injection Prevention)
        # และใช้ TRIM() เพื่อดักช่องว่างในฐานข้อมูล
        auth_query = f"""
            SELECT ReportName
            FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
            WHERE LOWER(TRIM(Email)) = LOWER(@email) 
            AND TRIM(ReportName) = @report_name
            LIMIT 1
        """
        
        job_config = bigquery.QueryJobConfig(
            query_parameters=[
                bigquery.ScalarQueryParameter("email", "STRING", email.lower().strip()),
                bigquery.ScalarQueryParameter("report_name", "STRING", report_name),
            ]
        )
        
        auth_df = bq_client.query(auth_query, job_config=job_config).to_dataframe()
        
        if not auth_df.empty:
            print(f"✅ [Success] พบสิทธิ์: {report_name}")
            return report_name # คืนค่าชื่อรายงานกลับไปใช้แสดงผล
            
        print(f"❌ [Denied] ไม่พบสิทธิ์สำหรับ Email: {email} ในรายงาน: {report_name}")
        return None
        
    except Exception as auth_error:
        print(f"🚨 [System Error] ไม่สามารถตรวจสอบสิทธิ์ได้: {str(auth_error)}")
        return None

# ♻️ Helper Function ดึงข้อมูลและสร้างไฟล์ Excel
def fetch_and_generate_excel(table_id: str, where_clause: str = "") -> Optional[pd.DataFrame]:
    # ดึงชื่อตารางท้ายสุด (เช่น VRptExpension)
    clean_table_id = table_id.split('.')[-1]
    
    try:
        # สร้าง SQL โดยใส่ WHERE clause ถ้ามีการส่งมา
        sql_where = f"WHERE {where_clause}" if where_clause else ""
        query = f"SELECT * FROM `{PROJECT_ID}.{DATASET_ID}.{clean_table_id}` {sql_where} LIMIT 5000"
        
        df = bq_client.query(query).to_dataframe()
        
        if df.empty:
            return None

        # ตั้งชื่อไฟล์โดยใช้ clean_table_id เพื่อความปลอดภัยของระบบไฟล์
        file_name = f"Report_{clean_table_id}.xlsx"
        file_path = REPORT_DIR / file_name
        
        df.to_excel(file_path, index=False)
        return df

    except Exception as e:
        print(f"🚨 Error: {str(e)}")
        return None

# --------------------------------------------------------
# 🛠️ MCP Server (Streamable HTTP — stateless)
# --------------------------------------------------------

# FIX: stateless_http=True ไม่ต้องการ session manager lifecycle
mcp = FastMCP(
    "sc-report-server",
    stateless_http=True,
    json_response=True,
    host="0.0.0.0",
)

@mcp.tool(
    name="generate_excel_report",
    description="ดึงข้อมูลรายงานจาก BigQuery โดยจะเช็คสิทธิ์ผู้ใช้งานจากตาราง AuthenByMenu อัตโนมัติก่อนสร้างไฟล์ Excel",
)
def mcp_generate_excel_report(table_id: str, filter_condition: str = None, user_email: str = None) -> str:
    # 1. จัดการเรื่อง Email และ Table ID
    email = user_email or CURRENT_REQUEST_USER_EMAIL.get() or DEFAULT_USER_EMAIL
    clean_table_id = table_id.split('.')[-1]

    if not validate_table_id(clean_table_id):
        return f"❌ table_id '{clean_table_id}' ไม่ถูกต้อง"

    # 2. เช็คสิทธิ์ (ใช้ clean_table_id เพื่อให้ตรงกับใน DB AuthenByMenu)
    report_name = check_user_permission(email, clean_table_id)
    if not report_name:
        return f"🙏 ขออภัยคุณ {email} ระบบตรวจสอบพบว่าคุณยังไม่มีสิทธิ์เข้าถึงรายงานนี้"

    try:
        # 3. เรียก fetch โดยส่ง filter_condition เข้าไปด้วย (ถ้า AI วิเคราะห์มาให้)
        df = fetch_and_generate_excel(clean_table_id, where_clause=filter_condition)
        
        if df is None:
            return f"❌ ไม่พบข้อมูลในรายงาน {report_name} " + (f"ตามเงื่อนไขที่ระบุ" if filter_condition else "")

        # 4. สร้าง Link ดาวน์โหลด (ใช้ clean_table_id ให้ตรงกับชื่อไฟล์ที่บันทึก)
        file_name = f"Report_{clean_table_id}.xlsx"
        download_url = f"{PUBLIC_BASE_URL.rstrip('/')}/download/{file_name}"

        return (
            f"✅ ระบบตรวจสอบสิทธิ์เรียบร้อยแล้ว\n"
            f"📊 ได้จัดเตรียม **รายงาน{report_name}** จำนวน {len(df)} แถว\n"
            f"💡 เงื่อนไขที่ใช้: {'ทั้งหมด' if not filter_condition else filter_condition}\n\n"
            f"📥 **ดาวน์โหลดที่นี่:**\n🔗 {download_url}"
        )
    except Exception as e:
        return f"🚨 เกิดข้อผิดพลาด: {str(e)}"

# Mount ที่ root ("") เพื่อให้ route ข้างในที่เป็น /mcp ยังคงเป็น /mcp ไม่เบิ้ลเป็น /mcp/mcp
app.mount("", mcp.streamable_http_app())

@app.on_event("startup")
async def startup_event():
    # FastMCP streamable-http จำเป็นต้องรัน session_manager
    app.state.mcp_session_manager_ctx = mcp.session_manager.run()
    await app.state.mcp_session_manager_ctx.__aenter__()

@app.on_event("shutdown")
async def shutdown_event():
    ctx = getattr(app.state, "mcp_session_manager_ctx", None)
    if ctx is not None:
        await ctx.__aexit__(None, None, None)

# --------------------------------------------------------
# REST API
# --------------------------------------------------------
@app.post("/generate_excel_report")
def generate_excel_report(request: GenerateReportRequest, http_request: Request):
    # 1. จัดการเรื่อง Table ID ให้เป็นมาตรฐานเดียวกัน (Clean Table ID)
    raw_table_id = request.table_id
    # ดึงเอาเฉพาะชื่อตารางท้ายสุด เช่น 'VRptExpension'
    clean_table_id = raw_table_id.split('.')[-1] 

    # 2. ตรวจสอบ Email (เน้น Header ก่อนเสมอ)
    user_email = http_request.headers.get("x-user-email") or request.user_email or DEFAULT_USER_EMAIL

    # 3. Validation (ควรใช้ clean_table_id ในการเช็ค)
    if not validate_table_id(clean_table_id):
        return {"success": False, "message": f"❌ ชื่อตาราง '{clean_table_id}' ไม่ถูกต้องในระบบ"}

    # 4. Check Permission (ส่ง clean_table_id เข้าไปเพื่อให้ Mapping ทำงานได้)
    report_name = check_user_permission(user_email, clean_table_id)
    if not report_name:
        return {
            "success": False, 
            "message": f"🙏 ขออภัยคุณ {user_email} ระบบไม่พบสิทธิ์เข้าถึงรายงานนี้ในฐานข้อมูล"
        }

    try:
        # 5. Fetch ข้อมูล (ขั้นตอนนี้อาจจะใช้ raw_table_id หรือ Full Path ตามที่ BigQuery ต้องการ)
        # แนะนำให้ fetch_and_generate_excel รับ Full Path ได้เพื่อความแม่นยำใน BigQuery
        full_table_path = f"sc-ai-uat.SCReport.{clean_table_id}"
        df = fetch_and_generate_excel(full_table_path)
        
        if df is None or df.empty:
            return {"success": False, "message": f"❌ ไม่พบข้อมูลในรายงาน {report_name}"}

        # 6. ตั้งชื่อไฟล์ให้สื่อความหมาย (ใช้ report_name)
        # ลบช่องว่างหรืออักขระพิเศษในชื่อไฟล์ออก
        safe_report_name = report_name.replace(" ", "_").replace("(", "").replace(")", "")
        file_name = f"Report_{safe_report_name}.xlsx"
        
        base_url = PUBLIC_BASE_URL
        download_url = f"{base_url.rstrip('/')}/download/{file_name}" if base_url else f"/download/{file_name}"

        return {
            "success": True, 
            "message": f"✅ จัดเตรียม **{report_name}** สำเร็จแล้วครับ\n🔗 [คลิกที่นี่เพื่อดาวน์โหลดรายงาน]({download_url})"
        }
    except Exception as e:
        print(f"🚨 Error in generate_excel_report: {str(e)}")
        return {"success": False, "message": f"🚨 เกิดข้อผิดพลาดขณะสร้างไฟล์: {str(e)}"}

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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)