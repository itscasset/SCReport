import os
import smtplib
import asyncio
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.encoders import encode_base64
import pandas as pd
from google.cloud import bigquery

# 📦 เรียกใช้แค่แกนหลักของ MCP สากล (ไม่ต้องใช้ FastAPI)
from mcp.server import Server
import mcp.server.stdio
import mcp.types as types

# 1. เริ่มต้นระบบ MCP Server แกนหลัก
mcp_server = Server("sc-report-server")

# ฟังก์ชันภายใน: ส่งอีเมลพร้อมแนบไฟล์ Excel อัตโนมัติ
def send_excel_email(file_path: str, file_name: str, to_email: str, table_id: str):
    smtp_host = "smtp.office365.com"   # 📝 ปรับเปลี่ยนตาม SMTP Server ของบริษัท
    smtp_port = 587
    sender_email = "bot-report@scasset.com"  # 📝 อีเมลผู้ส่งของระบบ
    sender_password = "your-app-password"    # 📝 รหัสผ่านแอปพลิเคชัน

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = f"[SC Report] รายงานข้อมูลตาราง {table_id}"

    body = f"เรียนพนักงาน SC Asset,\n\nระบบได้ดำเนินการดึงข้อมูลจาก BigQuery ตาราง {table_id} และแปลงเป็นไฟล์ Excel เรียบร้อยแล้ว รายละเอียดปรากฏตามไฟล์แนบครับ\n\nขอบคุณครับ"
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={file_name}")
        msg.attach(part)

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)

# 2. ลงทะเบียนประกาศเครื่องมือ (Tools Registration)
@mcp_server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="generate_excel_report",
            description="ดึงข้อมูลจาก Google BigQuery ตามชื่อตารางที่ระบุ แปลงเป็นไฟล์ Excel (.xlsx) และส่งเข้าอีเมลโดยตรง",
            inputSchema={
                "type": "object",
                "properties": {
                    "table_id": {"type": "string", "description": "ชื่อตารางใน BigQuery เช่น AuthenByMenu"},
                    "requester_email": {"type": "string", "description": "อีเมลพนักงานที่ระบุให้จัดส่ง"}
                },
                "required": ["table_id", "requester_email"],
            },
        )
    ]

@mcp_server.call_tool()
async def handle_call_tool(name: str, arguments: dict | None) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
    if name != "generate_excel_report":
        raise ValueError(f"Unknown tool: {name}")

    if not arguments:
        raise ValueError("Missing arguments")

    table_id = arguments.get("table_id")
    requester_email = arguments.get("requester_email")

    try:
        # ดึงข้อมูลจาก BigQuery
        client = bigquery.Client(project="sc-ai-uat")
        query = f"SELECT * FROM `sc-ai-uat.SCReport.{table_id}` LIMIT 2000"
        df = client.query(query).to_dataframe()
        
        if df.empty:
            return [types.TextContent(type="text", text=f"❌ ไม่พบข้อมูลในตาราง {table_id} บน BigQuery")]

        # สร้างไฟล์ Excel บนเซิร์ฟเวอร์ชั่วคราว
        file_name = f"Report_{table_id}.xlsx"
        df.to_excel(file_name, index=False, engine='openpyxl')

        # เรียกใช้ฟังก์ชันส่งเมล
        send_excel_email(file_path=file_name, file_name=file_name, to_email=requester_email, table_id=table_id)

        if os.path.exists(file_name):
            os.remove(file_name)

        return [types.TextContent(type="text", text=f"✅ ดึงข้อมูลตาราง {table_id} สำเร็จ และส่งไฟล์ Excel ไปที่เมล {requester_email} เรียบร้อยแล้ว!")]

    except Exception as e:
        return [types.TextContent(type="text", text=f"🚨 เกิดข้อผิดพลาดในระบบหลังบ้าน: {str(e)}")]

# 3. ฟังก์ชันหลักสำหรับรันระบบผ่านท่อตรง (Stdio Transport)
async def main():
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await mcp_server.run(
            read_stream,
            write_stream,
            mcp_server.create_initialization_options()
        )

if __name__ == "__main__":
    asyncio.run(main())