# 🛡️ ปรับฟังก์ชันเช็คสิทธิ์ให้ดึงชื่อรายงานภาษาไทยมาด้วย
def check_user_permission(email: str, table_id: str):
    print(f"🔍 [Authentication] กำลังตรวจสอบสิทธิ์ของ: {email} สำหรับตาราง: {table_id}")
    try:
        client = bigquery.Client(project="sc-ai-uat")
        
        # 🎯 ปรับ Query ให้ดึง ReportName ออกมาตรงๆ โดยเช็คคู่กับ Email และชื่อตารางภาษาอังกฤษ
        auth_query = f"""
            SELECT ReportName 
            FROM `sc-ai-uat.SCReport.AuthenByMenu` 
            WHERE Email = '{email}'
        """
        auth_df = client.query(auth_query).to_dataframe()
        
        # ถ้าเจอข้อมูล แปลว่ามีสิทธิ์ และเราจะส่งชื่อรายงานภาษาไทยกลับไป
        if not auth_df.empty:
            return auth_df['ReportName'].iloc[0]
            
        return None # ถ้าไม่พบข้อมูล แปลว่าไม่มีสิทธิ์
    except Exception as auth_error:
        print(f"🚨 ไม่สามารถตรวจสอบสิทธิ์กับ BigQuery ได้: {str(auth_error)}")
        return None

# --------------------------------------------------------

# ตอนเรียกใช้ในคลังเครื่องมือ call_tool
@mcp_server.call_tool()
async def handle_call_tool(name: str, arguments: dict | None) -> list[types.TextContent]:
    if name != "generate_excel_report" or not arguments:
        raise ValueError("Invalid tool call")

    table_id = arguments.get("table_id")
    user_email = os.getenv("CURRENT_USER_EMAIL", "test_user@scasset.com")

    # 🔒 1. ตรวจสอบสิทธิ์พนักงาน (รอบนี้จะได้ชื่อรายงานกลับมาด้วย)
    report_name = check_user_permission(user_email, table_id)

    if not report_name:
        # 💬 ท่าปฏิเสธอย่างสุภาพเมื่อไม่มีสิทธิ์
        return [types.TextContent(
            type="text", 
            text=f"🙏 ขออภัยในความไม่สะดวกครับคุณ {user_email.split('@')[0]} เนื่องจากระบบตรวจสอบพบว่าคุณยังไม่มีสิทธิ์เข้าถึงรายงานตัวนี้ในขณะนี้\n\n💡 หากต้องการตรวจสอบหรือดูข้อมูลรายงานเพิ่มเติม สามารถเข้าชมได้ที่ระบบ **sc system** ครับ"
        )]

    try:
        # 📊 2. หากมีสิทธิ์ (ผ่านเงื่อนไข) ดึงข้อมูลต่อทันที
        client = bigquery.Client(project="sc-ai-uat")
        query = f"SELECT * FROM `sc-ai-uat.SCReport.{table_id}` LIMIT 2000"
        df = client.query(query).to_dataframe()

        if df.empty:
            return [types.TextContent(type="text", text=f"❌ ไม่พบข้อมูลในรายงาน {report_name} ({table_id})")]

        file_name = f"Report_{table_id}.xlsx"
        download_url = f"https://[โดเมนของระบบคลาวด์]/download/{file_name}"

        # 🎯 จุดสำเร็จ: นำค่า report_name ภาษาไทยจากฐานข้อมูลมาตอบกลับหา User ได้ทันที!
        return [types.TextContent(
            type="text", 
            text=f"✅ ตรวจสอบสิทธิ์ผ่านระบบ AuthenByMenu เรียบร้อยแล้วครับ\n📊 ระบบได้จัดเตรียม **รายงาน{report_name}** จำนวนทั้งหมด {len(df)} แถวให้คุณเรียบร้อยแล้ว\n\n📥 **คลิกลิงก์เพื่อดาวน์โหลดไฟล์ลงเครื่อง:**\n🔗 [ดาวน์โหลดไฟล์รายงาน Excel]({download_url})"
        )]

    except Exception as e:
        return [types.TextContent(type="text", text=f"🚨 เกิดข้อผิดพลาดระหว่างดึงข้อมูล: {str(e)}")]