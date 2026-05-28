import asyncio
import pandas as pd
from google.cloud import bigquery
import os

# 🎯 จำลองการทำงานของฟังก์ชันดึงข้อมูลแบบไม่ต้องผ่านท่อ MCP
async def test_generate_report(table_id: str):
    print(f"🔄 กำลังเริ่มดึงข้อมูลจาก BigQuery ตาราง: {table_id} ...")
    try:
        # 1. เรียกใช้ Client เชื่อมต่อไปยัง BigQuery
        client = bigquery.Client(project="sc-ai-uat")
        query = f"SELECT * FROM `sc-ai-uat.SCReport.{table_id}` LIMIT 2000"
        
        # 2. ยิงคิวรีและแปลงเป็น DataFrame
        df = client.query(query).to_dataframe()
        
        if df.empty:
            print(f"❌ ไม่พบข้อมูลในตาราง {table_id} บน BigQuery")
            return

        # 3. สั่งสร้างและ Auto Save ไฟล์ Excel ลงเครื่องตรงๆ
        file_name = f"Report_{table_id}.xlsx"
        absolute_path = os.path.abspath(file_name)
        df.to_excel(absolute_path, index=False, engine='openpyxl')

        print("\n==================================================")
        print(f"✅ สำเร็จแล้วครับคุณ Kachatharn!")
        print(f"📂 บันทึกไฟล์สำเร็จที่: {absolute_path}")
        print(f"📊 ดึงข้อมูลมาได้ทั้งหมด: {len(df)} แถว")
        print("==================================================")

    except Exception as e:
        print(f"\n🚨 เกิดข้อผิดพลาดระหว่างทดสอบ: {str(e)}")
        print("💡 แนะนำให้เช็คว่าในเครื่องมีไฟล์ JSON Key ของ BigQuery และเซ็ตค่าสิทธิ์ถูกต้องหรือยังนะครับ")

# สั่งให้สคริปต์เริ่มทำงานทดสอบ (ระบุชื่อตารางที่ต้องการเทสตรงนี้ได้เลย)
if __name__ == "__main__":
    # 📝 สามารถเปลี่ยนชื่อตารางตรง 'AuthenByMenu' เป็นตารางอื่นที่ต้องการเทสได้ครับ
    asyncio.run(test_generate_report("AuthenByMenu"))