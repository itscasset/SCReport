import requests

# 1. ที่อยู่ของเซิร์ฟเวอร์
url = "http://localhost:8080/generate_chart_report"

# 2. ข้อมูล Request Payload (สำหรับส่งให้ FastAPI)
# หมายเหตุ: table_id ต้องเป็นตารางที่คุณมีสิทธิ์ในระบบ UAT เช่น vrptexpension
payload = {
    "table_id": "vrptexpension",
    "limit": 1000,
    "code": """
import seaborn as sns
import matplotlib.pyplot as plt

# ตารางข้อมูลจะถูกแปลงคอลัมน์เป็นชื่อภาษาไทยให้เรียบร้อยแล้ว
# คุณสามารถเรียกใช้งานคอลัมน์ชื่อไทยอย่าง 'ชื่อกลุ่มงาน' ได้ทันที
if not df.empty:
    sns.scatterplot(
        data=df, 
        x='มูลค่างานหลัง Vat', 
        y='มูลค่างาน (รวม VAT) - มูลค่าสุทธิ'
    )
    plt.title("Scatter Plot: Value vs Value After Vat")
""",
    # อีเมลผู้ใช้งานที่ต้องใช้เช็คสิทธิ์การเข้าถึงข้อมูลตาราง
    "user_email": ""
}

# 3. กำหนด Headers (ระบุ email ยืนยันตัวตน)
headers = {
    "current-user": "",
    "Content-Type": "application/json"
}

print(f"🚀 กำลังส่ง HTTP POST request ไปที่: {url}...")
try:
    response = requests.post(url, json=payload, headers=headers)
    print("\n--- ผลลัพธ์จากเซิร์ฟเวอร์ (API Response) ---")
    print(f"HTTP Status Code: {response.status_code}")
    print("Response JSON:")
    import json
    print(json.dumps(response.json(), indent=2, ensure_ascii=False))
except Exception as e:
    print(f"\n🚨 ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้: {e}")
    print("💡 คำแนะนำ: ตรวจสอบว่าได้รันเซิร์ฟเวอร์ `python screportserver.py` หรือยังนะครับ")
