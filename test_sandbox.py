import pandas as pd
from graph_mcp import run_visualization_code
from pathlib import Path

# 1. จำลองข้อมูล (DataFrame) ที่จะเอาไปวาดกราฟ
df = pd.DataFrame({
    'ชื่อกลุ่มงาน': ['โครงสร้าง', 'สถาปัตย์', 'ระบบไฟ', 'ประปา'],
    'มูลค่างาน (รวม VAT) - มูลค่าสุทธิ': [120000.0, 95000.0, 45000.0, 30000.0]
})

# 2. กำหนดโค้ดไวยากรณ์ที่จะให้รัน (สามารถแก้โค้ดตรงนี้เพื่อเทสแบบต่าง ๆ ได้)
code = """
print("This is test for print")
"""

output_path = Path("chart_output.png").resolve()
print(f"🔄 กำลังรันโค้ดและเซฟผลลัพธ์ไปที่: {output_path}...")

# 3. สั่งรันโค้ดผ่านระบบ Sandbox
res = run_visualization_code(df, code, str(output_path))

print("\n--- ผลการทดสอบ (Test Result) ---")
print("Success:", res["success"])
if res["success"]:
    print("✅ สำเร็จ! ไฟล์รูปภาพสร้างขึ้นที่:", output_path)
    print("ขนาดไฟล์ภาพ:", output_path.stat().st_size, "bytes")
else:
    print("❌ เกิดข้อผิดพลาด:")
    print(res["message"])
