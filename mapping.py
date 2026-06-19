import re

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

SENSITIVE_COLUMNS = {"username", "email"}

TABLE_TO_BQ_TABLE_NAME: dict[str, str] = {
    "vrptexpension": "VRptExpension",
    "vrptexpensionexpmodule": "VRptExpensionExpModule",
}

ALLOWED_OPERATORS: set[str] = {"=", ">", "<", ">=", "<=", "LIKE", "!="}

# รองรับ Username ที่มีช่องว่าง เช่น "Rattanachote Petpansri" และอักษรไทย
USER_REGEX = re.compile(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$|^[a-zA-Z0-9\u0E00-\u0E7F _.-]{2,80}$")
BQ_TABLE_REGEX = re.compile(r"^[A-Za-z0-9_.-]+$")

# --------------------------------------------------------
# NOTE: Column mapping (TABLE_COLUMN_MAPPING) ถูกลบออกแล้ว
# ดึงข้อมูล Column Mapping แบบ Dynamic จาก Firestore แทน
# ผ่านฟังก์ชัน get_table_column_mapping() ใน screportserver.py
# Firestore path: data_dictionary/{DATASET_ID}/tables/{bq_table_name}
# --------------------------------------------------------
