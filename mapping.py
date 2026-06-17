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
# ✅ รองรับ Username ที่มีช่องว่าง เช่น "Rattanachote Petpansri" และอักษรไทย
USER_REGEX = re.compile(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$|^[a-zA-Z0-9\u0E00-\u0E7F _.-]{2,80}$")
BQ_TABLE_REGEX = re.compile(r"^[A-Za-z0-9_.-]+$")
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
