from google.cloud import bigquery

PROJECT_ID = "sc-ai-uat"
DATASET_ID = "SCReport"
AUTH_TABLE = "AuthenByMenu"
bq_client = bigquery.Client(project=PROJECT_ID)

email = ""
report_names = [
    "รายงานตรวจสอบการจ่ายเงิน (ต้นทุน)",
    "รายงานตรวจสอบการจ่ายเงิน(ต้นทุน)",
    "ทะเบียนคุมการเบิกจ่าย (ต้นทุน)",
    "ทะเบียนคุมการเบิกจ่าย(ต้นทุน)",
]

placeholders = ", ".join([f"@name{i}" for i in range(len(report_names))])
query = f"""
    SELECT ReportName
    FROM `{PROJECT_ID}.{DATASET_ID}.{AUTH_TABLE}`
    WHERE LOWER(TRIM(Email)) = LOWER(TRIM(@email))
    AND TRIM(ReportName) IN ({placeholders})
"""
print("Query:", query)

params = [bigquery.ScalarQueryParameter("email", "STRING", email.strip())]
for i, name in enumerate(report_names):
    params.append(bigquery.ScalarQueryParameter(f"name{i}", "STRING", name))

job_config = bigquery.QueryJobConfig(query_parameters=params)
try:
    job = bq_client.query(query, job_config=job_config)
    results = list(job)
    print("Result length:", len(results))
    for row in results:
        print("Row:", repr(row['ReportName']))
except Exception as e:
    print("Error:", e)
