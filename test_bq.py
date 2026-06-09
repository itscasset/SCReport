from google.cloud import bigquery

client = bigquery.Client(project='sc-ai-uat')
query = "SELECT Email, ReportName FROM `sc-ai-uat.SCReport.AuthenByMenu` WHERE Email LIKE '%rattanachote%'"
job = client.query(query)

with open('db_results.txt', 'w', encoding='utf-8') as f:
    for row in job:
        f.write(f"Email: {repr(row['Email'])}, ReportName: {repr(row['ReportName'])}\n")
