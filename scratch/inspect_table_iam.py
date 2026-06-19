import sys
from google.cloud import bigquery

client = bigquery.Client(project="sc-ai-uat")

def inspect_policy(table_id):
    print(f"Inspecting IAM Policy for {table_id}...")
    try:
        table = client.get_table(table_id)
        # BigQuery Python client table policy method:
        # Note: BigQuery tables support IAM policies.
        # We can try to get IAM policy using the client:
        # client.get_iam_policy(table)
        policy = client.get_iam_policy(table.reference)
        print(f"Bindings for {table_id}:")
        for binding in policy.bindings:
            print(f"  Role: {binding['role']}")
            print(f"  Members: {binding['members']}")
    except Exception as e:
        print(f"Error getting IAM policy: {e}")

if __name__ == "__main__":
    inspect_policy("sc-ai-uat.SCReport.AiLog")
