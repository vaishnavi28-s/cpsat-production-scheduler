"""
Airflow DAG — Print Job Combo Optimizer

Runs twice daily (06:00 and 14:00 UTC) to:
  1. Fetch print jobs from Snowflake
  2. Run CP-SAT optimization
  3. Export color-coded Excel report
  4. Upload report to SharePoint

Environment variables required (set in Airflow Connections or .env):
  SF_ACCOUNT, SF_USER, SF_PASSWORD, SF_OTP, SF_ROLE,
  SF_WAREHOUSE, SF_DATABASE, SF_SCHEMA,
  SHAREPOINT_SITE_URL, SHAREPOINT_FOLDER, SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD
"""

import os
import sys
from datetime import datetime, timedelta

from airflow import DAG
from airflow.operators.python import PythonOperator

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.data_loader import load_from_snowflake, load_from_csv
from src.optimizer import run_optimizer
from src.export import export_to_excel

OUTPUT_PATH = "/tmp/combo_output.xlsx"
SAMPLE_CSV = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "sample_jobs.csv")

DEFAULT_ARGS = {
    "owner": "ai-team",
    "depends_on_past": False,
    "email_on_failure": True,
    "email_on_retry": False,
    "retries": 1,
    "retry_delay": timedelta(minutes=5),
}


def fetch_and_optimize(**context):
    """Fetch jobs from Snowflake, run optimizer, export to Excel."""
    use_snowflake = os.getenv("USE_SNOWFLAKE", "true").lower() == "true"

    if use_snowflake:
        jobs = load_from_snowflake(limit=5000)
    else:
        jobs = load_from_csv(SAMPLE_CSV)

    print(f"Loaded {len(jobs)} jobs")
    total_runs = run_optimizer(jobs)
    export_to_excel(total_runs, jobs, path=OUTPUT_PATH)
    context["ti"].xcom_push(key="output_path", value=OUTPUT_PATH)
    print(f"Exported results to {OUTPUT_PATH}")


def upload_to_sharepoint(**context):
    """Upload the Excel report to SharePoint."""
    output_path = context["ti"].xcom_pull(key="output_path", task_ids="fetch_and_optimize")

    sharepoint_url = os.getenv("SHAREPOINT_SITE_URL")
    sharepoint_folder = os.getenv("SHAREPOINT_FOLDER", "Shared Documents/PrintJobReports")
    username = os.getenv("SHAREPOINT_USERNAME")
    password = os.getenv("SHAREPOINT_PASSWORD")

    if not all([sharepoint_url, username, password]):
        print("SharePoint credentials not configured — skipping upload.")
        return

    try:
        from office365.runtime.auth.user_credential import UserCredential
        from office365.sharepoint.client_context import ClientContext

        ctx = ClientContext(sharepoint_url).with_credentials(
            UserCredential(username, password)
        )

        file_name = os.path.basename(output_path)
        target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder)

        with open(output_path, "rb") as f:
            target_folder.upload_file(file_name, f.read()).execute_query()

        print(f"Uploaded {file_name} to SharePoint: {sharepoint_folder}")

    except ImportError:
        print("office365-rest-python-client not installed. Skipping SharePoint upload.")
    except Exception as e:
        print(f"SharePoint upload failed: {e}")
        raise


with DAG(
    dag_id="print_job_optimizer",
    default_args=DEFAULT_ARGS,
    description="Optimize print job combinations and export to SharePoint",
    schedule_interval="0 6,14 * * *",
    start_date=datetime(2025, 1, 1),
    catchup=False,
    tags=["print", "optimization", "or-tools"],
) as dag:

    task_optimize = PythonOperator(
        task_id="fetch_and_optimize",
        python_callable=fetch_and_optimize,
        provide_context=True,
    )

    task_upload = PythonOperator(
        task_id="upload_to_sharepoint",
        python_callable=upload_to_sharepoint,
        provide_context=True,
    )

    task_optimize >> task_upload
