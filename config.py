import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Application paths
REFINITIV_PATH = r"C:\Program Files\Refinitiv\Workspace\workspace.exe"
EXCEL_MACRO_PATH = r"C:\Automation\Reports\macro_report.xlsm"
EXCEL_OUTPUT_PATH = "report_output.xlsx"
CSV_OUTPUT_PATH = "final_output.csv"
LOG_FILE_PATH = "automation_log.txt"

# Web application settings
WEBAPP_URL = os.getenv("WEBAPP_URL", "https://your-web-app.com")
WEBAPP_USERNAME = os.getenv("WEBAPP_USERNAME", "")
WEBAPP_PASSWORD = os.getenv("WEBAPP_PASSWORD", "")

# Excel macro settings
MACRO_NAME = "Module1.MainMacro"

# Retry settings
MAX_RETRIES = 3
RETRY_DELAY = 5  # seconds 