import os
import time
import logging
from retry import retry
import pandas as pd
import win32com.client
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from config import *

# Configure logging
logging.basicConfig(
    filename=LOG_FILE_PATH,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class AutomationPipeline:
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    @retry(tries=MAX_RETRIES, delay=RETRY_DELAY, backoff=2)
    def launch_refinitiv(self):
        """Launch Refinitiv Workspace"""
        try:
            self.logger.info("Launching Refinitiv Workspace...")
            os.startfile(REFINITIV_PATH)
            time.sleep(10)  # Wait for application to start
            self.logger.info("Refinitiv Workspace launched successfully")
        except Exception as e:
            self.logger.error(f"Error launching Refinitiv: {str(e)}")
            raise

    @retry(tries=MAX_RETRIES, delay=RETRY_DELAY, backoff=2)
    def run_excel_macro(self):
        """Launch Excel and run the macro"""
        try:
            self.logger.info("Starting Excel and running macro...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(EXCEL_MACRO_PATH)
            excel.Application.Run(f"{os.path.basename(EXCEL_MACRO_PATH)}!{MACRO_NAME}")
            wb.Save()
            wb.Close()
            excel.Quit()
            self.logger.info("Excel macro completed successfully")
        except Exception as e:
            self.logger.error(f"Error in Excel macro execution: {str(e)}")
            raise

    @retry(tries=MAX_RETRIES, delay=RETRY_DELAY, backoff=2)
    def process_data(self):
        """Process data using pandas"""
        try:
            self.logger.info("Processing data with pandas...")
            df = pd.read_excel(EXCEL_OUTPUT_PATH)
            df_filtered = df[df['Status'] == 'Active']
            df_filtered.to_csv(CSV_OUTPUT_PATH, index=False)
            self.logger.info("Data processing completed successfully")
        except Exception as e:
            self.logger.error(f"Error in data processing: {str(e)}")
            raise

    @retry(tries=MAX_RETRIES, delay=RETRY_DELAY, backoff=2)
    def upload_to_webapp(self):
        """Upload processed file to web application"""
        try:
            self.logger.info("Starting web upload process...")
            chrome_options = Options()
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")

            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            wait = WebDriverWait(driver, 10)

            # Login to web application
            driver.get(f"{WEBAPP_URL}/login")
            wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(WEBAPP_USERNAME)
            wait.until(EC.presence_of_element_located((By.ID, "password"))).send_keys(WEBAPP_PASSWORD)
            wait.until(EC.element_to_be_clickable((By.ID, "login-btn"))).click()

            # Navigate to upload page and submit file
            driver.get(f"{WEBAPP_URL}/upload")
            upload_field = wait.until(EC.presence_of_element_located((By.NAME, "file")))
            upload_field.send_keys(os.path.abspath(CSV_OUTPUT_PATH))
            wait.until(EC.element_to_be_clickable((By.ID, "submit-btn"))).click()

            driver.quit()
            self.logger.info("Web upload completed successfully")
        except Exception as e:
            self.logger.error(f"Error in web upload: {str(e)}")
            raise

    def run_pipeline(self):
        """Execute the complete automation pipeline"""
        try:
            self.logger.info("Starting automation pipeline...")
            self.launch_refinitiv()
            self.run_excel_macro()
            self.process_data()
            self.upload_to_webapp()
            self.logger.info("Automation pipeline completed successfully")
        except Exception as e:
            self.logger.error(f"Pipeline failed: {str(e)}")
            raise

if __name__ == "__main__":
    pipeline = AutomationPipeline()
    pipeline.run_pipeline() 