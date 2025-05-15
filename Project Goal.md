🧑‍💻 Project Goal
Automate a routine on a Windows 11 machine that:

1. Launches Refinitiv Workspace (data bridge)
2. Opens Excel with a pre-built macro for data processing
3. Executes a Python script to process/export CSV files
4. Uploads the CSVs into a browser-based custom web application via a web form

📚 Tech Stack & Tools Used
| Component | Technology/Tool |
| --------- | --------------- |
| Desktop Automation | Power Automate Desktop / Task Scheduler |
| Data Processing | Excel + VBA macros |
| Script Execution | Python (with subprocess, pandas) |
| File Upload to Web App | Python + selenium (headless browser) |
| Error Logging & Retry Logic | Python logging + conditional flows |
| Scheduling | Windows Task Scheduler |


🧭 Step-by-Step Process
1. Launch Data Bridge
Use Power Automate Desktop or Python + os.system to start the Refinitiv Workspace .exe:
```python
import os
os.startfile("C:\\Program Files\\Refinitiv\\Workspace\\workspace.exe")
```

2. Launch Excel and Run Macro
Launch Excel and open the .xlsm file with a macro:

```python
import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(r"C:\Automation\Reports\macro_report.xlsm")
excel.Visible = False
excel.Application.Run("macro_report.xlsm!Module1.MainMacro")
wb.Save()
wb.Close()
excel.Quit()
```


3. Run Python Script for Data Handling
Script reads from the Excel-generated file, processes data using pandas, and exports CSV:

```python
import pandas as pd

df = pd.read_excel("report_output.xlsx")
df_filtered = df[df['Status'] == 'Active']
df_filtered.to_csv("final_output.csv", index=False)
```

4. Automate Web Upload
Use selenium to log in and upload the CSV:

```python
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

driver = webdriver.Chrome()
driver.get("https://your-web-app.com/login")

driver.find_element(By.ID, "username").send_keys("youruser")
driver.find_element(By.ID, "password").send_keys("yourpass")
driver.find_element(By.ID, "login-btn").click()

time.sleep(2)
driver.get("https://your-web-app.com/upload")

upload_field = driver.find_element(By.NAME, "file")
upload_field.send_keys(r"C:\Automation\final_output.csv")
driver.find_element(By.ID, "submit-btn").click()

driver.quit()
```

5. Optional Logging
Add logging to track the process:

```python
import logging
logging.basicConfig(filename='automation_log.txt', level=logging.INFO)
logging.info("Upload completed successfully.")
```

✅ Outcome
This project:

- Automates multi-app workflows on a dedicated Windows 11 machine
- Requires no user interaction once triggered
- Handles failure cases with retry logic and logging
