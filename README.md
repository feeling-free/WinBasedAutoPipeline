# Windows Automation Pipeline

This project automates a workflow that includes launching Refinitiv Workspace, processing Excel data, and uploading results to a web application.

## Prerequisites

- Windows 11
- Python 3.8 or higher
- Refinitiv Workspace installed
- Google Chrome (for web automation)
- Excel with macro capabilities

## Setup

1. Clone this repository to your local machine.

2. Create a virtual environment and activate it:
```bash
python -m venv venv
.\venv\Scripts\activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

4. Copy `.env.template` to `.env` and fill in your credentials:
```bash
cp .env.template .env
```

5. Update the paths in `config.py` to match your system:
- REFINITIV_PATH
- EXCEL_MACRO_PATH
- Other paths as needed

## Usage

1. Ensure all applications (Refinitiv Workspace, Excel) are closed before running the automation.

2. Run the automation:
```bash
python automation.py
```

The script will:
- Launch Refinitiv Workspace
- Run Excel macro for data processing
- Generate CSV output
- Upload results to web application
- Log all activities in `automation_log.txt`

## Error Handling

- The script includes retry logic for all major operations
- Errors are logged to `automation_log.txt`
- Each operation will retry up to 3 times with exponential backoff

## Scheduling

To schedule this automation:

1. Open Windows Task Scheduler
2. Create a new task
3. Set the trigger (time/frequency)
4. Action: Start a program
   - Program/script: `python`
   - Arguments: `automation.py`
   - Start in: `[path-to-script-directory]`

## Troubleshooting

1. Check `automation_log.txt` for detailed error messages
2. Ensure all paths in `config.py` are correct
3. Verify web application credentials in `.env`
4. Confirm Refinitiv Workspace and Excel are properly installed

## Support

For issues and questions, please check the logs first and then create an issue in the repository. 