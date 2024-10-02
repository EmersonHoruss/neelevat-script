# Project Overview: Excel Data Refresh Script

This project consists of a Python script designed to automate the refreshing of data in an Excel workbook using the `win32com.client` library. The script opens an Excel application instance, refreshes the data connections, and saves the workbook. This automation is particularly useful for ensuring that the latest data is always available without manual intervention.

## Key Features
- Automates the process of opening an Excel workbook.
- Refreshes data connections within the workbook.
- Logs actions and errors for monitoring and debugging purposes.
- Closes the workbook and the Excel application cleanly after refreshing.

## Requirements

To successfully run the script, ensure you have the following installed:

1. **Python**: The script is compatible with Python 3.x. Make sure to download and install it from [python.org](https://www.python.org/downloads/).

2. **Required Libraries**: The script utilizes the `pywin32` package to interface with Excel. Install it using pip:
   ```bash 
   pip install pywin32
3. **Excel**: It is required to have excel installed, to have this it is necessary to be in a windows operative system.

4. **Change the path of files**: for the moment it is hardcoded the path, so, each time you want download the project you have to change the path.

## Integration with Task Scheduler
You can schedule the script to run every time you want. We achive this using Task Scheduler.

### Setup 
To run python script and make changes on excel files is necessary to activate the interactive user. You can achive this by following the next steps:
1. Write in the search bar Component Services.
2. Open Component Services Windows application.
3. Look for Microsoft Excel Application
4. Right click and select Properties
5. Click on Identity tab
6. Select: The interactive user option
