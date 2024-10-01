import win32com.client as win32
import logging
import traceback

logging.basicConfig(
    filename=r'C:\Users\emerson\Downloads\projects\script-neelevat\logs.txt', 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def main():
    try:
        excel_app = win32.Dispatch('Excel.Application') # Open Excel Application
        excel_app.Visible = False  # Set to True for debugging
        excel_app.DisplayAlerts = False

        workbook = excel_app.Workbooks.Open(r'C:\Users\emerson\Downloads\projects\script-neelevat\data.xlsx')

        workbook.RefreshAll()
        workbook.Save()
        workbook.Close()

        excel_app.Quit()
        logging.info("Script ran successfully! - data refreshed")

    except Exception as e:
        logging.error("An error occurred: %s", str(e))
        logging.error(traceback.format_exc())


if __name__ == '__main__':
    main()