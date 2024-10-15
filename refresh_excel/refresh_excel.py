import win32com.client as win32
import logging
import traceback
import time

logging.basicConfig(
    filename=r'C:\Users\emerson\Downloads\projects\script-neelevat\logs.txt', 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def wait_for_query_refresh_to_complete(workbook):
    """Wait until all queries/connections are done refreshing."""
    try:
        while True:
            refreshing = False
            for connection in workbook.Connections:
                try:
                    if connection.OLEDBConnection and connection.OLEDBConnection.Refreshing:
                        refreshing = True
                except AttributeError:
                    continue
            
            if not refreshing:
                break

            logging.info("Waiting for query refresh to complete...")
            time.sleep(2)

    except Exception as e:
        logging.error(f"Error while waiting for query refresh: {e}")
        raise
    
def main():
    try:
        excel_app = win32.DispatchEx('Excel.Application') # Open Excel Application
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        workbook = excel_app.Workbooks.Open(r'C:\Users\emerson\Downloads\projects\script-neelevat\data.xlsx')

        workbook.RefreshAll()

        wait_for_query_refresh_to_complete(workbook)
        
        workbook.Save()
        workbook.Close()
        excel_app.Quit()

        del workbook
        del excel_app
        logging.info("Script ran successfully! - data refreshed")

    except Exception as e:
        logging.error("An error occurred: %s", str(e))
        logging.error(traceback.format_exc())


if __name__ == '__main__':
    main()