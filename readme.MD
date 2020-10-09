###############################
Excel To Google Spread Sheet
###############################

Basic aim of this program is to take a backup of Excel File to Google Sheet.
Not a lot of free programs are available at market for doing the same. It is very dificult to convert an Excel File to Google Sheet automatically, for daily use.

################################
Structure of the program
################################

This program is started with reading a local excel file, then create a google sheet service, clear the previous data in the google sheet, append data from the excel file to the google sheet. If everything done successfully run crete a log file named LOGS.txt for store status and last date and time of backup done. If program gives error store status also in LOGS.txt

################################
How to run the program
################################

Download the repositry as a zip file. Extract in a folder

Main program is stored in excelToSheet.py
supply correct information for the required variables in the file excelToSheet.py

IP_ADDR = None # if required the change the value else default is None
PORT_NO = None # if required the change the value else default is None

# json file must be in the same folder where Google.py and this file exists.
CLIENT_SECRET_FILE = '<client_secret_file_name credentials in json from google>'

# pass value through variables for Google Sheets.
GSPREADSHEETS_ID = '<Google Spreadsheet ID>'
SHEETS_NAME = '<name of the google sheet>'
SHEETS_CELL_NAME = '<Cell Name of Google Sheet from where file need to be added>'

# pass value through variables for Excel Sheet.
EXCEL_FILE = '<Excel File Name local excel file>'
OPEN_PASSWORD = None # if required the change the value else default is None
WRITE_PASSWORD = None # if required the change the value else default is None
WORKSHEET_NAME = '<Excel Worksheet Name>'
EXCEL_CELL_NAME = '<Excel Worksheet Cell Name from where records started>'

Now run the python file with the command 'python excelToSheet.py' in windows operating system, in linux 'python3 excelToSheet.py'

To run this program in background after a certain time, you can use Task Scheduler of Windows.
Any suggestion for improvement is always welcome.


################################
Known Issues
################################

Sometimes giving error "Unable to find server at www.googleapis.com" while working behind a proxy.
Any suggestion for removing the issue is always welcome.

################################
Improvement
################################

Please give your valueable feedback to improve this program