from Google import create_service, dateconverter
from extra import writeFile
import win32com.client as w32c
import json


'''
Set all the below mentioned variables as per your need.
'''
IP_ADDR = None # if required the change the value else default is None
PORT_NO = None # if required the change the value else default is None
# pass value through variables for Google.py to create a service
API_NAME = 'sheets'
API_VERSION = 'v4'
# json file must be in the same folder where Google.py and this file exists.
CLIENT_SECRET_FILE = '<client_secret_file_name credentials in json from google>'
# If modifying these scopes, delete the token.pickle file.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

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


'''
#############################################
#############################################
### Main Program is running from here. ######
#############################################
#############################################
'''
# workbook and worksheet variables
wb = None
ws = None
range_value = None

xlapp = w32c.gencache.EnsureDispatch('Excel.Application')
xlapp.Visible = False

print('Finding The Excel File')
# check if OPEN_PASSWORD is available
if OPEN_PASSWORD is not None:
    wb = xlapp.Workbooks.Open(EXCEL_FILE, Password= OPEN_PASSWORD, WriteResPassword= WRITE_PASSWORD, IgnoreReadOnlyRecommended=True)
else:
    wb = xlapp.Workbooks.Open(EXCEL_FILE)

if wb is not None:
    print('Reading Excel File')
    wb.RefreshAll
    ws = wb.Sheets(WORKSHEET_NAME)
    range_value = ws.Range(EXCEL_CELL_NAME).CurrentRegion()
    ws = None
    wb = None
    xlapp = None
    print('Converting to json')
    range_value = json.dumps(range_value, default=dateconverter)

# create a service for google
print('Creating Service for Google Sheets')
service = create_service(IP_ADDR, PORT_NO, CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

# if service is not none
if service is not None:
    try:
        print('Clearing Google Sheet')
        # connect with google sheets.
        response = service.spreadsheets().values().clear(
            spreadsheetId = GSPREADSHEETS_ID,
            range = SHEETS_NAME
        ).execute()
        print('Appending Data to Google Sheet')
        response = service.spreadsheets().values().append(
            spreadsheetId = GSPREADSHEETS_ID,
            valueInputOption = "RAW",
            range = SHEETS_NAME + "!" +SHEETS_CELL_NAME,
            body = dict(
                majorDimension = 'ROWS',
                values = json.loads(range_value)
            )
        ).execute()
        print('Backup Completed Successfully')
        writeFile('Success.')

        service = None
        response = None
    except Exception as e:
        print('Error! Backup Not Completed Successfully')
        print(e)
        writeFile('Error- ' + str(e))
        service = None
        response = None
else:
    writeFile('Error- Connection Problem')
