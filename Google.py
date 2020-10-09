import pickle
import os.path
import datetime
import socks
import httplib2
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


def create_service(ip_addr, port_no, client_secret_file, api_name, api_version, *scopes):
    IP_ADDR = ip_addr
    PORT_NO = port_no
    if IP_ADDR is not None:
        socks.set_default_proxy(socks.PROXY_TYPE_HTTP,IP_ADDR, PORT_NO)
        socks.wrap_module(httplib2)
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = [scope for scope in scopes[0]]
    CLIENT_SECRET_FILE = client_secret_file
    API_NAME = api_name
    API_VERSION = api_version

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    try:
        service = build(API_NAME, API_VERSION, credentials=creds)
        print(API_NAME, "service created successfully.")
        return service
    except Exception as e:
        print("Unable to connect.")
        print(e)
        return None

def dateconverter(o):
    if isinstance(o, datetime.datetime):
        return o.strftime("%d-%m-%Y")