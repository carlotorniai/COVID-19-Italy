## Here daving the excel file into google docs
from apiclient.discovery  import build  
from httplib2 import Http  
from oauth2client import file, client, tools  
from oauth2client.contrib import gce  
from apiclient.http import MediaFileUpload

CLIENT_SECRET = "./client_secret.json"

SCOPES='https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets'  
store = file.Storage('token.json')  
creds = store.get()  
if not creds or creds.invalid:  
    flow = client.flow_from_clientsecrets(CLIENT_SECRET, SCOPES)
    creds = tools.run_flow(flow, store)
SERVICE = build('drive', 'v3', http=creds.authorize(Http()))  
SS_SERVICE = build('sheets', 'v4', http=creds.authorize(Http()))