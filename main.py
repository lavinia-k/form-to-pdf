from __future__ import print_function
import json

from datetime import datetime
from google.oauth2 import service_account
from googleapiclient import discovery
from utils import create_a_copy, generate_slide_requests, generate_sheet_requests


datetime_now = datetime.now()

FOLDER_ID = '1xhk_PfBuXjib-nQUtCo_266RkD2k6cFR'  # Pre-created Google Drive file
TEMPLATE_PRESENTATION_ID = '1i3jbxK48AQmDMdYfdslPd3Emn7ydpCutiG9lVY32V8E'
TEMPLATE_SPREADSHEET_ID = '1Q0NtNj3dkznONXEv8y8xaQU5cHwjw7aOYjYyCIRNesQ'


# Google Auth #
SCOPES = (
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/spreadsheets',
)

SERVICE_ACCOUNT_FILE = './.secrets/service-account-key.json'
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

DRIVE = discovery.build('drive',  'v3', credentials=credentials)
SLIDES = discovery.build('slides', 'v1', credentials=credentials)
SHEETS = discovery.build('sheets', 'v4', credentials=credentials)
####################


# 1. Duplicate report charts Google Sheet
NEW_SPREADSHEET_ID = create_a_copy(TEMPLATE_SPREADSHEET_ID, FOLDER_ID, DRIVE)

# 2. Update Google Sheet chart values
print('Updating Google Sheet chart values')
body = generate_sheet_requests()
SHEETS.spreadsheets().values().batchUpdate(spreadsheetId=NEW_SPREADSHEET_ID, body=body).execute()
spreadsheet = SHEETS.spreadsheets().get(spreadsheetId=NEW_SPREADSHEET_ID).execute()
sheets = spreadsheet.get("sheets")

# 3. Duplicate main report Google Slides
NEW_PRESENTATION_ID = create_a_copy(TEMPLATE_PRESENTATION_ID, FOLDER_ID, DRIVE)

# 4. Update Google Slides
print('Updating Google Slides')
reqs = generate_slide_requests(NEW_SPREADSHEET_ID)
SLIDES.presentations().batchUpdate(body={'requests': reqs}, presentationId=NEW_PRESENTATION_ID).execute()

print('Process completed!')
