from __future__ import print_function

from datetime import datetime
from google.oauth2 import service_account
from googleapiclient import discovery
from utils import create_a_copy, generate_slide_requests, generate_sheet_requests


datetime_now = datetime.now()

generated_content_folder_id = '1xhk_PfBuXjib-nQUtCo_266RkD2k6cFR'  # Pre-created Google Drive file

REPORT_TEMPLATE = 'HealthCheckReport_Template'
REPORT_CHART_TEMPLATE = 'HealthCheckReport_Charts'

# Google Auth
SCOPES = (
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/spreadsheets',
)

SERVICE_ACCOUNT_FILE = './service-account-key.json'
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

DRIVE = discovery.build('drive',  'v3', credentials=credentials)
SLIDES = discovery.build('slides', 'v1', credentials=credentials)
SHEETS = discovery.build('sheets', 'v4', credentials=credentials)

# 1. Duplicate main report Google Slides
REPORT_ID = create_a_copy(REPORT_TEMPLATE, generated_content_folder_id, DRIVE)

# 2. Duplicate report charts Google Sheet
CHARTS_ID = create_a_copy(REPORT_CHART_TEMPLATE, generated_content_folder_id, DRIVE)

# 3. Get values from form (currently hardcoded)
# TODO: Get from form

fill_values = {
    'COMPANY_NAME': 'Random Company',
    'MONTH': datetime_now.strftime('%B'),
    'YEAR': datetime_now.strftime('%Y'),
    'DATE': datetime_now.strftime('%d %B %Y'),
    'USER': 'TestUser',
    'TCH_WMN': '14',
    'TCH_WMN_EVAL': 'low',
    'ENG_WMN': '4',
    'INF_WMN': '6',
    'DAT_WMN': '20',
    'PRD_WMN': '9',
    'SEC_WMN': '0',
    'TECHLEAD_WMN': '2',
    'LDRSHIP_WMN': '25',
    'SNR_LDRSHIP_WMN': '50',
    'PRMY_CARER_LEAVE_WKS': '8',
    'SEC_CARER_LEAVE_WKS': '2',
    'QLFYNG_PERIOD_MNTHS': '12',
    'TCH_MEN': '86'
}

chart_fill_values = {
    'ENG': ['ENG_WMN', 50],
    'DATA': ['DAT_WMN', 50],
    'PRODUCT': ['PRD_WMN', 50],
    'INFRA': ['INF_WMN', 50],
    'SECURITY': ['SEC_WMN', 50],
    'TECH_LEADS': ['TECHLEAD_WMN', 50],
    'TECH_LEADERSHIP': ['LDRSHIP_WMN', 50],
    'CO_LEADERSHIP_TEAM': ['SNR_LDRSHIP_WMN', 50],
    'OVERALL_TECH': ['TCH_WMN', 'TCH_MEN']
}

# 4. Update Google Sheet chart values
print('4. Update Google Sheet chart values')
body = generate_sheet_requests(chart_fill_values, fill_values)
SHEETS.spreadsheets().values().batchUpdate(spreadsheetId=CHARTS_ID, body=body).execute()

# 5. Update Google Slides
print('5. Update Google Slides')
reqs = generate_slide_requests(fill_values)
SLIDES.presentations().batchUpdate(body={'requests': reqs}, presentationId=REPORT_ID).execute()

print('Process completed!')

# print('** Get slide objects, search for image placeholder')
# slide = SLIDES.presentations().get(presentationId=DECK_ID,
#         fields='slides').execute().get('slides')[0]
# obj = None
# for obj in slide['pageElements']:
#     if obj['shape']['shapeType'] == 'RECTANGLE':
#         break
#
# print('** Searching for icon file')
# rsp = DRIVE.files().list(q="name='%s'" % IMG_FILE).execute().get('files')[0]
# print(' - Found image %r' % rsp['name'])
# img_url = '%s&access_token=%s' % (
#         DRIVE.files().get_media(fileId=rsp['id']).uri, creds.access_token)
#
# # reqs = [
# #     # {'createImage': {
# #     #     'url': img_url,
# #     #     'elementProperties': {
# #     #         'pageObjectId': slide['objectId'],
# #     #         'size': obj['size'],
# #     #         'transform': obj['transform'],
# #     #     }
# #     # }},
# #     # {'deleteObject': {'objectId': obj['objectId']}},
# # ]