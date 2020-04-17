from __future__ import print_function

from datetime import datetime
from google.oauth2 import service_account
from googleapiclient import discovery

filename_datetime_format = '%Y%m%d_%H%M%S'
generated_content_folder_id = '1xhk_PfBuXjib-nQUtCo_266RkD2k6cFR'
datetime_now = datetime.now()

def create_a_copy(unique_src_filename_substr, dest_file_folder_id):
    # Get ID of source file
    response = DRIVE.files().list(q="name contains '%s'" % unique_src_filename_substr).execute()
    src_file = response.get('files')[0]
    
    # Generate payload for new file
    dest_file_name = f'{datetime_now.strftime(filename_datetime_format)}_HealthCheckReport'
    payload = {'name': dest_file_name,
               'parents': [dest_file_folder_id]}

    # Copy file
    print('** Copying template %r as %r' % (src_file['name'], dest_file_name))
    dest_file_id = DRIVE.files().copy(body=payload, fileId=src_file['id']).execute().get('id')
    print(f'** Successfully made a copy | FILE NAME: {dest_file_name}  |  FILE ID: {dest_file_id}')

    return dest_file_id

IMG_FILE = 'google-slides.png'
REPORT_TEMPLATE = 'HealthCheckReport_Template'
REPORT_CHART_TEMPLATE = 'HealthCheckReport_Charts'
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

# Duplicate main report Google Slides
# REPORT_ID = create_a_copy(REPORT_TEMPLATE, generated_content_folder_id)

# Duplicate report charts Google Sheet
CHARTS_ID = create_a_copy(REPORT_CHART_TEMPLATE, generated_content_folder_id)


# Update report charts

fill_values = {
    'COMPANY_NAME': 'Canva',
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

data = []
for key, value in chart_fill_values.items():
    value_range = {
        "range": f'{key}!B2:B3',
        "majorDimension": 'COLUMNS',
        "values": [[fill_values[value[0]], value[1]]]
    }
    data.append(value_range)

# body = {
#   "range": "B2:B3",
#   "majorDimension": "COLUMNS",
#   "values": [
#     [29, 32]
#   ]
# }

body = {
    "data": data,
    "valueInputOption": 'USER_ENTERED'
}


SHEETS.spreadsheets().values().batchUpdate(
    spreadsheetId=CHARTS_ID,
    body=body
).execute()





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
# print('** Replacing placeholder text and icon')


# reqs = []
# for key, value in fill_values.items():
#     req_object = \
#         {
#             'replaceAllText':
#                 {
#                     'containsText': {'text': '{{' + key + '}}'},
#                     'replaceText': value
#                 }
#         }
#     reqs.append(req_object)

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
# SLIDES.presentations().batchUpdate(body={'requests': reqs},
#                                    presentationId=DECK_ID).execute()
print('DONE')