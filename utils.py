from datetime import datetime

datetime_now = datetime.now()

fill_values = {
    'COMPANY_NAME': 'Random Company',
    'MONTH': datetime_now.strftime('%B'),
    'YEAR': datetime_now.strftime('%Y'),
    'DATE': datetime_now.strftime('%d %B %Y'),
    'USER': 'TestUser',
    'TCH_WMN': '14',
    'TCH_WMN_EVAL': 'low',
    'ENG_WMN': '4',
    'INF_WMN': '0',
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
    'SECURITY': ['SEC_WMN', 0],
    'TECH_LEADS': ['TECHLEAD_WMN', 50],
    'TECH_LEADERSHIP': ['LDRSHIP_WMN', 50],
    'CO_LEADERSHIP_TEAM': ['SNR_LDRSHIP_WMN', 50],
    'OVERALL_TECH': ['TCH_WMN', 50]
}

charts = {
    "g5784b994ba_3_0": {
        'chartId': 715383549,  # Security Chart
        'transform':
        {
            "scaleX": 105,
            "scaleY": 105,
            "translateX": 1822649.67,
            "translateY": 4691775,
            "unit": "EMU"
        }
    },
    "g5784b994ba_0_0": {
        'chartId': 1334235346,  # Infrastructure Chart
        'transform':
        {
            "scaleX": 105,
            "scaleY": 105,
            "translateX": 3865363,
            "translateY": 1978925,
            "unit": "EMU"
        }
    },
    "g63b241fd6a_0_0": {
        'chartId': 353279936,  # Data Chart
        'transform':
        {
            "scaleX": 105,
            "scaleY": 105,
            "translateX": 5908077,
            "translateY": 1978925,
            "unit": "EMU"
        }
    },
    "g54cefca8fe_0_87": {
        'chartId': 607026404,  # Co Senior Leadership Team Chart
        'transform':
        {
            "scaleX": 105,
            "scaleY": 105,
            "translateX": 7950791,
            "translateY": 4691775,
            "unit": "EMU"
        }
    },
    "g54cefca8fe_0_88": {
        'chartId': 557365308,  # Tech Leadership Chart
        'transform':
        {
            "scaleX": 105,
            "scaleY": 105,
            "translateX": 5908077,
            "translateY": 4691775,
            "unit": "EMU"
        }
    },
    "g54cefca8fe_0_89": {
        'chartId': 1658241655,  # Tech Leads Chart
        'transform':
        {
            "scaleX": 105,
            "scaleY": 105,
            "translateX": 3865363,
            "translateY": 4691775,
            "unit": "EMU"
        }
    },
    "g54cefca8fe_0_90": {
        'chartId': 1365614958,  # Engineering Chart
        'transform':
        {
            "scaleX": 105,
            "scaleY": 105,
            "translateX": 1822649,
            "translateY": 1978925,
            "unit": "EMU"
        }
    },
    "g54cefca8fe_0_92": {
        'chartId': 513287435,  # Product Management Chart
        'transform':
        {
            "scaleX": 105,
            "scaleY": 105,
            "translateX": 7950791,
            "translateY": 1978925,
            "unit": "EMU"
        }
    },

}

reqs = []



def create_a_copy(file_id, dest_folder_id, drive_api):

    filename_datetime_format = '%Y%m%d_%H%M%S'
    datetime_now = datetime.now()

    # TODO: Add error handling
    # Generate payload for new file
    dest_file_name = f'{datetime_now.strftime(filename_datetime_format)}_HealthCheckReport'
    payload = {'name': dest_file_name,
               'parents': [dest_folder_id]}

    # Copy file
    print('** Copying template %r as %r' % (file_id, dest_file_name))
    dest_file_id = drive_api.files().copy(body=payload, fileId=file_id).execute().get('id')
    print(f'** Successfully made a copy | FILE NAME: {dest_file_name}  |  FILE ID: {dest_file_id}')

    return dest_file_id


def generate_slide_requests(new_spreadsheet_id):
    reqs = []

    for key, value in fill_values.items():
        req_object = \
            {
                'replaceAllText':
                    {
                        'containsText': {'text': '{{' + key + '}}'},
                        'replaceText': value
                    }
            }
        reqs.append(req_object)

    for object_id, chart in charts.items():
        req = \
            {
                "createSheetsChart": {
                    "objectId": object_id,
                    "elementProperties": {
                        "pageObjectId": "g54cefca8fe_0_86",
                        "size": {
                            "width": {
                                "magnitude": 30000,
                                "unit": "EMU"
                            },
                            "height": {
                                "magnitude": 18550,
                                "unit": "EMU"
                            }
                        },
                        "transform": chart['transform']
                    },
                  "spreadsheetId": new_spreadsheet_id,
                  "chartId": chart['chartId'],
                  "linkingMode": "LINKED"
                }
            }
        reqs.append(req)

    return reqs


def generate_sheet_requests():
    data = []
    for sheet, values in chart_fill_values.items():
        value_range = {
            "range": f'{sheet}!B2:B3',
            "majorDimension": 'COLUMNS',
            "values": [[fill_values[values[0]], values[1]]]
        }
        data.append(value_range)

    body = {
        "data": data,
        "valueInputOption": 'USER_ENTERED'
    }

    return body
