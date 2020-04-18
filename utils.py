from datetime import datetime


def create_a_copy(unique_src_filename_substr, dest_folder_id, drive_api):

    filename_datetime_format = '%Y%m%d_%H%M%S'
    datetime_now = datetime.now()

    # TODO: Add error handling

    # Get ID of source file
    response = drive_api.files().list(q="name contains '%s'" % unique_src_filename_substr).execute()

    # TODO: check file count
    src_file = response.get('files')[0]

    # Generate payload for new file
    dest_file_name = f'{datetime_now.strftime(filename_datetime_format)}_HealthCheckReport'
    payload = {'name': dest_file_name,
               'parents': [dest_folder_id]}

    # Copy file
    print('** Copying template %r as %r' % (src_file['name'], dest_file_name))
    dest_file_id = drive_api.files().copy(body=payload, fileId=src_file['id']).execute().get('id')
    print(f'** Successfully made a copy | FILE NAME: {dest_file_name}  |  FILE ID: {dest_file_id}')

    return dest_file_id


def generate_slide_requests(fill_values):
    reqs = []
    # TODO: Update charts
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

    return reqs


def generate_sheet_requests(chart_fill_values, fill_values):
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
