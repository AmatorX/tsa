import datetime
from time import sleep

from .add_results_tables import create_materials_table
from .add_photos_tables import create_photos_table


# def create_monthly_sheets(service, spreadsheet_url, materials):
#     # service = get_service()
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#     year = datetime.datetime.now().year
#
#     month_names = ["January", "February", "March", "April", "May", "June",
#                    "July", "August", "September", "October", "November", "December"]
#
#     requests = []
#
#     for i, month in enumerate(month_names, start=1):
#         # Создаем лист для фотографий
#         photos_sheet_title = f"photos_{month}_{year}"
#         photos_request = {
#             'addSheet': {
#                 'properties': {
#                     'title': photos_sheet_title,
#                     'gridProperties': {
#                         'rowCount': 1000,
#                         'columnCount': 26
#                     }
#                 }
#             }
#         }
#         requests.append(photos_request)
#
#         # Создаем лист для результатов
#         results_sheet_title = f"results_{month}_{year}"
#         results_request = {
#             'addSheet': {
#                 'properties': {
#                     'title': results_sheet_title,
#                     'gridProperties': {
#                         'rowCount': 1000,
#                         'columnCount': 26
#                     }
#                 }
#             }
#         }
#         requests.append(results_request)
#
#     # Добавляем все листы одним запросом
#     body = {'requests': requests}
#     service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
#
#     for month in month_names:
#         photos_sheet_title = f"photos_{month}_{year}"
#         results_sheet_title = f"results_{month}_{year}"
#
#         create_materials_table(service, spreadsheet_url, month, materials)
#         create_photos_table(service, spreadsheet_url, month)
#
#     return f"Monthly sheets for {year} created in spreadsheet: {spreadsheet_url}"


def create_monthly_sheets(service, spreadsheet_url, materials):
    spreadsheet_id = spreadsheet_url.split("/")[5]
    year = datetime.datetime.now().year

    all_month_names = ["January", "February", "March", "April", "May", "June",
                       "July", "August", "September", "October", "November", "December"]
    current_month = datetime.datetime.now().month
    month_names = all_month_names[current_month - 1:]

    # Получение списка существующих листов в таблице
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    existing_sheets = [sheet['properties']['title'] for sheet in sheet_metadata.get('sheets', '')]

    requests = []

    for i, month in enumerate(month_names, start=1):
        # Создаем лист для фотографий, если он еще не существует
        photos_sheet_title = f"photos_{month}_{year}"
        if photos_sheet_title not in existing_sheets:
            photos_request = {
                'addSheet': {
                    'properties': {
                        'title': photos_sheet_title,
                        'gridProperties': {
                            'rowCount': 1000,
                            'columnCount': 26
                        }
                    }
                }
            }
            requests.append(photos_request)

        # Создаем лист для результатов, если он еще не существует
        results_sheet_title = f"results_{month}_{year}"
        if results_sheet_title not in existing_sheets:
            results_request = {
                'addSheet': {
                    'properties': {
                        'title': results_sheet_title,
                        'gridProperties': {
                            'rowCount': 1000,
                            'columnCount': 26
                        }
                    }
                }
            }
            requests.append(results_request)

    # Добавляем все новые листы одним запросом, если есть что добавлять
    if requests:
        body = {'requests': requests}
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    for month in month_names:
        photos_sheet_title = f"photos_{month}_{year}"
        results_sheet_title = f"results_{month}_{year}"

        if photos_sheet_title not in existing_sheets:
            create_photos_table(service, spreadsheet_url, month)

        if results_sheet_title not in existing_sheets:
            create_materials_table(service, spreadsheet_url, month, materials)
        sleep(0.2)
    return f"Monthly sheets for {year} checked/created in spreadsheet: {spreadsheet_url}"

# create_monthly_sheets('https://docs.google.com/spreadsheets/d/1gKsdz-icvXkx7km6Dl5RIqcEkGhn9GmFp80BIt3kVco/edit#gid=0')
