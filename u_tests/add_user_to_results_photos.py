from service import get_service
# from u_tests.get_start_end_indexes_rows import get_start_end_table
from get_start_end_indexes_rows import get_start_end_table


def copy_and_paste_table(spreadsheet_url, sheet_name, usernames):
    service = get_service()
    spreadsheet_id = spreadsheet_url.split("/")[5]
    
    sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)
    start_copy, end_copy, start_paste, end_paste = get_start_end_table(service, spreadsheet_id, sheet_name)
    for username in usernames:
        # start_copy, end_copy, start_paste, end_paste = get_start_end_table(service, spreadsheet_id, sheet_name)
        print(f"start_copy: {start_copy}, end_copy: {end_copy}, start_paste: {start_paste}, end_paste: {end_paste}")

        requests = []

        # Добавление запроса на копирование и вставку диапазона
        requests.append({
            "copyPaste": {
                "source": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_copy,
                    "endRowIndex": end_copy,
                    "startColumnIndex": 0,
                    "endColumnIndex": 26
                },
                "destination": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_paste,
                    "endRowIndex": end_paste,
                    "startColumnIndex": 0,
                    "endColumnIndex": 26
                },
                "pasteType": "PASTE_NORMAL",
                "pasteOrientation": "NORMAL"
            }
        })

        # Добавление запроса на изменение значения в ячейке с именем пользователя
        requests.append({
            "updateCells": {
                "rows": [{
                    "values": [{"userEnteredValue": {"stringValue": username}}]
                }],
                "fields": "userEnteredValue",
                "start": {
                    "sheetId": sheet_id,
                    "rowIndex": start_paste,
                    "columnIndex": 1
                }
            }
        })

        # Выполнение запросов
        body = {
            "requests": requests
        }
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        start_paste = end_paste + 4
        end_paste = start_paste + end_copy

    return response

def get_sheet_id(service, spreadsheet_id, sheet_name):
    # Получение информации о всех листах в таблице
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск идентификатора листа по имени
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']

    raise ValueError(f"Sheet with name '{sheet_name}' not found.")



spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1JFbB5TucFZn-xl4NtTVpcYAvqP89DZmOy5xxB18rBsI/edit#gid=299250498'
sheet_name = 'results_January_2024'
usernames = ['Pavel', 'Misha', 'Vlad']

copy_and_paste_table(spreadsheet_url, sheet_name, usernames)