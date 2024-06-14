# from .get_service import get_service
from time import sleep

from .get_start_end_indexes_rows import get_start_end_table
from common.service import service


service = service.get_service()


def add_users_to_results_photos(spreadsheet_url, usernames):
    # service = get_service()
    spreadsheet_id = spreadsheet_url.split("/")[5]

    # Получение списка всех листов
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    for sheet in sheets:
        sheet_name = sheet['properties']['title']

        # Проверяем соответствие названия листа шаблону
        if sheet_name.startswith("results_") or sheet_name.startswith("photos_"):
            sheet_id = sheet['properties']['sheetId']

            start_copy, end_copy, start_paste, end_paste = get_start_end_table(service, spreadsheet_id, sheet_name)

            for username in usernames:
                try:
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
                    body = {"requests": requests}
                    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

                    start_paste = end_paste + 4
                    end_paste = start_paste + end_copy
                except Exception as e:
                    print(f"An error occurred: {e}")
            sleep(0.3)

    return "Tables copied and pasted successfully for all matching sheets."
