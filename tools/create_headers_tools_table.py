from googleapiclient.errors import HttpError

from common.get_service import get_service

service = get_service()


def create_headers_tools_table(tools_sheet_instance):
    try:
        # Извлечение идентификатора таблицы из URL
        spreadsheet_id = tools_sheet_instance.sh_url.split('/')[-2]
        # sheet_name = 'Лист1'
        sheet_name = 'Sheet1'

        # Определите диапазон для заголовков
        headers_range = f'{sheet_name}!A1:F1'

        # Заголовки
        headers = [
            ['Tool ID', 'Name', 'Price', 'Date of Issue', 'Assigned To', 'Building object']
        ]

        # Подготовка тела запроса для записи заголовков
        body = {
            'values': headers
        }

        # Запись заголовков в Google Sheets
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=headers_range,
            valueInputOption='RAW',
            body=body
        ).execute()

        # Определяем диапазон для обведения границ
        borders_range = {
            "sheetId": 0,  # Идентификатор листа, может потребоваться корректировка
            "startRowIndex": 0,
            "endRowIndex": 1,
            "startColumnIndex": 0,  # Столбец A
            "endColumnIndex": 6  # Столбец D (индекс последнего столбца не включительно)
        }

        # Запрос на обведение ячеек линией
        requests = [{
            "updateBorders": {
                "range": borders_range,
                "top": {"style": "SOLID"},
                "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"},
                "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"},
                "innerVertical": {"style": "SOLID"}
            }
        }]

        # Выполнение запроса на обновление границ
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()

    except HttpError as error:
        print(f"An error occurred: {error}")

