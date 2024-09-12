from googleapiclient.errors import HttpError

from common.service import service

service = service.get_service()


def add_tool_to_sheet(sh_url, instance):
    try:
        spreadsheet_id = sh_url.split('/')[-2]
        sheet_name = 'Sheet1'
        # sheet_name = 'Лист1'

        # Получаем текущее количество строк в таблице
        sheet_data = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A:F'
        ).execute()

        num_existing_rows = len(sheet_data.get('values', []))

        # Определяем диапазон для записи новых данных
        start_row_index = num_existing_rows
        end_row_index = start_row_index + 1  # Только одна строка будет добавлена

        sheet_range = f'{sheet_name}!A{start_row_index + 1}:D{start_row_index + 1}'
        if instance:
            price = str(instance.price)
            # Данные для записи
            values = [
                [
                    instance.tool_id,
                    instance.name,
                    price,
                    instance.date_of_issue.strftime('%Y-%m-%d') if instance.date_of_issue else '',
                    instance.assigned_to.name if instance.assigned_to else '',
                    instance.assigned_to.build_obj.name if instance.assigned_to and instance.assigned_to.build_obj else '',
                ]
            ]

            # Запись данных в Google Sheets
            body = {
                'values': values
            }

            service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=sheet_range,
                valueInputOption='RAW',
                body=body
            ).execute()

            # Определяем диапазон для обведения границ
            borders_range = {
                "sheetId": 0,  # Идентификатор листа, возможно, нужно будет уточнить.
                "startRowIndex": start_row_index,
                "endRowIndex": end_row_index,
                "startColumnIndex": 0,  # Столбец A
                "endColumnIndex": 6,
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







