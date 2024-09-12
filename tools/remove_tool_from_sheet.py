from googleapiclient.errors import HttpError

from common.get_service import get_service


service = get_service()


# def remove_tool_from_sheet(sheet_url, instance):
#     try:
#         spreadsheet_id = sheet_url.split('/')[-2]
#         row_index = find_row_with_id(spreadsheet_id, instance.tool_id)
#
#         if not row_index:
#             print(f"ID {instance.tool_id} not found.")
#             return
#
#         # Конвертируем индекс строки в диапазон для удаления
#         sheet_name = 'Лист1'  # Убедитесь, что имя листа указано правильно
#         # sheet_id = get_sheet_id(spreadsheet_id, sheet_name)  # Получить идентификатор листа
#
#         # Запрос на удаление строки
#         requests = [
#             {
#                 "deleteDimension": {
#                     "range": {
#                         "sheetId": spreadsheet_id,
#                         "dimension": "ROWS",
#                         "startIndex": row_index - 1,  # Индекс строки в запросе начинается с 0
#                         "endIndex": row_index  # Индекс строки в запросе не включительно
#                     }
#                 }
#             }
#         ]
#
#         # Выполнение запроса на удаление строки
#         service.spreadsheets().batchUpdate(
#             spreadsheetId=spreadsheet_id,
#             body={'requests': requests}
#         ).execute()
#
#         print(f"Row {row_index} successfully deleted.")
#
#     except HttpError as error:
#         print(f"An error occurred: {error}")


def remove_tool_from_sheet(sheet_url, instance):
    try:
        # Извлечение идентификатора таблицы из URL
        spreadsheet_id = sheet_url.split('/')[-2]

        # Нахождение индекса строки, которую нужно удалить
        row_index = find_row_with_id(spreadsheet_id, instance.tool_id)
        if not row_index:
            print(f"ID {instance.tool_id} not found.")
            return

        # Имя листа и идентификатор листа (0 для первого листа)
        # sheet_name = 'Лист1'
        sheet_name = 'Sheet1'
        sheet_id = 0  # Используйте 0 для первого листа

        # Запрос на удаление строки
        requests = [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": row_index - 1,  # Индекс строки начинается с 0
                        "endIndex": row_index  # Индекс строки не включительно
                    }
                }
            }
        ]

        # Выполнение запроса на удаление строки
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()

        print(f"Row {row_index} successfully deleted.")

    except HttpError as error:
        print(f"An error occurred: {error}")


def find_row_with_id(spreadsheet_id, search_id):
    try:
        # Имя листа
        sheet_name = 'Sheet1'
        # sheet_name = 'Лист1'
        # Диапазон для чтения данных (все ячейки столбца A)
        # range_name = f'{sheet_name}!A:A'
        range_name = f"{sheet_name}!A1:A"

        # Получение данных из Google Sheets
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()

        values = result.get('values', [])

        # Поиск строки с нужным ID
        for index, row in enumerate(values):
            if row and row[0] == search_id:
                # Возвращаем номер строки (индексация с 1)
                return index + 1

        # Возвращаем None, если ID не найден
        return None

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None
