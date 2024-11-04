from googleapiclient.errors import HttpError

from common.get_service import get_service


service = get_service()


def update_tool_user(sheet_url, instance):
    print('Start >>>>>>> update_tool_user')
    try:
        spreadsheet_id = sheet_url.split('/')[-2]
        start_row_index = find_row_with_id(spreadsheet_id, instance.tool_id)
        if not start_row_index:
            print(f"ID {instance.tool_id} not found.")
            return

        # Конвертируем индекс строки в диапазон для обновления
        # sheet_name = 'Лист1'
        sheet_name = 'Sheet1'
        sheet_range = f'{sheet_name}!A{start_row_index}:F{start_row_index}'
        if instance:
            price = str(instance.price)
            # Данные для записи, с пустым значением в колонке Assigned To
            values = [
                [
                    instance.tool_id,
                    instance.name,
                    price,
                    instance.date_of_issue.strftime('%Y-%m-%d') if instance.date_of_issue else '',
                    instance.assigned_to.name if instance.assigned_to else '',  # Присвоить пустое значение, чтобы указать, что Assigned To теперь пусто
                    instance.assigned_to.build_obj.name if instance.assigned_to and instance.assigned_to.build_obj else '',
                ]
            ]

            # Подготовка тела запроса для записи данных
            body = {
                'values': values
            }

            # Запись данных в Google Sheets
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=sheet_range,
                valueInputOption='RAW',
                body=body
            ).execute()

    except HttpError as error:
        print(f"An error occurred: {error}")


def find_row_with_id(spreadsheet_id, search_id):
    print('Start >>>>>>> find_row_with_id')
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


def find_rows_with_name(spreadsheet_id, search_name):
    print(f'Start >>>>>>>> find_rows_with_name')
    try:
        sheet_name = 'Sheet1'
        # sheet_name = 'Лист1'

        # Диапазон для чтения данных (все ячейки столбца B)
        range_name = f'{sheet_name}!D1:D'

        # Получение данных из Google Sheets
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()

        values = result.get('values', [])

        # Список для хранения найденных индексов строк
        matching_rows = []

        # Поиск строк с нужным именем
        for index, row in enumerate(values):
            if row and row[0] == search_name:
                # Добавляем индекс строки (индексация с 1) в список
                matching_rows.append(index + 1)

        # Возвращаем список найденных строк
        return matching_rows

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


def update_building_in_users_rows(sheet_url, name, building):
    print(f'Start >>>>>>>>>> update_building_in_users_rows')
    sheet_name = 'Sheet1'
    # sheet_name = 'Лист1'
    spreadsheet_id = sheet_url.split('/')[-2]
    target_rows = find_rows_with_name(spreadsheet_id, name)
    # Обновляем значение в 5-й ячейке (столбец E) для каждой строки
    for row in target_rows:
        try:
            range_name = f'{sheet_name}!E{row}'

            # Обновляем значение в Google Sheets
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body={'values': [[building]]}
            ).execute()

            print(f'Строка {row}: значение успешно обновлено.')

        except HttpError as error:
            print(f'An error occurred while updating row {row}: {error}')
