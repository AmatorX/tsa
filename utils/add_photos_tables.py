import datetime
import calendar


def create_color_format_request(sheet_id, start_row, end_row, start_column, end_column, color):
    """
    Создает запрос на форматирование фона для заданного диапазона ячеек.

    :param sheet_id: Идентификатор листа
    :param start_row: Начальный индекс строки (0-based)
    :param end_row: Конечный индекс строки (не включается)
    :param start_column: Начальный индекс столбца (0-based)
    :param end_column: Конечный индекс столбца (не включается)
    :param color: Словарь с ключами 'red', 'green', 'blue' для определения цвета фона
    :return: Словарь с запросом на форматирование
    """

    return {
        'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': start_row,
                'endRowIndex': end_row,
                'startColumnIndex': start_column,
                'endColumnIndex': end_column
            },
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': color
                }
            },
            'fields': 'userEnteredFormat.backgroundColor'
        }
    }


def create_photos_table(service, spreadsheet_url, month, user_name='Example'):
    # service = get_service()
    spreadsheet_id = spreadsheet_url.split("/")[5]

    year = datetime.datetime.now().year  # Текущий год

    # Формирование названия листа
    sheet_name = f"photos_{month}_{year}"

    # Получение списка листов в документе
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск листа по названию
    sheet_id = None
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            sheet_id = sheet['properties']['sheetId']
            break

    # Проверка, найден ли лист
    if sheet_id is None:
        raise ValueError(f"Sheet with name '{sheet_name}' not found in the spreadsheet.")

    # Определение количества дней в месяце
    month_number = datetime.datetime.strptime(month, '%B').month
    days_in_month = calendar.monthrange(year, month_number)[1]

    # Создание данных таблицы
    values = [["Day"] + [user_name] + [''] * (25)]  # Заголовки таблицы, до столбца Z
    values += [[str(day)] + [''] * 25 for day in range(1, days_in_month + 1)]  # Дни месяца и пустые ячейки

    # Добавление данных в лист
    range_name = f"{sheet_name}!A1"
    body = {'values': values}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=range_name, body=body, valueInputOption='USER_ENTERED'
    ).execute()

    # Создание запроса на добавление границ

    requests = [{
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': 0,
                # 'endRowIndex': days_in_month + 3,
                'endRowIndex': days_in_month + 1,
                'startColumnIndex': 0,
                'endColumnIndex': 26,  # До столбца Z включительно
            },
            'top': {'style': 'SOLID', 'width': 1},
            'bottom': {'style': 'SOLID', 'width': 1},
            'left': {'style': 'SOLID', 'width': 1},
            'right': {'style': 'SOLID', 'width': 1},
            'innerHorizontal': {'style': 'SOLID', 'width': 1},
            'innerVertical': {'style': 'SOLID', 'width': 1},
        }
    }]

    # Применение запросов форматирования
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={'requests': requests}
    ).execute()

    return f"User table for {month} created for {user_name} in sheet '{sheet_name}'"


# spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1gKsdz-icvXkx7km6Dl5RIqcEkGhn9GmFp80BIt3kVco/edit#gid=0'
# month = 'January'

# create_photos_table(spreadsheet_url, month)
