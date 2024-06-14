import datetime
import calendar


def create_color_format_request(sheet_id, start_row, end_row, start_column, end_column, color):
    """
    Создает запрос на форматирование фона для заданного диапазона ячеек.

    :param sheet_id: Идентификатор листа
    :param start_row: Начальный индекс строки
    :param end_row: Конечный индекс строки (не включается)
    :param start_column: Начальный индекс столбца
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


def add_formulas_to_table(values, days_in_month, materials_start_col=1):  # Изменено на 1 для столбца "B"
    formula_start_row = len(values) - 3  # Индекс строки для "Total materials"

    # Добавление формул для каждого столбца начиная со "B"
    for col_index in range(materials_start_col, 26):  # Проходим по столбцам начиная со столбца "B"
        column_letter = chr(65 + col_index)  # Преобразуем индекс столбца в букву

        # Формула для "Total materials" для текущего столбца
        sum_formula = f'=SUM({column_letter}6:{column_letter}{5 + days_in_month})'
        values[formula_start_row][col_index] = sum_formula

        # Формула для "Earned" для текущего столбца
        price_cell = f'{column_letter}3'  # Цена за единицу в столбце "B"
        earned_formula = f'={price_cell}*{column_letter}{formula_start_row + 1}'
        values[formula_start_row + 1][col_index] = earned_formula

    # Формула для "Total earned"
    total_earned_formula = f"=SUM({chr(65 + materials_start_col)}{formula_start_row + 2}:{chr(64 + 26)}{formula_start_row + 2})"
    values[formula_start_row + 2][1] = total_earned_formula  # Использование индекса 1 для столбца "B"

    return values


def create_materials_table(service, spreadsheet_url, month, materials, user_name='Example'):
    # service = get_service()

    # Извлечение spreadsheet_id из URL
    spreadsheet_id = spreadsheet_url.split("/")[5]

    # Получение текущего года
    year = datetime.datetime.now().year

    # Формирование названия листа
    sheet_name = f"results_{month}_{year}"

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
    month_number = list(calendar.month_name).index(month)
    days_in_month = calendar.monthrange(year, month_number)[1]

    # Создание данных таблицы
    values = [
        ["Name", user_name],
        ["Month", month],
        ["Price"] + [material[1] for material in materials],
        ["Material"] + [material[0] for material in materials],
        ["Day"],
    ] + [[str(day)] + [''] * (26 - 1) for day in range(1, days_in_month + 1)]  # 26 столбцов до Z включительно

    # Добавление трех строк для подсчетов
    values += [
        ['Total materials'] + [''] * 25,
        ['Earned'] + [''] * 25,
        ['Total earned'] + [''] * 25
    ]

    # Добавление формул
    values = add_formulas_to_table(values, days_in_month)

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
                'endRowIndex': days_in_month + 5 + 3,  # +5 для учета заголовков и строки "Day", +3 для формул в конце
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

    return f"Materials table with borders created in sheet '{sheet_name}', extended to column Z with two empty rows at the end"



# spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1gKsdz-icvXkx7km6Dl5RIqcEkGhn9GmFp80BIt3kVco/edit#gid=0'
# month = 'January'
# materials = [['Wood', 100], ['Metal', 200], ['Plastic', 50]]

# create_materials_table(spreadsheet_url, month, materials)
