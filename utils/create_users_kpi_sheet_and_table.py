import calendar
from datetime import datetime
from time import sleep


def find_last_filled_row_in_column(service, spreadsheet_id, sheet_name, column="A"):
    """
    Функция поиска последней заполненной строки в столбце.
    Нужна для определения индекса начальной строки для создания таблицы
    """
    # Определение диапазона для поиска в столбце A
    range_name = f"{sheet_name}!{column}1:{column}1000"  # Предполагаем, что строк не будет более 1000

    # Получение данных из столбца A
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    # Поиск последней заполненной ячейки
    last_filled_row = len(values) + 1 if values else 1  # Если нет значений, начинаем с первой строки
    print(f'Последняя заполненная строка: {last_filled_row}')

    # Прибавляем 3 к индексу последней заполненной строки для отступа
    return last_filled_row + 3



def create_border_format_request(sheet_id, start_row, end_row, start_column, end_column):
    """
    Создает запрос на форматирование границ.

    :return: Словарь с запросом на форматирование
    """
    return {
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_column,
                    "endColumnIndex": end_column
                },
                "top": {"style": "SOLID", "width": 1},
                "bottom": {"style": "SOLID", "width": 1},
                "left": {"style": "SOLID", "width": 1},
                "right": {"style": "SOLID", "width": 1},
                "innerHorizontal": {"style": "SOLID", "width": 1},
                "innerVertical": {"style": "SOLID", "width": 1}
            }
        }


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


def get_sheet_id(service, spreadsheet_id, sheet_name):
    # Получение информации о всех листах в таблице
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск идентификатора листа по имени
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']

    raise ValueError(f"Sheet with name '{sheet_name}' not found.")


def create_users_kpi_sheet(service, spreadsheet_url, sheet_name):
    spreadsheet_id = spreadsheet_url.split("/")[5]

    # Получение списка существующих листов
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Проверка наличия листа с именем sheet_name
    if not any(sheet['properties']['title'] == sheet_name for sheet in sheets):
        # Лист с таким именем не найден, создаем новый с 52 столбцами
        body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                        'gridProperties': {
                            'rowCount': 1000,  # Укажите подходящее количество строк в зависимости от ваших потребностей
                            'columnCount': 52  # Указываем количество столбцов
                        }
                    }
                }
            }]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        print(f"Sheet '{sheet_name}' with 52 columns created.")
    else:
        print(f"Sheet '{sheet_name}' already exists.")


def get_column_letter(num):
    """Преобразует номер столбца в буквенное обозначение."""
    letter = ''
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


def create_users_kpi_table(service, spreadsheet_url, sheet_name='Users KPI'):
    current_day = datetime.now().day

    # Проверяем, что сегодня первое число месяца
    # if current_day != 1:
    #     print("Сегодня не первое число месяца, таблица для текущего месяца уже должна существовать.")
    #     return
    spreadsheet_id = spreadsheet_url.split("/")[5]

    # Проверяем существует ли лист, если нет - создаем
    create_users_kpi_sheet(service, spreadsheet_url, sheet_name)
    sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)

    # Получаем текущий месяц и год
    current_month = datetime.now().month
    current_year = datetime.now().year

    # Получаем количество дней в текущем месяце
    days_in_month = calendar.monthrange(current_year, current_month)[1]
    month_name = calendar.month_name[current_month]

    # Начальный индекс строки для таблицы
    row_index = find_last_filled_row_in_column(service, spreadsheet_id, sheet_name)

    # Заголовок таблицы
    headers = [month_name, "Total"] + list(range(1, days_in_month + 1))

    # Создаем запросы для добавления заголовков и формулы суммы
    requests = [
        {
            "updateCells": {
                "start": {"sheetId": sheet_id, "rowIndex": row_index, "columnIndex": 0},
                "rows": [{"values": [{"userEnteredValue": {"stringValue": str(header)}} for header in headers]}],
                "fields": "userEnteredValue"
            }
        },
        {
            "updateCells": {
                "start": {"sheetId": sheet_id, "rowIndex": row_index + 1, "columnIndex": 1},
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": f"=SUM(C{row_index + 2}:{get_column_letter(2 + days_in_month)}{row_index + 2})"}}]}],
                "fields": "userEnteredValue"
            }
        }
    ]

    # Добавляем запросы для установки границ и форматирования
    requests += [
        create_border_format_request(sheet_id, row_index, row_index + 2, 0, len(headers)),
        create_color_format_request(sheet_id, row_index, row_index + 1, 0, len(headers), {'red': 0.85, 'green': 0.93, 'blue': 0.98}),
        create_color_format_request(sheet_id, row_index + 1, row_index + 2, 0, len(headers), {'red': 0.85, 'green': 0.93, 'blue': 0.98})
    ]

    # Выполняем все запросы одной транзакцией
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()
    sleep(2)
    print(f'Таблица KPI для месяца "{month_name}" создана')
