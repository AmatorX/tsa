from datetime import datetime
from time import sleep

from common.service import service
from kpi_utils.expenses_materials_in_objects import create_materials_table
# from kpi_utils.expenses_materials_in_objects import create_materials_table
from utils.create_object_kpi_sheet import create_monthly_kpi_table
from common.last_non_empty_row import get_last_non_empty_row


service = service.get_service()


def is_first_of_the_month():
    """
    Функция проверки числа месяца.
    Вернет True если 1-е число, иначе вернет False
    """
    today = datetime.now()
    return today.day == 1


def get_remaining_budget(service, spreadsheet_url, sheet_name):
    spreadsheet_id = spreadsheet_url.split("/")[5]
    range_name = f"{sheet_name}!E1:E"  # Запрашиваем данные только из первого столбца
    # Получаем данные столбца
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    # Определяем номер последней заполненной строки
    last_row = len(values) - 1 if values else 0
    if last_row >= 0:
        budget = values[last_row][0]  # Получаем значение из последней непустой строки
    else:
        budget = None  # Если в столбце нет значений, возвращаем None
    return budget


def get_sheet_id(service, spreadsheet_id, sheet_name):
    # Получение идентификатора листа по его имени
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    return None


def check_materials_in_table(spreadsheet_id, materials, month_row):
    """
    Функция проверяет все ли материалы из materials существуют в таблицах
    """
    start_column = 'I'
    end_column = ''  # Последний столбец который нужен, если оставить пустым, берутся все до конца
    row_range = f"Object KPI!{start_column}{month_row}:{end_column}{month_row}"
    row_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=row_range).execute()

    row_values = row_result.get('values', [])
    materials_list_in_table = [item if item else "" for sublist in row_values for item in sublist]
    # print(f'flat row values: {materials_list_in_table}')
    # print(f'materials: {materials}')
    # TODO Реализовать добавление материалов которых  нет в таблице, но они переданы в списке
    return materials_list_in_table


def check_material_table(spreadsheet_url, materials):
    """
    Функция проверяет есть ли таблица для учета расхода материалов по объекту.
    Если таблицы нет, вызывается фукнция создания таблицы create_materials_table.
    :return: Индекс строки текущего месяца
    """
    spreadsheet_id = spreadsheet_url.split("/")[5]
    range_name = "Object KPI!H1:I"
    now = datetime.now()
    current_month = now.strftime("%B")
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    flat_values = [item[0] if item else "" for item in values]

    last_entered_row = len(values) if values else 0

    if last_entered_row != 0:
        start_row = last_entered_row + 4
    else:
        start_row = last_entered_row + 1

    if current_month not in flat_values:
        # Создаем таблицу расход материалов
        create_materials_table(spreadsheet_url, materials, start_row=start_row)
        month_row = start_row
        materials_list_in_table = check_materials_in_table(spreadsheet_id, materials, month_row)
    else:
        month_row = flat_values.index(current_month) + 1
        materials_list_in_table = check_materials_in_table(spreadsheet_id, materials, month_row)

    return materials_list_in_table, month_row


def create_materials_table(spreadsheet_url, materials, start_row):
    """
    Функция создания таблицы для учета материалов в листе Object KPI
    """
    # Извлечение spreadsheet_id из URL
    spreadsheet_id = spreadsheet_url.split("/")[5]

    # Получение текущего месяца
    now = datetime.now()
    current_month = now.strftime("%B")

    # Определение диапазона для вставки данных
    sheet_name = "Object KPI"
    start_column = "H"
    end_column = chr(ord(start_column) + len(materials))

    # Подготовка данных для вставки
    values = []

    # Первая строка: текущий месяц и имена материалов
    header_row = [current_month] + materials
    values.append(header_row)

    # Вторая строка: "Units" и пустые ячейки
    units_row = ["Units"] + [""] * len(materials)
    values.append(units_row)

    # Остальные строки: числа текущего месяца и пустые ячейки
    days_in_month = (datetime(now.year, now.month + 1, 1) - datetime(now.year, now.month, 1)).days
    for day in range(1, days_in_month + 1):
        row = [day] + [""] * len(materials)
        values.append(row)

    # Запрос на обновление значений
    body = {
        'values': values
    }
    range_name = f"{sheet_name}!{start_column}{start_row}:{end_column}{start_row + len(values) - 1}"

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        body=body,
        valueInputOption='USER_ENTERED'
    ).execute()

    # Запросы на форматирование ячеек (границы)
    requests = []

    # Границы для всей таблицы
    requests.append({
        'updateBorders': {
            'range': {
                'sheetId': get_sheet_id(service, spreadsheet_id, sheet_name),
                'startRowIndex': start_row - 1,
                'endRowIndex': start_row - 1 + len(values),
                'startColumnIndex': ord(start_column) - ord('A'),
                'endColumnIndex': ord(end_column) - ord('A') + 1
            },
            'top': {'style': 'SOLID'},
            'bottom': {'style': 'SOLID'},
            'left': {'style': 'SOLID'},
            'right': {'style': 'SOLID'},
            'innerHorizontal': {'style': 'SOLID'},
            'innerVertical': {'style': 'SOLID'}
        }
    })

    # Применение запросов на форматирование
    body = {'requests': requests}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def if_object_kpi_tables_exist(data, sheet_name='Object KPI'):
    print(f'DATA: {data}')
    for spreadsheet_url, materials in data.items():
    # Если сегодня 1-е число месяца
        if is_first_of_the_month():
            last_row = get_last_non_empty_row(service, spreadsheet_url, sheet_name)
            start_row = last_row + 3
            total_budget = get_remaining_budget(service, spreadsheet_url, sheet_name)

            # Создание таблицы KPI для текущего месяца
            create_monthly_kpi_table(service, spreadsheet_url, total_budget=total_budget, sheet_name=sheet_name,
                                     start_row=start_row)
            print(f"Updated KPI table for sheet {sheet_name} at {spreadsheet_url}")
            materials_list_in_table, month_row_index = check_material_table(spreadsheet_url, materials)
        else:
            print('Сщздание таблиц не требуется')

