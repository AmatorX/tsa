from datetime import datetime
import calendar
from time import sleep

from common.return_results_curr_month_year import generate_results_filename
from common.service import service


service = service.get_service()


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


def get_sheet_id(service, spreadsheet_id, sheet_name):
    # Получение идентификатора листа по его имени
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    return None


def merge_dictionaries(dict1, dict2):
    result_dict = {}

    for key, material in dict1.items():
        if material in dict2:
            result_dict[key] = dict2[material]
        else:
            result_dict[key] = 0  # или любое другое значение по умолчанию, если материал не найден в dict2
    print(f'Result Dict {result_dict}')
    return result_dict


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


def create_materials_dict(materials):
    """
    Функция принимает список материалов и возвращает словарь, где ключами являются
    последовательные буквы начиная с 'I', а значениями - элементы списка.
    """
    start_char_code = ord('I')
    materials_dict = {}

    for index, material in enumerate(materials):
        column_letter = chr(start_char_code + index)
        materials_dict[column_letter] = material

    return materials_dict


def get_row_indexes_total_materials(spreadsheet_url, days_in_month):
    """
    Функция вернет список индексов строк Total materials из таблицы results_current_month_current_year
    """
    spreadsheet_id = spreadsheet_url.split("/")[5]
    sh_name = generate_results_filename()
    range_name = f"{sh_name}!A1:A"
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    flat_values = [item[0] if item else "" for item in values]
    total_users = flat_values.count("Name") - 1
    # print(f"Flat values: {flat_values}")
    # print(f"Total users: {total_users}")

    indexes_total_materials = []
    first_element = 2 * days_in_month + 18
    indexes_total_materials.append(first_element)
    for i in range(total_users - 1):
        next_element = indexes_total_materials[-1] + days_in_month + 12
        indexes_total_materials.append(next_element)
    return indexes_total_materials


def update_materials_values(spreadsheet_url, materials, materials_list_in_table, month_row_index):
    """
    Функция обновления данныъ сколько выполнено по количеству материалов.
    Обновляет данные в таблице в Object KPI
    """
    day = datetime.now().day
    day_row = month_row_index + day + 1
    material_letters = create_materials_dict(materials_list_in_table)

    current_date = datetime.today()
    # Получаем количество дней в текущем месяце
    days_in_month = calendar.monthrange(current_date.year, current_date.month)[1]

    row_indexes_total_materials = get_row_indexes_total_materials(spreadsheet_url, days_in_month)

    # TODO Необходимо спарсить для каждого материала количесвто из строк Total materials каждого пользователя
    # таблиц результатов в месяце, получить общую сумму по каждому материалу,
    # Затем создаем словарь вида {'material_1': 1234, 'material_2': 1234, 'material_3': 1234}
    # Затем по идексам колонок из material_letters пишем данные каждого материала в
    # соответвующую месяцу таблицу в Object KPI

    print(f'Spreadsheet url: {spreadsheet_url}')
    print(f'Day row in Object KPI table: {day_row}')
    print(f'Materials: {materials}')
    print(f'Materials in table: {materials_list_in_table}')
    print(f'Materials Letters in table: {material_letters}')
    print(f'Indexes of the material rows in the results table for the current month: {row_indexes_total_materials}')
    material_total_used = calculate_sums_materials_in_results(spreadsheet_url, row_indexes_total_materials, sh_name=generate_results_filename())
    data_dict = merge_dictionaries(material_letters, material_total_used)
    update_spreadsheet(spreadsheet_url, data_dict)



# def update_spreadsheet(spreadsheet_url, data_dict):
#     """
#     Функция обновляет данные за текущий день по использованию материалов
#     """
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#     sheet_name = "Object KPI"
#
#     # Определение текущего месяца и дня
#     now = datetime.now()
#     current_month = now.strftime("%B")
#     current_day = now.day
#
#     # Получение всех значений в столбце H
#     range_name = f"{sheet_name}!H:H"
#     result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
#     values = result.get('values', [])
#
#     # Поиск строки с текущим месяцем
#     month_row = None
#     for idx, row in enumerate(values):
#         if row and row[0] == current_month:
#             month_row = idx
#             break
#
#     if month_row is None:
#         raise ValueError("Текущий месяц не найден в столбце H.")
#
#     # Определение целевой строки
#     target_row = month_row + current_day + 1
#
#     # Подготовка данных для обновления
#     updates = []
#     for col_letter, value in data_dict.items():
#         # Получение значения из ячейки month_row + 1 в соответствующем столбце
#         source_range = f"{sheet_name}!{col_letter}{month_row + 2}"
#         source_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=source_range).execute()
#         source_value = source_result.get('values', [[]])[0][0]
#         source_value = float(source_value) if source_value else 0.0
#
#         # Вычисление нового значения и добавление его в список обновлений
#         new_value = source_value - value
#         target_range = f"{sheet_name}!{col_letter}{target_row + 1}"
#         updates.append({
#             'range': target_range,
#             'values': [[new_value]]
#         })
#
#     # Выполнение обновления
#     body = {
#         'valueInputOption': 'USER_ENTERED',
#         'data': updates
#     }
#     service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()



def update_spreadsheet(spreadsheet_url, data_dict):
    spreadsheet_id = spreadsheet_url.split("/")[5]
    sheet_name = "Object KPI"

    # Определение текущего месяца и дня
    now = datetime.now()
    current_month = now.strftime("%B")
    current_day = now.day

    # Получение всех значений в столбце H
    range_name = f"{sheet_name}!H:H"
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    # Поиск строки с текущим месяцем
    month_row = None
    for idx, row in enumerate(values):
        if row and row[0] == current_month:
            month_row = idx
            break

    if month_row is None:
        raise ValueError("Текущий месяц не найден в столбце H.")

    # Определение целевой строки
    target_row = month_row + current_day + 1

    # Подготовка данных для обновления
    updates = []
    for col_letter, value in data_dict.items():
        # Получение значения из ячейки month_row + 1 в соответствующем столбце
        source_range = f"{sheet_name}!{col_letter}{month_row + 2}"
        source_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=source_range).execute()
        source_values = source_result.get('values', [])
        sleep(2)

        # Проверка наличия значения в ячейке
        if source_values and source_values[0]:
            source_value = float(source_values[0][0])
        else:
            source_value = 0.0

        # Вычисление нового значения и добавление его в список обновлений
        new_value = source_value - value
        target_range = f"{sheet_name}!{col_letter}{target_row + 1}"
        updates.append({
            'range': target_range,
            'values': [[new_value]]
        })

    # Выполнение обновления
    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': updates
    }
    service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def calculate_sums_materials_in_results(spreadsheet_url, row_indexes, sh_name):

    spreadsheet_id = spreadsheet_url.split("/")[5]

    # Запрашиваем диапазон данных из колонок B, C и далее
    range_name = f"{sh_name}!B:Z"
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    # Получаем первую строку для названий
    header_row = values[3] if len(values) > 3 else []

    # Инициализация словаря для хранения результатов
    results = {}

    # Проход по каждому столбцу, начиная с колонки B (индекс 1)
    for col_index in range(len(header_row)):
        # Суммируем значения в целевых ячейках для текущего столбца
        sum_x = sum(float(values[row_index - 1][col_index]) for row_index in row_indexes if row_index - 1 < len(values) and col_index < len(values[row_index - 1]))

        # Получаем имя из ячейки в строке 4 для текущего столбца
        name = header_row[col_index]

        # Создаем пару ключ: значение в словаре
        results[name] = sum_x
    print(f'Results {results}')
    sleep(4)
    return results


def update_materials_on_objects(data: dict):
    """
    Функция обновления данных в таблице учета материалов объекта.
    Обновляет данные на текущий день
    """
    for spreadsheet_url, materials in data.items():
        materials_list_in_table, month_row_index = check_material_table(spreadsheet_url, materials)
        update_materials_values(spreadsheet_url, materials, materials_list_in_table,  month_row_index)
        sleep(4)




# spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1gKsdz-icvXkx7km6Dl5RIqcEkGhn9GmFp80BIt3kVco/edit#gid=1862401406'
# materials = ['Test1', 'Test2', 'Test3', 'Test4']
#
# check_material_table(spreadsheet_url, materials)