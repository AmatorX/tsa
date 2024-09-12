import datetime
from calendar import monthrange
import logging
import string
from time import sleep
from logging_config import setup_logging
from common.return_results_curr_month_year import generate_results_filename
from common.service import service


service = service.get_service()
setup_logging()


def make_lst(data: list):
    return [material[0] for material in data]


def get_index_pairs(spreadsheet_id, sheet_name):
    logging.debug(f'Sheet name: {sheet_name}')
    range_name = f"{sheet_name}!A1:A"
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    indices = {'Price': [], 'Material': []}
    for i, val in enumerate(values):
        if val in [['Price'], ['Material']]:
            indices[val[0]].append(i + 1)

    return list(zip(indices['Price'], indices['Material']))


def get_value_pairs(spreadsheet_id, sheet_name, index_pairs):
    """
    Функция get_value_pairs извлекает значения строк для Price и Material из Google Sheets, а затем создает словарь,
    где ключами являются буквы столбцов, а значениями - списки, содержащие material и price.
    """
    if not index_pairs:
        return {}

    price_idx, material_idx = index_pairs[0]
    price_range = f"{sheet_name}!B{price_idx}:Z{price_idx}"
    material_range = f"{sheet_name}!B{material_idx}:Z{material_idx}"

    price_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=price_range).execute()
    material_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=material_range).execute()

    price_values = price_result.get('values', [[]])[0]
    material_values = material_result.get('values', [[]])[0]

    columns = [chr(i) for i in range(ord('B'), ord('Z') + 1)]
    return {col: [material, price] for col, material, price in zip(columns, material_values, price_values)}


def add_missing_materials(spreadsheet_id, sheet_name, index_pairs, value_pairs, materials_info):
    columns = [chr(i) for i in range(ord('B'), ord('Z') + 1)]
    used_columns = set(value_pairs.keys())

    for material, price in materials_info:
        if any(material in values for values in value_pairs.values()):
            continue  # Skip materials already in the value_pairs

        # Find the first available column that isn't used
        for col in columns:
            if col not in used_columns:
                used_columns.add(col)
                for price_idx, material_idx in index_pairs:
                    material_cell = f"{sheet_name}!{col}{material_idx}"
                    price_cell = f"{sheet_name}!{col}{price_idx}"

                    material_value_result = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=material_cell
                    ).execute()
                    # sleep(1)
                    price_value_result = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=price_cell
                    ).execute()
                    # sleep(1)

                    material_value = material_value_result.get('values', [[]])[0]
                    price_value = price_value_result.get('values', [[]])[0]

                    if not (material_value and material_value[0].startswith("=")):
                        # print(f"Updating material cell {material_cell} with value {material}")
                        service.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range=material_cell,
                            valueInputOption="RAW",
                            body={"values": [[material]]}
                        ).execute()
                        # sleep(1)

                    if not (price_value and price_value[0].startswith("=")):
                        # print(f"Updating price cell {price_cell} with value {price}")
                        service.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range=price_cell,
                            valueInputOption="RAW",
                            body={"values": [[price]]}
                        ).execute()
                        # sleep(0.1)
                break


# ####################################################################################################
# Здесь функции добавления материалов в таблицы учета материалов Object KPI
def update_materials_in_object_kpi(spreadsheet_id, materials_list):
    row_index = get_row_for_current_month(spreadsheet_id, sheet_name='Object KPI', column='H')
    in_stock = get_materials_dict(spreadsheet_id, row_index, sheet_name='Object KPI')
    # sleep(60)
    values_to_record = find_missing_materials(spreadsheet_id, row_index, in_stock, materials_list, sheet_name='Object KPI')
    # sleep(60)
    write_materials_to_sheet(spreadsheet_id, row_index, values_to_record, sheet_name='Object KPI')
    # sleep(60)


def get_row_for_current_month(spreadsheet_id, sheet_name='Object KPI', column='H'):
    """
    Парсит столбец H на листе Object KPI, находит название текущего месяца и возвращает индекс строки текущего месяца + 1.
    """

    # Получение текущего месяца в формате полного названия месяца (например, "July")
    current_month = datetime.datetime.now().strftime("%B")

    # Определение диапазона для чтения данных из столбца H
    range_name = f"{sheet_name}!{column}:{column}"

    # Получение данных из столбца H
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    # Поиск строки, содержащей название текущего месяца
    for i, row in enumerate(values):
        if current_month in row:
            logging.debug(f'Индекс строки текущего месяца {i + 1}')
            return i + 1  # Возвращаем индекс строки + 1

    return None  # Если текущий месяц не найден


def get_materials_dict(spreadsheet_id, row_index, sheet_name='Object KPI'):
    """
    Парсит все ячейки по индексу строки, начиная с ячейки в столбце I, и возвращает словарь in_stock.
    Ключ - название материала, значение - буква столбца.
    """

    # Определение диапазона для чтения данных из строки начиная с колонки I
    start_column = 'I'
    end_column = 'Z'  # Установите максимальный конечный столбец для чтения
    range_name = f"{sheet_name}!{start_column}{row_index}:{end_column}{row_index}"

    # Получение данных из строки
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name, majorDimension='ROWS').execute()
    row_values = result.get('values', [])[0]  # Получаем значения строки, если пусто, возвращаем пустой список
    logging.debug(f'ROW VALUES: {row_values}')
    in_stock = {}

    # Обработка значений строки
    for idx, value in enumerate(row_values):
        if value:  # Проверяем, что значение не пустое
            column_letter = string.ascii_uppercase[idx + 8]  # +8 потому что I - это 9-я колонка (0-индексированный список)
            in_stock[value] = column_letter
    logging.debug(f'Список материалов в таблице: {in_stock}')
    return in_stock


def get_next_column_letter(last_column_letter):
    """
    Возвращает следующую букву столбца после заданной.
    """
    last_index = string.ascii_uppercase.index(last_column_letter)
    if last_index == len(string.ascii_uppercase) - 1:
        # Если последний столбец Z, следующий будет AA
        return 'AA'
    else:
        return string.ascii_uppercase[last_index + 1]


def find_missing_materials(spreadsheet_id, row_index, in_stock, materials_list, sheet_name='Object KPI'):
    """
    Находит материалы, которых нет в in_stock, но которые есть в materials_list.
    Создает словарь values_to_record, где ключ - название материала, значение - буква колонки.
    """

    # Получение последней буквы столбца из in_stock
    if in_stock:
        logging.debug(f'IN SOCK {in_stock}')
        last_column_letter = max(in_stock.values(), key=lambda x: string.ascii_uppercase.index(x))
    else:
        last_column_letter = 'H'  # Начнем с колонки I, если in_stock пуст
    logging.debug(f'MATERIALS LIST {materials_list}')
    values_to_record = {}
    for material in materials_list:
        if material not in in_stock:
            next_column_letter = get_next_column_letter(last_column_letter)
            values_to_record[material] = next_column_letter
            last_column_letter = next_column_letter
    logging.debug(f'VALUES TO RECORD: {values_to_record}')
    return values_to_record


def get_sheet_id(service, spreadsheet_id, sheet_name):
    """
    Возвращает идентификатор листа по его имени.
    """
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    for sheet in sheets:
        if sheet.get("properties", {}).get("title", "") == sheet_name:
            return sheet.get("properties", {}).get("sheetId", "")
    return None


def write_materials_to_sheet(spreadsheet_id, row_index, values_to_record, sheet_name='Object KPI'):
    """
    Записывает данные в столбцы, указанные в values_to_record, начиная с row_index.
    Для каждого столбца записывает название материала и очерчивает границы.
    """
    sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)

    if sheet_id is None:
        logging.error(f"Sheet with name '{sheet_name}' not found.")
        raise ValueError(f"Sheet with name '{sheet_name}' not found.")

    # Определение текущего месяца и количества дней в нем
    today = datetime.datetime.now()
    current_year = today.year
    current_month = today.month
    _, num_days = monthrange(current_year, current_month)

    requests = []

    for material, column_letter in values_to_record.items():
        start_column_index = ord(column_letter) - ord('A')
        end_column_index = start_column_index + 1
        start_row_index = row_index - 1
        end_row_index = start_row_index + num_days + 2  # Исправление индекса

        # Очерчивание всех ячеек в столбце
        requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row_index,
                    "endRowIndex": end_row_index,
                    "startColumnIndex": start_column_index,
                    "endColumnIndex": end_column_index
                },
                "top": {
                    "style": "SOLID",
                    "width": 1,
                    "color": {"red": 0, "green": 0, "blue": 0}
                },
                "bottom": {
                    "style": "SOLID",
                    "width": 1,
                    "color": {"red": 0, "green": 0, "blue": 0}
                },
                "left": {
                    "style": "SOLID",
                    "width": 1,
                    "color": {"red": 0, "green": 0, "blue": 0}
                },
                "right": {
                    "style": "SOLID",
                    "width": 1,
                    "color": {"red": 0, "green": 0, "blue": 0}
                },
                "innerHorizontal": {
                    "style": "SOLID",
                    "width": 1,
                    "color": {"red": 0, "green": 0, "blue": 0}
                },
                "innerVertical": {
                    "style": "SOLID",
                    "width": 1,
                    "color": {"red": 0, "green": 0, "blue": 0}
                }
            }
        })

        # Запись названия материала в первую ячейку столбца
        requests.append({
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row_index,
                    "endRowIndex": start_row_index + 1,
                    "startColumnIndex": start_column_index,
                    "endColumnIndex": end_column_index
                },
                "rows": [{
                    "values": [{
                        "userEnteredValue": {"stringValue": material}
                    }]
                }],
                "fields": "userEnteredValue"
            }
        })

    body = {
        "requests": requests
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()
    logging.info(f"Materials {list(values_to_record.keys())} have been added to the sheet '{sheet_name}'"
                 f" starting at row {row_index}.")
    return (f"Materials {list(values_to_record.keys())} have been added to the sheet '{sheet_name}'"
            f" starting at row {row_index}.")


def update_materials_in_sheets(sh_url, materials_info):
    logging.debug(f'URL листа {sh_url} \n материалы {materials_info}')
    spreadsheet_id = sh_url.split("/")[5]
    sheet_name = generate_results_filename()

    index_pairs = get_index_pairs(spreadsheet_id, sheet_name)
    logging.debug(f'Index pairs: {index_pairs}\n')
    value_pairs = get_value_pairs(spreadsheet_id, sheet_name, index_pairs)
    logging.debug(f'Value pairs: {value_pairs}\n')
    add_missing_materials(spreadsheet_id, sheet_name, index_pairs, value_pairs, materials_info)
    # Обновление данных расхода материалов в таблицах листов Object KPI

    # Создаем список с материалами, которые сейчамс привязаны к объекту
    materials_list = make_lst(materials_info)
    # Функция обновления материалов в таблицах учета расхода материалов

    logging.debug(f'Список материалов перед добавление в таблицу учета в Object KPI {materials_list}')
    update_materials_in_object_kpi(spreadsheet_id, materials_list)

