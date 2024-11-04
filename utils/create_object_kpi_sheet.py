import calendar
from datetime import datetime
from time import sleep

from common.get_sheet_id import get_sheet_id
from kpi_utils.expenses_materials_in_objects import check_material_table


def get_materials_list(building_object):
    building_objects_dict = {}
    materials = building_object.material.all()
    if materials:  # Проверка наличия связанных объектов материалов
        material_names = [material.name for material in materials]
        # building_objects_dict[building_object.sh_url] = material_names
        sh_url = building_object.sh_url
        return sh_url, material_names



# Добавляю создание таблицы учета расходов материалов
def create_object_kpi_sheet(service, spreadsheet_url, build_object, sheet_name):
    """
    Функция ищет лист Object KPI
    Если находит его то возвращает его id
    Если такого листа нет, создает и возвращает его id
    """
    spreadsheet_id = spreadsheet_url.split("/")[5]
    # Получение информации о всех листах в таблице
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск идентификатора листа по имени
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            print(f"Sheet '{sheet_name}' already exists.")
            return sheet['properties']['sheetId']

    # Если лист не найден, создаем его
    print(f"Sheet '{sheet_name}' not found. Creating new sheet.")
    request_body = {
        'requests': [{
            'addSheet': {
                'properties': {
                    'title': sheet_name
                }
            }
        }]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request_body).execute()
    sheet_id = response.get('replies', [])[0].get('addSheet', {}).get('properties', {}).get('sheetId')

    # Создаем таблицу Object KPI
    total_budget = build_object.total_budget
    # Сoздаем таблицу учета материалов в Object KPI
    create_monthly_kpi_table(service, spreadsheet_url, total_budget, sheet_name)
    sh_url, materials = get_materials_list(build_object)
    check_material_table(sh_url, materials)
    return sheet_id


def create_monthly_kpi_table(service, spreadsheet_url, total_budget, sheet_name, start_row=0):
    """
    Создаёт таблицу KPI для текущего месяца с указанием выполненных работ за каждый день месяца,
    процента выполнения относительно цели и остатка бюджета.

    Args:
        service: Экземпляр сервиса Google Sheets API.
        spreadsheet_url: URL адрес Google таблицы.
        total_budget: Общий бюджет на месяц.
        sheet_name: Имя листа, где будет создаваться таблица.
        start_row: Начальная строка для таблицы (по умолчанию 0).

    Returns:
        None: Функция ничего не возвращает, но выводит сообщение о создании таблицы.
    """
    now = datetime.now()
    month_name = now.strftime("%B")
    days_in_month = calendar.monthrange(now.year, now.month)[1]
    spreadsheet_id = spreadsheet_url.split("/")[5]
    sheet_id = get_sheet_id(service, spreadsheet_url, sheet_name)

    # Создание списка запросов
    requests = []

    # Заголовки таблицы
    headers = [month_name, "Done today $", "Done today %", "Total left %", "Budget balance"]
    # Первая строка: Заголовки
    requests.append({
        "updateCells": {
            "start": {"sheetId": sheet_id, "rowIndex": start_row, "columnIndex": 0},
            "rows": [{"values": [{"userEnteredValue": {"stringValue": header}} for header in headers]}],
            "fields": "userEnteredValue"
        }
    })

    # Вторая строка: пустая, кроме ячейки "Budget balance"
    requests.append({
        "updateCells": {
            "start": {"sheetId": sheet_id, "rowIndex": start_row + 1, "columnIndex": 4},
            "rows": [{"values": [{"userEnteredValue": {"numberValue": total_budget}}]}],
            "fields": "userEnteredValue"
        }
    })

    # Следующие строки: данные по дням с формулами
    for day in range(1, days_in_month + 1):
        day_row = [
            {"userEnteredValue": {"numberValue": day}},  # День месяца
            {"userEnteredValue": {"stringValue": ""}},  # Пустая ячейка
            {"userEnteredValue": {"formulaValue": f"=ROUND((B{start_row + 2 + day} / E{start_row + 2 + day -1}) * 100; 2)"}},  # Done today %
            {"userEnteredValue": {"formulaValue": f"=ROUND((E{start_row + 2 + day} / E{start_row + 2}) * 100; 2)"}},  # Total done %
            {"userEnteredValue": {"formulaValue": f"=(E{start_row + 2 + day - 1} - B{start_row + 2 + day})"}}  # Budget balance
        ]
        requests.append({
            "updateCells": {
                "start": {"sheetId": sheet_id, "rowIndex": start_row + 1 + day, "columnIndex": 0},
                "rows": [{"values": day_row}],
                "fields": "userEnteredValue"
            }
        })

    # Установка границ таблицы
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": start_row + days_in_month + 2,
                "startColumnIndex": 0,
                "endColumnIndex": len(headers)
            },
            "top": {"style": "SOLID", "width": 1},
            "bottom": {"style": "SOLID", "width": 1},
            "left": {"style": "SOLID", "width": 1},
            "right": {"style": "SOLID", "width": 1},
            "innerHorizontal": {"style": "SOLID", "width": 1},
            "innerVertical": {"style": "SOLID", "width": 1}
        }
    })

    # Выполнение запросов
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()
    sleep(1)
    print(f'Таблица KPI для месяца "{month_name}" создана')

