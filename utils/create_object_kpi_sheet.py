import calendar
from datetime import datetime
from common.get_sheet_id import get_sheet_id


# def convert_to_float(value):
#     """Функция для преобразования строки в число с плавающей точкой"""
#     if value:
#         # Замена запятых на точки и удаление пробелов
#         value = value.replace(',', '.').strip()
#         try:
#             # Попытка преобразовать строку в число с плавающей точкой
#             return float(value)
#         except ValueError:
#             # В случае неудачи возвращаем 0
#             return 0
#     else:
#         return 0


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
    # print(f"New sheet '{sheet_name}' created with Sheet ID: {sheet_id}")
    # Создаем таблицу Object KPI
    # create_monthly_kpi_table(service, sheet_id, spreadsheet_id, build_object)
    total_budget = build_object.total_budget
    create_monthly_kpi_table(service, spreadsheet_url, total_budget, sheet_name)
    return sheet_id


# def create_monthly_kpi_table(service, spreadsheet_url, total_budget, sheet_name, start_row=0):
#     """
#     Функция создания таблицы Objext KPI для текущего месяца
#     """
#
#     # Определение текущего месяца и количества дней в нем
#     now = datetime.now()
#     month_name = now.strftime("%B")
#     days_in_month = calendar.monthrange(now.year, now.month)[1]
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#     sheet_id = get_sheet_id(service, spreadsheet_url, sheet_name)
#
#     # Начальный индекс строки и столбца
#     # start_row_index = 0
#     start_column = 0
#
#     # Формирование заголовков
#     headers = [month_name, "Done today $", "Done today %", "Total done %", "Budget balance"]
#     days_headers = list(range(1, days_in_month + 1))
#
#     # Создаем запросы для добавления заголовков и дней месяца
#     requests = [{
#         "updateCells": {
#             "start": {"sheetId": sheet_id, "rowIndex": start_row, "columnIndex": start_column},
#             "rows": [{"values": [{"userEnteredValue": {"stringValue": str(header)}} for header in headers]}],
#             "fields": "userEnteredValue"
#         }
#     }]
#
#     # Добавляем дни месяца
#     for i, day in enumerate(days_headers, start=start_row + 1):
#         requests.append({
#             "updateCells": {
#                 "start": {"sheetId": sheet_id, "rowIndex": i, "columnIndex": start_column},
#                 "rows": [{"values": [{"userEnteredValue": {"numberValue": day}}]}],
#                 "fields": "userEnteredValue"
#             }
#         })
#
#     # Значение Budget balance во второй строке
#     requests.append({
#         "updateCells": {
#             "start": {"sheetId": sheet_id, "rowIndex": start_row + 1, "columnIndex": len(headers) - 1},
#             "rows": [{"values": [{"userEnteredValue": {"stringValue": str(total_budget)}}]}],
#             "fields": "userEnteredValue"
#         }
#     })
#
#     # Установка границ для всей таблицы
#     requests.append({
#         "updateBorders": {
#             "range": {
#                 "sheetId": sheet_id,
#                 "startRowIndex": start_row,
#                 "endRowIndex": start_row + days_in_month + 1,
#                 "startColumnIndex": start_column,
#                 "endColumnIndex": len(headers)
#             },
#             "top": {"style": "SOLID", "width": 1},
#             "bottom": {"style": "SOLID", "width": 1},
#             "left": {"style": "SOLID", "width": 1},
#             "right": {"style": "SOLID", "width": 1},
#             "innerHorizontal": {"style": "SOLID", "width": 1},
#             "innerVertical": {"style": "SOLID", "width": 1}
#         }
#     })
#
#     # Выполняем все запросы одной транзакцией
#     service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()
#     print(f'Таблица KPI для месяца "{month_name}" создана')


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

    print(f'TOTAL_BUDGET -----------------> {type(total_budget)} <--------------------')

    # Заголовки таблицы
    headers = [month_name, "Done today $", "Done today %", "Total done %", "Budget balance"]
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
            {"userEnteredValue": {"formulaValue": f"=(B{start_row + 2 + day} / E{start_row + 2 + day -1}) * 100"}},  # Done today %
            {"userEnteredValue": {"formulaValue": f"=(E{start_row + 2 + day} / E{start_row + 2}) * 100"}},  # Total done %
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
    print(f'Таблица KPI для месяца "{month_name}" создана')

