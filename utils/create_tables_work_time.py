import datetime
from time import sleep

from common.days_list import create_date_list
from common.get_service import get_service
from common.last_non_empty_row import get_last_non_empty_row

# ИСПОЛЬЗОВАТЬ для создания таблицы вручную
# from days_list import create_date_list
# from get_service import get_service

service = get_service()


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


# def create_work_time_tables(service, spreadsheet_url, sheet_name):
#     # # Извлечение spreadsheet_id из URL
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#
#     # Создание сервиса
#     # service = get_service()
#
#     sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
#     sheets = sheet_metadata.get('sheets', '')
#
#     # Поиск sheetId по имени листа
#     sheet_id = None
#     for sheet in sheets:
#         if sheet['properties']['title'] == sheet_name:
#             sheet_id = sheet['properties']['sheetId']
#             break
#
#     # Если лист не найден, добавляем новый лист с указанным именем
#     if sheet_id is None:
#         body = {
#             'requests': [{
#                 'addSheet': {
#                     'properties': {'title': sheet_name}
#                 }
#             }]
#         }
#         response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
#         # Получаем sheetId добавленного листа
#         sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
#
#     # Получение данных из функции create_date_list
#     date_chunks = create_date_list()
#
#     # Инициализация списка значений и запросов для форматирования
#     values = []
#     requests = []
#
#     current_row = 1  # Текущий номер строки для запросов форматирования
#     row_offset = 0  # Смещение строки для каждой новой таблицы
#
#     header_color = {'red': 0.85, 'green': 0.93, 'blue': 0.98}
#     row_color = {'red': 0.85, 'green': 0.93, 'blue': 0.98}
#
#     # Создание таблицы для каждого чанка
#     for chunk in date_chunks:
#
#         # Форматирование заголовков для каждой таблицы
#         headers = ["Per hour", "User Name"] + [f"{month} {day}" for month, _, day in chunk] + ["Total", "Salary"]
#         values.append(headers)
#         requests.append(create_color_format_request(sheet_id, current_row - 1, current_row, 0, len(headers), header_color))
#         requests.append(create_color_format_request(sheet_id, current_row, current_row + 1, 0, len(headers), row_color))
#
#         real_row = 2 + row_offset  # Реальный номер строки с учетом смещения
#
#         row = ['', ''] + [''] * len(chunk) + ['', '']
#         values.append(row)
#
#         # Добавление запроса для добавления границ
#         end_row = current_row + 1
#         requests.append({
#             'updateBorders': {
#                 'range': {
#                     'sheetId': sheet_id,
#                     'startRowIndex': current_row - 1,
#                     'endRowIndex': end_row,
#                     'startColumnIndex': 0,
#                     'endColumnIndex': len(headers)
#                 },
#                 'top': {'style': 'SOLID', 'width': 1},
#                 'bottom': {'style': 'SOLID', 'width': 1},
#                 'left': {'style': 'SOLID', 'width': 1},
#                 'right': {'style': 'SOLID', 'width': 1},
#                 'innerHorizontal': {'style': 'SOLID', 'width': 1},
#                 'innerVertical': {'style': 'SOLID', 'width': 1}
#             }
#         })
#
#         # Обновление текущего номера строки и добавление пустых строк
#         current_row += 2
#         for _ in range(3):
#             values.append([''] * len(headers))
#             current_row += 1
#         row_offset = current_row - 1  # Новое смещение для следующей таблицы
#
#     # Получение метаданных листов
#     sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
#     sheets = sheet_metadata.get('sheets', '')
#
#     # Добавление нового листа, если он не существует
#     if not any(sheet['properties']['title'] == sheet_name for sheet in sheets):
#         body = {
#             'requests': [{
#                 'addSheet': {
#                     'properties': {'title': sheet_name}
#                 }
#             }]
#         }
#         service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
#
#     # Заполнение листа данными
#     body = {
#         'values': values
#     }
#     range_name = f"{sheet_name}!A1"
#     service.spreadsheets().values().update(
#         spreadsheetId=spreadsheet_id, range=range_name, body=body, valueInputOption='USER_ENTERED'
#     ).execute()
#
#     # Применение запросов форматирования
#     if requests:
#         service.spreadsheets().batchUpdate(
#             spreadsheetId=spreadsheet_id, body={'requests': requests}
#         ).execute()
#
#     return f'Table created in sheet "{sheet_name}" in the spreadsheet: {spreadsheet_url}'


def create_work_time_tables(spreadsheet_url, sheet_name='Work Time'):
    # Извлечение spreadsheet_id из URL
    spreadsheet_id = spreadsheet_url.split("/")[5]

    # Получение данных о всех листах
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск sheetId по имени листа
    sheet_id = None
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            sheet_id = sheet['properties']['sheetId']
            break

    # Если лист не найден, добавляем новый лист
    if sheet_id is None:
        body = {
            'requests': [{
                'addSheet': {
                    'properties': {'title': sheet_name}
                }
            }]
        }
        response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']

    # Получение всех дат в виде чанков
    date_chunks = create_date_list()

    # Определение текущей даты
    today = datetime.datetime.now()
    current_month = today.strftime("%B")  # Например, "May"
    current_day = today.strftime("%d").lstrip("0")  # Удаление ведущих нулей, например, "8" вместо "08"

    # Поиск чанка, содержащего текущую дату
    current_chunk = next((chunk for chunk in date_chunks if
                          any(day == current_day and month == current_month for month, _, day in chunk)), None)

    if not current_chunk:
        return "No chunk found for today's date."

    # Инициализация списка значений и запросов для форматирования
    values = []
    requests = []
    
    # TODO Добавить проверку какая стока сейчас пустая
    # current_row = 1
    #############
    current_row = get_last_non_empty_row(service, spreadsheet_url, sheet_name='Work Time')
    if current_row == 0:
        current_row += 1
    else:
        current_row += 5
    print(f'Current row = {current_row}')
    sleep(5)
    # #############

    # Формирование таблицы для текущего чанка
    headers = ["Per hour", "User Name"] + [f"{month} {day}" for month, _, day in current_chunk] + ["Total", "Salary"]
    values.append(headers)

    # Форматирование заголовков и строк
    header_color = {'red': 0.85, 'green': 0.93, 'blue': 0.98}
    row_color = {'red': 0.85, 'green': 0.93, 'blue': 0.98}

    requests.append(create_color_format_request(sheet_id, current_row - 1, current_row, 0, len(headers), header_color))
    requests.append(create_color_format_request(sheet_id, current_row, current_row + 1, 0, len(headers), row_color))

    # Добавление пустой строки данных
    row = ['', ''] + [''] * len(current_chunk) + ['', '']
    values.append(row)

    # Добавление запроса для отрисовки границ
    end_row = current_row + 1
    requests.append({
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': current_row - 1,
                'endRowIndex': end_row,
                'startColumnIndex': 0,
                'endColumnIndex': len(headers)
            },
            'top': {'style': 'SOLID', 'width': 1},
            'bottom': {'style': 'SOLID', 'width': 1},
            'left': {'style': 'SOLID', 'width': 1},
            'right': {'style': 'SOLID', 'width': 1},
            'innerHorizontal': {'style': 'SOLID', 'width': 1},
            'innerVertical': {'style': 'SOLID', 'width': 1}
        }
    })

    # Заполнение листа данными
    body = {
        'values': values
    }
    ##############
    # range_name = f"{sheet_name}!A1"
    range_name = f"{sheet_name}!A{current_row}"
    #############
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=range_name, body=body, valueInputOption='USER_ENTERED'
    ).execute()

    # Применение запросов форматирования
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={'requests': requests}
        ).execute()

    return f'Table created in sheet "{sheet_name}" in the spreadsheet: {spreadsheet_url}'


# Пример вызова функции
# create_work_time_tables(
#     spreadsheet_url='https://docs.google.com/spreadsheets/d/1JFbB5TucFZn-xl4NtTVpcYAvqP89DZmOy5xxB18rBsI/edit#gid=661746781',
#     sheet_name='Work Time')
