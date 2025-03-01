import logging
from datetime import datetime
from time import sleep

from googleapiclient.errors import HttpError

from utils.create_users_kpi_sheet_and_table import create_users_kpi_table
from common.service import service


service = service.get_service()
logger = logging.getLogger(__name__)


def find_users_kpi_table(service, spreadsheet_url):
    # Извлечение идентификатора таблицы из URL
    spreadsheet_id = spreadsheet_url.split("/")[5]

    # Имя текущего месяца
    current_month = datetime.now().strftime("%B")

    # Диапазон ячеек для чтения (столбец A от 1-й до 1000-й)
    range_name = 'Users KPI!A:A'

    try:
        # Получение значений ячеек из указанного диапазона
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])

        # Проверка наличия текущего имени месяца в значениях ячеек
        for row in values:
            if row and current_month in row:
                print(f'Таблица для {current_month} есть.')
                return True
        print(f'Таблицы для {current_month} нет, она будет создана')
        return False

    except Exception as e:
        print(f"Error reading data from spreadsheet: {e}")
        return False


def update_kpi_values(data_list):
    """
    Функция для обновления данных KPI для users
    :param data_list:
    :return:
    """
    logger.info(f'Функция записи данных для [Users KPI] запущена')
    now = datetime.now()
    current_day = now.day
    current_month_name = now.strftime("%B")

    logger.info(f'Текущий месяц: {current_month_name}, Текущий день: {current_day}')

    for spreadsheet_url, user_data in data_list:
        if not find_users_kpi_table(service, spreadsheet_url):
            create_users_kpi_table(service, spreadsheet_url)
        spreadsheet_id = spreadsheet_url.split("/")[5]
        sheet_name = "Users KPI"

        range_name = f"{sheet_name}!A:A"
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        sleep(2)
        values = result.get('values', [])

        month_row_index = None
        for row_index, row in enumerate(values):
            if row and row[0] == current_month_name:
                month_row_index = row_index + 1
                break

        if month_row_index is not None:
            sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)

            for name, value in user_data:

                name_found = False
                empty_row_index = None

                user_rows = values[month_row_index:] + [[] for _ in range(100 - len(values[month_row_index:]))]

                # user_rows = values[month_row_index:month_row_index + 100]
                logger.info(f'Поиск имени {name} в строках {month_row_index}:{month_row_index + 100}')
                if not user_rows:
                    empty_row_index = month_row_index + 1
                else:
                    for name_row_index, name_row in enumerate(user_rows, start=month_row_index + 1):
                        if name_row and name_row[0] == name:
                            logger.info(f'Записываются данные {value} в таблицу [Users KPI] для пользователя {name}')
                            name_found = True
                            update_value_range = f"{sheet_name}!{get_column_letter(current_day)}{name_row_index}"
                            update_value_body = {'values': [[value]]}
                            service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=update_value_range, body=update_value_body, valueInputOption='USER_ENTERED').execute()
                            sleep(3)
                            break
                        if not name_row:
                            empty_row_index = name_row_index
                            break

                if not name_found and empty_row_index:
                    copy_paste_request = {
                        "copyPaste": {
                            "source": {"sheetId": sheet_id, "startRowIndex": empty_row_index - 1, "endRowIndex": empty_row_index},
                            "destination": {"sheetId": sheet_id, "startRowIndex": empty_row_index, "endRowIndex": empty_row_index + 1},
                            "pasteType": "PASTE_NORMAL"
                        }
                    }
                    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [copy_paste_request]}).execute()
                    sleep(2)

                    update_name_range = f"{sheet_name}!A{empty_row_index}"
                    # update_name_body = {'values': [[name]]}
                    # service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=update_name_range, body=update_name_body, valueInputOption='USER_ENTERED').execute()
                    # sleep(2)

                    update_value_range = f"{sheet_name}!{get_column_letter(current_day)}{empty_row_index}"
                    # update_value_body = {'values': [[value]]}
                    # service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=update_value_range, body=update_value_body, valueInputOption='USER_ENTERED').execute()
                    # sleep(2)
                    # Объединяем запросы на обновление имени и значения
                    update_requests = [
                        {
                            "range": update_name_range,
                            "values": [[name]]
                        },
                        {
                            "range": update_value_range,
                            "values": [[value]]
                        }
                    ]

                    # Выполняем два обновления в одном запросе
                    service.spreadsheets().values().batchUpdate(
                        spreadsheetId=spreadsheet_id,
                        body={
                            "valueInputOption": "USER_ENTERED",
                            "data": update_requests
                        }
                    ).execute()
                    sleep(2)

                    # Обновляем переменную values для отражения последних изменений в таблице
                    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
                    values = result.get('values', [])

###########################################################################
# def get_month_row_index(service, spreadsheet_id, sheet_name, current_month_name):
#     """Ищет индекс строки текущего месяца."""
#     range_name = f"{sheet_name}!A:A"
#     result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
#     values = result.get('values', [])
#     for row_index, row in enumerate(values):
#         if row and row[0] == current_month_name:
#             return row_index + 1
#     return None
#
#
# def update_user_kpi(service, spreadsheet_id, sheet_name, user_name, day, value, month_row_index):
#     """Обновляет значение KPI для указанного пользователя."""
#     range_name = f"{sheet_name}!A:A"
#     result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
#     values = result.get('values', [])
#     user_rows = values[month_row_index:] + [[] for _ in range(100 - len(values[month_row_index:]))]
#
#     for name_row_index, name_row in enumerate(user_rows, start=month_row_index + 1):
#         if name_row and name_row[0] == user_name:
#             update_range = f"{sheet_name}!{get_column_letter(day)}{name_row_index}"
#             body = {'values': [[value]]}
#             service.spreadsheets().values().update(
#                 spreadsheetId=spreadsheet_id, range=update_range, body=body, valueInputOption='USER_ENTERED'
#             ).execute()
#             return True  # Обновление выполнено
#     return False  # Пользователь не найден
#
#
# def add_user_to_sheet(service, spreadsheet_id, sheet_name, user_name, day, value, empty_row_index, sheet_id):
#     """Добавляет нового пользователя и записывает его KPI."""
#     copy_paste_request = {
#         "copyPaste": {
#             "source": {"sheetId": sheet_id, "startRowIndex": empty_row_index - 1, "endRowIndex": empty_row_index},
#             "destination": {"sheetId": sheet_id, "startRowIndex": empty_row_index, "endRowIndex": empty_row_index + 1},
#             "pasteType": "PASTE_NORMAL"
#         }
#     }
#     service.spreadsheets().batchUpdate(
#         spreadsheetId=spreadsheet_id, body={"requests": [copy_paste_request]}
#     ).execute()
#
#     update_requests = [
#         {"range": f"{sheet_name}!A{empty_row_index}", "values": [[user_name]]},
#         {"range": f"{sheet_name}!{get_column_letter(day)}{empty_row_index}", "values": [[value]]}
#     ]
#     service.spreadsheets().values().batchUpdate(
#         spreadsheetId=spreadsheet_id,
#         body={"valueInputOption": "USER_ENTERED", "data": update_requests}
#     ).execute()
#
#
# def update_kpi_values(data_list):
#     """Основная функция для обновления данных KPI."""
#     now = datetime.now()
#     current_day = now.day
#     current_month_name = now.strftime("%B")
#
#     for spreadsheet_url, user_data in data_list:
#         if not find_users_kpi_table(service, spreadsheet_url):
#             create_users_kpi_table(service, spreadsheet_url)
#
#         spreadsheet_id = spreadsheet_url.split("/")[5]
#         sheet_name = "Users KPI"
#         month_row_index = get_month_row_index(service, spreadsheet_id, sheet_name, current_month_name)
#
#         if month_row_index is None:
#             logging.error(f"Не найдена строка для месяца {current_month_name} в {spreadsheet_id}")
#             continue
#
#         sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)
#
#         for user_name, value in user_data:
#             if not update_user_kpi(service, spreadsheet_id, sheet_name, user_name, current_day, value, month_row_index):
#                 empty_row_index = month_row_index + len(user_data)
#                 add_user_to_sheet(service, spreadsheet_id, sheet_name, user_name, current_day, value, empty_row_index, sheet_id)
####################################################################


# def get_month_row_index(service, spreadsheet_id, sheet_name, current_month_name):
#     """Ищет индекс строки текущего месяца."""
#     try:
#         range_name = f"{sheet_name}!A:A"
#         result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
#         values = result.get('values', [])
#         for row_index, row in enumerate(values):
#             if row and row[0] == current_month_name:
#                 return row_index + 1
#         return None
#     except HttpError as error:
#         logging.error(f"Ошибка при получении данных: {error}")
#         if error.resp.status == 429:  # Превышен лимит запросов
#             sleep(2)
#         return None
#
#
# def update_user_kpi(service, spreadsheet_id, sheet_name, user_name, day, value, month_row_index):
#     """Обновляет значение KPI для указанного пользователя."""
#     try:
#         range_name = f"{sheet_name}!A:A"
#         result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
#         values = result.get('values', [])
#         user_rows = values[month_row_index:] + [[] for _ in range(100 - len(values[month_row_index:]))]
#
#         for name_row_index, name_row in enumerate(user_rows, start=month_row_index + 1):
#             if name_row and name_row[0] == user_name:
#                 update_range = f"{sheet_name}!{get_column_letter(day)}{name_row_index}"
#                 body = {'values': [[value]]}
#                 service.spreadsheets().values().update(
#                     spreadsheetId=spreadsheet_id, range=update_range, body=body, valueInputOption='USER_ENTERED'
#                 ).execute()
#                 return True  # Обновление выполнено
#         return False  # Пользователь не найден
#     except HttpError as error:
#         logging.error(f"Ошибка при обновлении KPI для {user_name}: {error}")
#         if error.resp.status == 429:
#             sleep(2)
#         return False
#
#
# def add_user_to_sheet(service, spreadsheet_id, sheet_name, user_name, day, value, empty_row_index, sheet_id):
#     """Добавляет нового пользователя и записывает его KPI."""
#     try:
#         # Копируем строку
#         copy_paste_request = {
#             "copyPaste": {
#                 "source": {"sheetId": sheet_id, "startRowIndex": empty_row_index - 1, "endRowIndex": empty_row_index},
#                 "destination": {"sheetId": sheet_id, "startRowIndex": empty_row_index, "endRowIndex": empty_row_index + 1},
#                 "pasteType": "PASTE_NORMAL"
#             }
#         }
#         service.spreadsheets().batchUpdate(
#             spreadsheetId=spreadsheet_id, body={"requests": [copy_paste_request]}
#         ).execute()
#
#         # Обновляем имя и значение KPI
#         update_requests = [
#             {"range": f"{sheet_name}!A{empty_row_index}", "values": [[user_name]]},
#             {"range": f"{sheet_name}!{get_column_letter(day)}{empty_row_index}", "values": [[value]]}
#         ]
#         service.spreadsheets().values().batchUpdate(
#             spreadsheetId=spreadsheet_id,
#             body={"valueInputOption": "USER_ENTERED", "data": update_requests}
#         ).execute()
#     except HttpError as error:
#         logging.error(f"Ошибка при добавлении пользователя {user_name}: {error}")
#         if error.resp.status == 429:
#             sleep(2)
#
#
# def update_kpi_values(data_list):
#     """Основная функция для обновления данных KPI."""
#     now = datetime.now()
#     current_day = now.day
#     current_month_name = now.strftime("%B")
#
#     for spreadsheet_url, user_data in data_list:
#         try:
#             if not find_users_kpi_table(service, spreadsheet_url):
#                 create_users_kpi_table(service, spreadsheet_url)
#
#             spreadsheet_id = spreadsheet_url.split("/")[5]
#             sheet_name = "Users KPI"
#             month_row_index = get_month_row_index(service, spreadsheet_id, sheet_name, current_month_name)
#
#             if month_row_index is None:
#                 logging.error(f"Не найдена строка для месяца {current_month_name} в {spreadsheet_id}")
#                 continue
#
#             sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)
#
#             for user_name, value in user_data:
#                 if not update_user_kpi(service, spreadsheet_id, sheet_name, user_name, current_day, value, month_row_index):
#                     empty_row_index = month_row_index + len(user_data)
#                     add_user_to_sheet(service, spreadsheet_id, sheet_name, user_name, current_day, value, empty_row_index, sheet_id)
#
#         except Exception as error:
#             logging.error(f"Общая ошибка при обработке таблицы {spreadsheet_url}: {error}")



def get_column_letter(column_index):
    """Преобразует индекс столбца в его буквенное обозначение ."""

    column_index += 1
    letters = []
    while column_index >= 0:
        column_index, remainder = divmod(column_index, 26)
        letters.append(chr(65 + remainder))
        column_index -= 1
    return ''.join(reversed(letters))


def get_sheet_id(service, spreadsheet_id, sheet_name):
    # Получение информации о всех листах в таблице
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск идентификатора листа по имени
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']

    raise ValueError(f"Sheet with name '{sheet_name}' not found.")