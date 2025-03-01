import datetime
import logging
from time import sleep

from common.return_results_curr_month_year import generate_results_filename
from .get_start_end_indexes_rows import get_start_end_table
from common.service import service


logger = logging.getLogger(__name__)
service = service.get_service()


def get_names_in_results(spreadsheet_id):
    """
    Возвращает значения ячеек
    """
    # Определение диапазона для чтения данных из столбца B
    sheet_name = generate_results_filename()
    range_name = f"{sheet_name}!B:B"

    # Получение данных из столбца B
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    print(f'Список ячеек для поиска пустой строки\n'
          f'--------------->'
          f' {values}')
    return values


def convert_to_string_list(input_list):
    return [item[0] for item in input_list if item]


# def add_users_to_results_photos(spreadsheet_url, usernames):
#
#     print(f'Функция add_users_to_results_photos -------> Имена для вставки {usernames}')
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#     names = get_names_in_results(spreadsheet_id)
#     names_list = convert_to_string_list(names)
#     target_name = usernames[0]
#     if target_name not in names_list:
#         print(f'Имя {target_name} отсутствует в таблицах результаты и фото, results, производится добавление')
#         # Получение списка всех листов
#         sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
#         sheets = sheet_metadata.get('sheets', '')
#
#         for sheet in sheets:
#             sheet_name = sheet['properties']['title']
#
#             # Проверяем соответствие названия листа шаблону
#             if sheet_name.startswith("results_") or sheet_name.startswith("photos_"):
#                 sheet_id = sheet['properties']['sheetId']
#
#                 start_copy, end_copy, start_paste, end_paste = get_start_end_table(service, spreadsheet_id, sheet_name)
#
#                 for username in usernames:
#                     try:
#                         requests = []
#
#                         # Добавление запроса на копирование и вставку диапазона
#                         requests.append({
#                             "copyPaste": {
#                                 "source": {
#                                     "sheetId": sheet_id,
#                                     "startRowIndex": start_copy,
#                                     "endRowIndex": end_copy,
#                                     "startColumnIndex": 0,
#                                     "endColumnIndex": 26
#                                 },
#                                 "destination": {
#                                     "sheetId": sheet_id,
#                                     "startRowIndex": start_paste,
#                                     "endRowIndex": end_paste,
#                                     "startColumnIndex": 0,
#                                     "endColumnIndex": 26
#                                 },
#                                 "pasteType": "PASTE_NORMAL",
#                                 "pasteOrientation": "NORMAL"
#                             }
#                         })
#
#                         # Добавление запроса на изменение значения в ячейке с именем пользователя
#                         requests.append({
#                             "updateCells": {
#                                 "rows": [{
#                                     "values": [{"userEnteredValue": {"stringValue": username}}]
#                                 }],
#                                 "fields": "userEnteredValue",
#                                 "start": {
#                                     "sheetId": sheet_id,
#                                     "rowIndex": start_paste,
#                                     "columnIndex": 1
#                                 }
#                             }
#                         })
#
#                         # Выполнение запросов
#                         body = {"requests": requests}
#                         service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
#
#                         start_paste = end_paste + 4
#                         end_paste = start_paste + end_copy
#                     except Exception as e:
#                         print(f"An error occurred: {e}")
#                 # sleep(0.01)
#
#         return "Tables copied and pasted successfully for all matching sheets."
#     else:
#         print(f'Имя {target_name} уже есть в таблицах results, добавление не нужно!')


# def add_users_to_results_photos(spreadsheet_url, usernames):
#     print(f'Функция add_users_to_results_photos -------> Имена для вставки {usernames}')
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#     names = get_names_in_results(spreadsheet_id)
#     names_list = convert_to_string_list(names)
#     target_name = usernames[0]
#
#     # Определение текущего месяца и года
#     current_year = datetime.datetime.now().year
#     current_month = datetime.datetime.now().strftime('%B')
#     current_suffix = f"{current_month}_{current_year}"
#
#     if target_name not in names_list:
#         print(f'Имя {target_name} отсутствует в таблицах, производится добавление')
#
#         # Получение списка всех листов
#         sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
#         sheets = sheet_metadata.get('sheets', '')
#
#         for sheet in sheets:
#             sheet_name = sheet['properties']['title']  # Название листа в нижнем регистре
#
#             # Проверяем, заканчивается ли имя листа на текущий месяц и год
#             if sheet_name.endswith(current_suffix):
#                 sheet_id = sheet['properties']['sheetId']
#                 start_copy, end_copy, start_paste, end_paste = get_start_end_table(service, spreadsheet_id, sheet_name)
#
#                 for username in usernames:
#                     try:
#                         requests = []
#
#                         # Добавление запроса на копирование и вставку диапазона
#                         requests.append({
#                             "copyPaste": {
#                                 "source": {
#                                     "sheetId": sheet_id,
#                                     "startRowIndex": start_copy,
#                                     "endRowIndex": end_copy,
#                                     "startColumnIndex": 0,
#                                     "endColumnIndex": 26
#                                 },
#                                 "destination": {
#                                     "sheetId": sheet_id,
#                                     "startRowIndex": start_paste,
#                                     "endRowIndex": end_paste,
#                                     "startColumnIndex": 0,
#                                     "endColumnIndex": 26
#                                 },
#                                 "pasteType": "PASTE_NORMAL",
#                                 "pasteOrientation": "NORMAL"
#                             }
#                         })
#
#                         # Добавление запроса на изменение значения в ячейке с именем пользователя
#                         requests.append({
#                             "updateCells": {
#                                 "rows": [{
#                                     "values": [{"userEnteredValue": {"stringValue": username}}]
#                                 }],
#                                 "fields": "userEnteredValue",
#                                 "start": {
#                                     "sheetId": sheet_id,
#                                     "rowIndex": start_paste,
#                                     "columnIndex": 1
#                                 }
#                             }
#                         })
#
#                         # Выполнение запросов
#                         body = {"requests": requests}
#                         service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
#
#                         # Обновление индексов для следующего добавления
#                         start_paste = end_paste + 4
#                         end_paste = start_paste + end_copy
#                     except Exception as e:
#                         print(f"An error occurred: {e}")
#         print("Пользователи успешно добавлены в таблицы текущего месяца.")
#         return "Пользователи успешно добавлены в таблицы текущего месяца."
#     else:
#         print(f'Имя {target_name} уже есть в таблицах, добавление не нужно!')

def add_users_to_results_photos(spreadsheet_url, usernames):
    print(f'Функция add_users_to_results_photos -------> Имена для вставки {usernames}')
    spreadsheet_id = spreadsheet_url.split("/")[5]
    names = get_names_in_results(spreadsheet_id)
    names_list = convert_to_string_list(names)
    target_name = usernames[0]

    print(f'Полученные имена из таблиц: {names_list}')
    print(f'Целевое имя для добавления: {target_name}')

    # Определение текущего месяца и года
    current_year = datetime.datetime.now().year
    current_month = datetime.datetime.now().strftime('%B')
    current_suffix = f"{current_month}_{current_year}"

    if target_name not in names_list:
        print(f'Имя {target_name} отсутствует в таблицах, производится добавление')

        # Получение списка всех листов
        try:
            sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheets = sheet_metadata.get('sheets', '')
            print(f'Список всех листов {sheets}')
            print(f'Получено метаданных листов: {len(sheets)}')
        except Exception as e:
            print(f"Ошибка при получении метаданных листов: {e}")
            return

        for sheet in sheets:
            sheet_name = sheet['properties']['title']  # Название листа в нижнем регистре
            print(f'Обработка листа: {sheet_name}')

            # Проверяем, заканчивается ли имя листа на текущий месяц и год
            if sheet_name.endswith(current_suffix):
                print(f'Лист {sheet_name} соответствует текущему месяцу и году, продолжаем...')
                sheet_id = sheet['properties']['sheetId']
                try:
                    start_copy, end_copy, start_paste, end_paste = get_start_end_table(service, spreadsheet_id, sheet_name)
                    print(f'Диапазоны: start_copy={start_copy}, end_copy={end_copy}, start_paste={start_paste}, end_paste={end_paste}')
                except Exception as e:
                    print(f"Ошибка при получении диапазонов: {e}")
                    continue

                for username in usernames:
                    try:
                        print(f'Добавление пользователя {username} в лист {sheet_name}')
                        requests = []

                        # Добавление запроса на копирование и вставку диапазона
                        requests.append({
                            "copyPaste": {
                                "source": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": start_copy,
                                    "endRowIndex": end_copy,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": 26
                                },
                                "destination": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": start_paste,
                                    "endRowIndex": end_paste,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": 26
                                },
                                "pasteType": "PASTE_NORMAL",
                                "pasteOrientation": "NORMAL"
                            }
                        })

                        # Добавление запроса на изменение значения в ячейке с именем пользователя
                        requests.append({
                            "updateCells": {
                                "rows": [{
                                    "values": [{"userEnteredValue": {"stringValue": username}}]
                                }],
                                "fields": "userEnteredValue",
                                "start": {
                                    "sheetId": sheet_id,
                                    "rowIndex": start_paste,
                                    "columnIndex": 1
                                }
                            }
                        })

                        # Выполнение запросов
                        body = {"requests": requests}
                        response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                        print(f'Запрос выполнен успешно: {response}')

                        # Обновление индексов для следующего добавления
                        start_paste = end_paste + 4
                        end_paste = start_paste + end_copy
                    except Exception as e:
                        print(f"Ошибка при добавлении пользователя {username}: {e}")
        print("Пользователи успешно добавлены в таблицы текущего месяца.")
        return "Пользователи успешно добавлены в таблицы текущего месяца."
    else:
        print(f'Имя {target_name} уже есть в таблицах, добавление не нужно!')


def add_many_users_to_results_photos(spreadsheet_url, usernames):
    logger.info(f"Начало функции add_users_to_results_photos. Имена для вставки: {usernames}")
    spreadsheet_id = spreadsheet_url.split("/")[5]
    names = get_names_in_results(spreadsheet_id)
    names_list = convert_to_string_list(names)

    logger.debug(f"Полученные имена из таблиц: {names_list}")

    # Определение текущего месяца и года
    current_year = datetime.datetime.now().year
    current_month = datetime.datetime.now().strftime('%B')
    current_suffix = f"{current_month}_{current_year}"

    # Фильтруем только те имена, которых еще нет в таблицах
    new_usernames = [username for username in usernames if username not in names_list]
    if not new_usernames:
        logger.info("Все имена уже существуют в таблицах, добавление не требуется.")
        return "Имена уже существуют в таблицах."

    logger.info(f"Имена для добавления: {new_usernames}")

    # Получение списка всех листов
    try:
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        logger.debug(f"Получено метаданных листов: {len(sheets)}")
    except Exception as e:
        logger.error(f"Ошибка при получении метаданных листов: {e}")
        return

    for sheet in sheets:
        sheet_name = sheet['properties']['title']
        if sheet_name.endswith(current_suffix):
            logger.info(f"Обработка листа: {sheet_name}")
            sheet_id = sheet['properties']['sheetId']
            try:
                start_copy, end_copy, start_paste, end_paste = get_start_end_table(service, spreadsheet_id, sheet_name)
                logger.debug(
                    f"Диапазоны: start_copy={start_copy}, end_copy={end_copy}, start_paste={start_paste}, end_paste={end_paste}")
            except Exception as e:
                logger.error(f"Ошибка при получении диапазонов: {e}")
                continue

            requests = []
            for username in new_usernames:
                # Добавление запроса на копирование диапазона
                requests.append({
                    "copyPaste": {
                        "source": {
                            "sheetId": sheet_id,
                            "startRowIndex": start_copy,
                            "endRowIndex": end_copy,
                            "startColumnIndex": 0,
                            "endColumnIndex": 26
                        },
                        "destination": {
                            "sheetId": sheet_id,
                            "startRowIndex": start_paste,
                            "endRowIndex": end_paste,
                            "startColumnIndex": 0,
                            "endColumnIndex": 26
                        },
                        "pasteType": "PASTE_NORMAL",
                        "pasteOrientation": "NORMAL"
                    }
                })

                # Добавление запроса на изменение значения в ячейке с именем пользователя
                requests.append({
                    "updateCells": {
                        "rows": [{
                            "values": [{"userEnteredValue": {"stringValue": username}}]
                        }],
                        "fields": "userEnteredValue",
                        "start": {
                            "sheetId": sheet_id,
                            "rowIndex": start_paste,
                            "columnIndex": 1
                        }
                    }
                })

                # Обновление индексов для следующего добавления
                start_paste = end_paste + 4
                end_paste = start_paste + (end_copy - start_copy)

            # Выполнение всех запросов за раз
            try:
                body = {"requests": requests}
                response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                logger.info(f"Запрос выполнен успешно: {response}")
            except Exception as e:
                logger.error(f"Ошибка при выполнении batchUpdate: {e}")

    logger.info("Пользователи успешно добавлены в таблицы текущего месяца.")
    return "Пользователи успешно добавлены в таблицы текущего месяца."

