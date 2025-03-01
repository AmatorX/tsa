from collections import defaultdict
from datetime import datetime
from time import sleep

import googleapiclient
import logging
# from logging_config import setup_logging

from kpi_utils.update_users_kpi_values import update_kpi_values
from common.service import service


# Здесь подготавливаются данные для записи в таблицах Users KPI

service = service.get_service()
logger = logging.getLogger(__name__)



# def get_sheet_id(service, spreadsheet_id, sheet_name):
#     # Получение информации о всех листах в таблице
#     sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
#     sheets = sheet_metadata.get('sheets', '')
#
#     # Поиск идентификатора листа по имени
#     for sheet in sheets:
#         if sheet['properties']['title'] == sheet_name:
#             return sheet['properties']['sheetId']
#     logging.error(f"Sheet with name '{sheet_name}' not found.")
#     raise ValueError(f"Sheet with name '{sheet_name}' not found.")
#
#
# def get_current_month_and_day():
#     now = datetime.now()
#     current_year = now.year
#     current_month = now.strftime("%B")  # Получаем название текущего месяца
#     current_day = now.day  # Получаем номер текущего дня
#     return current_month, current_day, current_year
#
#
# def find_sheet_id_by_name(spreadsheet_url, service, sheet_name='Work Time'):
#     """
#     Получение ID листа 'Work Time'
#     """
#     # Извлечение spreadsheet_id из URL
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#
#     # Получение списка всех листов в таблице
#     sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
#     sheets = sheet_metadata.get('sheets', '')
#
#     # Поиск ID листа по названию
#     for sheet in sheets:
#         if sheet['properties']['title'] == sheet_name:
#             return sheet['properties']['sheetId']
#
#     # Если лист с таким названием не найден, возвращаем None
#     return None
#
#
# # Шаг 4
# def find_cell_by_date(spreadsheet_url):
#     """
#     Находит ячейку с текущей датой (в формате "Месяц День") на листе "Work Time" в Google Sheets.
#
#     Функция извлекает идентификатор таблицы из предоставленного URL, проверяет наличие листа с именем "Work Time",
#     и ищет текущую дату, соответствующую формату "Месяц День" (например, "December 5"), в диапазоне ячеек `A:Z`.
#
#     Args:
#         spreadsheet_url (str): URL Google Sheets, содержащий таблицу для поиска.
#
#     Returns:
#         tuple | None:
#             Если дата найдена:
#                 - str: `spreadsheet_id` — идентификатор таблицы.
#                 - int: Номер строки, где находится дата (индексация с 1).
#                 - int: Номер столбца, где находится дата (индексация с 1).
#             Если дата не найдена или лист "Work Time" отсутствует:
#                 - None.
#     """
#
#     # Извлечение spreadsheet_id из URL
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#
#     date = get_current_month_and_day()
#     month_day = date[0] + ' ' + str(date[1])
#     logging.debug(f'Текущая дата: {month_day}')
#
#     try:
#         # Получение метаданных таблицы
#         sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
#         sleep(2)
#         sheets = sheet_metadata.get('sheets', '')
#         sheet_names = [sheet['properties']['title'] for sheet in sheets]
#         logging.debug(f'Доступные листы: {sheet_names}')
#
#         # Проверьте, существует ли лист 'Work Time'
#         if 'Work Time' not in sheet_names:
#             logging.debug(f"Лист 'Work Time' не найден в таблице. Доступные листы: {sheet_names}")
#             return None
#
#         # Определение диапазона для поиска
#         range_name = 'Work Time!A:Z'
#
#         # Получение данных из листа
#         result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
#         sleep(2)
#         values = result.get('values', [])
#         logging.debug(f'Values: {values}')
#
#         # Поиск ячейки с заданным содержимым
#         for row_index, row in enumerate(values):
#             for col_index, value in enumerate(row):
#                 if month_day == value:
#                     logging.info(f'Функция find_cell_by_date возвращает\n'
#                                  f'spreadsheet_id: {spreadsheet_id}, row_index: {row_index + 1}, col_index: {col_index + 1}')
#                     return spreadsheet_id, row_index + 1, col_index + 1
#
#     except googleapiclient.errors.HttpError as err:
#         logging.error(f"An error occurred: {err}")
#         if err.resp.status == 403:
#             logging.error("Access denied. Make sure the service account has access to the spreadsheet.")
#         return None
#
#     return None
#
#
# # Шаг 3
# def calculate_salary_by_names(spreadsheet_url, user_info_list):
#     """
#     Рассчитывает заработную плату пользователей на основе их часовой ставки и рабочего времени, указанного в Google Sheets.
#
#     Функция ищет имена пользователей в столбце B листа "Work Time", начиная с строки после текущей даты.
#     Затем находит количество отработанных часов для каждого пользователя в соответствующей колонке текущей даты.
#     Заработная плата рассчитывается как произведение количества часов на ставку пользователя.
#
#     Args:
#         spreadsheet_url (str): URL Google Sheets, содержащий данные о рабочем времени и ставках пользователей.
#         user_info_list (list of tuples): Список пользователей в формате [(имя, ставка), ...], где:
#             - имя (str): Имя пользователя.
#             - ставка (float или str): Почасовая ставка пользователя.
#
#     Returns:
#         tuple:
#             - str: URL таблицы (`spreadsheet_url`), из которой были получены данные.
#             - list of lists: Список с рассчитанной зарплатой в формате [[имя, заработная плата], ...].
#
#     Raises:
#         ValueError: Если значения часов или ставки не могут быть преобразованы в числа.
#         googleapiclient.errors.HttpError: Если произошла ошибка HTTP при работе с Google API.
#
#     Notes:
#         - Если имя пользователя отсутствует в таблице, ему присваивается заработная плата 0.
#         - Если данные о часах или ставке некорректны, они обрабатываются как 0 по умолчанию.
#     """
#
#     logging.info(f'Функция calculate_salary_by_names запущена, [ЩАГ № 3]')
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#     salaries = []
#
#     # Находим индекс ячейки с текущей датой
#     _, date_row_index, date_col_index = find_cell_by_date(spreadsheet_url)
#
#     # Определение диапазона для поиска имен пользователей начиная со следующей строки после даты
#     range_end = date_row_index + len(user_info_list) + 1  # +1 для учета дополнительной строки после списка пользователей
#     range_name = f'Work Time!B{date_row_index + 1}:B{range_end}'
#
#     # Получение данных из листа
#     result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
#     sleep(2)
#     values = result.get('values', [])
#
#     for user in user_info_list:
#         user_name, user_salary = user
#         user_found = False
#
#         # Поиск имени пользователя в списке значений
#         for row_index, row in enumerate(values):
#             if row and row[0] == user_name:
#                 user_found = True
#                 # Получение часов из соответствующего столбца даты для найденного пользователя
#                 hours_range = f'Work Time!{chr(65 + date_col_index - 1)}{date_row_index + row_index + 1}'
#                 hours_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
#                                                                    range=hours_range).execute()
#                 sleep(2)
#                 hours = hours_result.get('values', [[0]])[0][0]
#
#                 # Проверка и предобработка данных
#                 try:
#                     hours = float(hours)
#                 except ValueError:
#                     hours = 0.0  # или любое другое значение по умолчанию
#
#                 try:
#                     user_salary = float(user_salary)
#                 except ValueError:
#                     user_salary = 0.0  # или любое другое значение по умолчанию
#
#                 # Вычисление заработной платы
#                 calculated_salary = hours * user_salary
#                 salaries.append([user_name, calculated_salary])
#                 break  # Прерываем цикл, так как нашли пользователя и рассчитали его зарплату
#
#         if not user_found:
#             # Если пользователь не найден, добавляем его с зарплатой 0
#             salaries.append([user_name, 0.0])
#     logging.info(f'[ШАГ № 3] возвращаем {salaries}')
#     return spreadsheet_url, salaries
#
#
# # Шаг 2
# def get_kpi_for_work_time(sh_and_users):
#     """
#         Рассчитывает KPI по рабочему времени для нескольких Google Sheets.
#
#         Функция принимает список данных, содержащий URL таблиц и информацию о пользователях, и вызывает
#         функцию `calculate_salary_by_names` для расчёта заработной платы пользователей на основе их часовой ставки
#         и рабочего времени. Результаты собираются в единый список.
#
#         Args:
#             sh_and_users (list of tuples): Список данных в формате [(spreadsheet_url, user_info_list), ...], где:
#                 - spreadsheet_url (str): URL Google Sheets с данными о рабочем времени.
#                 - user_info_list (list of tuples): Список пользователей в формате [(имя, ставка), ...].
#
#         Returns:
#             list of tuples:
#                 - Каждый элемент списка — это результат работы `calculate_salary_by_names` для одной таблицы, включающий:
#                     - str: URL таблицы (`spreadsheet_url`).
#                     - list of lists: Рассчитанные заработные платы пользователей в формате [[имя, заработная плата], ...].
#     """
#     logging.info(f'Функция get_kpi_for_work_time запущена, [ШАГ № 2]')
#     work_time_kpi = []
#     for data_list in sh_and_users:
#         spreadsheet_url = data_list[0]
#         user_info_list = data_list[1]
#         name_salary_work_time = calculate_salary_by_names(spreadsheet_url, user_info_list)
#         work_time_kpi.append(name_salary_work_time)
#         sleep(2)
#     logging.info(f'KPI WORk TIME\n{work_time_kpi}')
#     return work_time_kpi
#
#
# # Рассчет для KPI по таблицам результатов
#
# def get_results_current_month_table(service, spreadsheet_url):
#     """
#         Генерирует имя таблицы для текущего месяца на основе текущей даты.
#
#         Функция формирует имя таблицы в формате `results_<месяц>_<год>`, где:
#         - `<месяц>` — текущий месяц (например, "December").
#         - `<год>` — текущий год (например, "2024").
#
#         Args:
#             service (googleapiclient.discovery.Resource): Сервис Google Sheets API для работы с таблицей.
#                 (Не используется в текущей реализации, но может быть полезен для расширения функциональности.)
#             spreadsheet_url (str): URL Google Sheets, где предполагается создание или использование таблицы.
#
#         Returns:
#             str: Имя таблицы в формате `results_<месяц>_<год>`.
#     """
#     current_date = get_current_month_and_day()
#     sheet_name = f'results_{current_date[0]}_{current_date[2]}'
#     return sheet_name
#
#
# def get_indexes_cell_name(service, spreadsheet_url, sheet_name, name):
#     """
#     Находит строку, содержащую заданное имя, в указанном листе Google Sheets.
#
#     Функция выполняет поиск имени в столбце B указанного листа и возвращает индекс строки, где найдено совпадение.
#     Если имя не найдено, возвращает `None`.
#
#     Args:
#         service (googleapiclient.discovery.Resource): Сервис Google Sheets API для работы с таблицей.
#         spreadsheet_url (str): URL Google Sheets, где будет выполнен поиск.
#         sheet_name (str): Имя листа, на котором выполняется поиск.
#         name (str): Имя, которое нужно найти в столбце B.
#
#     Returns:
#         tuple or None:
#             - Если имя найдено: `(str, int)` — кортеж из имени и индекса строки (нумерация с 1).
#             - Если имя не найдено: `None`.
#     """
#
#     spreadsheet_id = spreadsheet_url.split("/")[5]
#     # Определение диапазона для поиска в столбце B
#     range_name = f"{sheet_name}!B:B"
#
#     # Получение данных из столбца B
#     result = service.spreadsheets().values().get(
#         spreadsheetId=spreadsheet_id,
#         range=range_name
#     ).execute()
#     sleep(2)
#
#     values = result.get('values', [])
#
#     # Поиск ячейки с заданным именем
#     for row_index, row in enumerate(values):
#         if row and row[0] == name:
#             logging.debug(f"Найдено совпадение в строке {row_index + 1}, столбце 2, для имени {name}")
#             return name, row_index + 1
#
#     # Если имя не найдено, возвращаем None
#     return None
#
#
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
#
#
# def calculate_total_material_cost(service, spreadsheet_url, sheet_name, name, start_row):
#     """
#     Вычисляет общую стоимость материалов для указанного имени на основе данных в таблице Google Sheets.
#
#     Функция извлекает значения стоимости материалов и количества, умножает их и возвращает итоговую стоимость.
#
#     Args:
#         service (googleapiclient.discovery.Resource): Сервис Google Sheets API для работы с таблицей.
#         spreadsheet_url (str): URL Google Sheets, где находятся данные.
#         sheet_name (str): Имя листа, на котором находятся данные.
#         name (str): Имя пользователя или объекта, для которого рассчитывается стоимость материалов.
#         start_row (int): Номер строки (начиная с 1), откуда начинается расчет.
#
#     Returns:
#         tuple:
#             - `name` (str): Имя пользователя или объекта.
#             - `total_cost` (float): Общая стоимость материалов.
#     """
#     # Определение текущего дня
#     now = datetime.now()
#     current_day = now.day
#     spreadsheet_id=spreadsheet_url.split("/")[5]
#     # Определение диапазона для material_price
#     material_price_range = f"{sheet_name}!B{start_row + 2}:Z{start_row + 2}"
#     # Получение данных из листа для material_price
#     material_price_result = service.spreadsheets().values().get(
#         spreadsheetId=spreadsheet_id,
#         range=material_price_range
#     ).execute()
#     sleep(2)
#     material_price_values = material_price_result.get('values', [[]])[0]
#
#     # Определение диапазона для salary
#     salary_range = f"{sheet_name}!B{start_row + 4 + current_day}:Z{start_row + 4 + current_day}"
#     # Получение данных из листа для salary
#     salary_result = service.spreadsheets().values().get(
#         spreadsheetId=spreadsheet_id,
#         range=salary_range
#     ).execute()
#     sleep(2)
#     salary_values = salary_result.get('values', [[]])[0]
#
#     # Вычисление суммы произведений
#     total_cost = 0
#     for material_price, salary in zip(material_price_values, salary_values):
#         if material_price:  # Проверка, что ячейка material_price не пустая
#             material_price = convert_to_float(material_price)
#             salary = convert_to_float(salary)
#             total_cost += material_price * salary
#         else:
#             break  # Прекращаем цикл, если встретили пустую ячейку в material_price
#     logging.debug(f"Функция calculate_total_material_cost вернула {name}: {total_cost}")
#     sleep(2)
#     return name, total_cost
#
#
# #############################################################################################################################
# def get_kpi_for_results(sh_and_users):
#     """
#     Генерирует KPI по результатам текущего месяца для списка пользователей и таблиц Google Sheets.
#
#     Функция рассчитывает общую стоимость материалов для каждого пользователя на основе данных,
#     извлеченных из таблиц Google Sheets, соответствующих текущему месяцу.
#     Итоговые данные возвращаются в виде списка кортежей.
#
#     Args:
#         sh_and_users (list): Список данных, содержащий:
#             - `spreadsheet_url` (str): URL Google Sheets.
#             - `user_info_list` (list): Список данных о пользователях, каждый элемент содержит:
#                 - `name` (str): Имя пользователя.
#                 - Дополнительные данные, которые могут быть использованы в других функциях.
#
#     Returns:
#         list: Список кортежей, где каждый кортеж содержит:
#             - `spreadsheet_url` (str): URL таблицы Google Sheets.
#             - `user_results` (list): Список списков, каждый из которых содержит:
#                 - `name` (str): Имя пользователя.
#                 - `total_cost` (float): Общая стоимость материалов.
#
#     Workflow:
#         1. Определение имени листа текущего месяца с помощью `get_results_current_month_table`.
#         2. Для каждого пользователя:
#             - Поиск начальной строки данных с помощью `get_indexes_cell_name`.
#             - Расчет общей стоимости материалов с использованием `calculate_total_material_cost`.
#         3. Результаты каждого пользователя добавляются в `user_results` для конкретного URL.
#         4. Итоговые данные по каждому URL сохраняются в словаре `kpi_results`.
#         5. Преобразование словаря `kpi_results` в список кортежей.
#
#     Example:
#         Входные данные:
#             sh_and_users = [
#                 (
#                     "https://docs.google.com/spreadsheets/d/spreadsheet_id/edit",
#                     [["User1"], ["User2"]]
#                 )
#             ]
#
#         Выходные данные:
#             [
#                 (
#                     "https://docs.google.com/spreadsheets/d/spreadsheet_id/edit",
#                     [["User1", 500.0], ["User2", 300.0]]
#                 )
#             ]
#     """
#     logging.info(f'Функция get_kpi_for_results запущена, [ШАГ № 4]')
#     # service = get_service()
#     kpi_results = {}
#     for data_list in sh_and_users:
#         spreadsheet_url = data_list[0]
#         user_info_list = data_list[1]
#         sheet_name = get_results_current_month_table(service=service, spreadsheet_url=spreadsheet_url)
#         user_results = []
#         for user_data in user_info_list:
#             name = user_data[0]
#             name, start_row = get_indexes_cell_name(service=service, spreadsheet_url=spreadsheet_url, sheet_name=sheet_name, name=name)
#             sleep(1)
#             name_salary_results = calculate_total_material_cost(
#                 service=service,
#                 spreadsheet_url=spreadsheet_url,
#                 sheet_name=sheet_name,
#                 name=name,
#                 start_row=start_row)
#             sleep(2)
#             user_results.append(list(name_salary_results))
#         kpi_results[spreadsheet_url] = user_results
#         sleep(1)
#
#     # Преобразование словаря в список кортежей
#     kpi_results = [(url, data) for url, data in kpi_results.items()]
#     logging.info(f'Функция get_kpi_for_results вернула -> {kpi_results}')
#     return kpi_results
# #######################################################################################################################
#
#
# # 1 Шаг
# def update_data_users_kpi(sh_and_users):
#     """
#     Обновляет данные KPI для пользователей, рассчитывая разницу между временем работы и результатами.
#
#     Функция выполняет следующие шаги:
#     1. Вызывает `get_kpi_for_work_time` для получения KPI по времени работы.
#     2. Вызывает `get_kpi_for_results` для получения KPI по результатам.
#     3. Преобразует результаты в словари для удобства доступа.
#     4. Для каждого пользователя рассчитывает разницу между значением результатов и времени работы.
#     5. Формирует список обновлённых данных для дальнейшего использования.
#     6. Передаёт данные в функцию `update_kpi_values` для обновления KPI в таблицах.
#
#     Args:
#         sh_and_users (list): Список данных, содержащий:
#             - `spreadsheet_url` (str): URL таблицы Google Sheets.
#             - `user_info_list` (list): Список данных о пользователях, каждый элемент содержит:
#                 - `name` (str): Имя пользователя.
#                 - Дополнительные данные, которые могут быть использованы в других функциях.
#
#     Returns:
#         list: Список кортежей, содержащий данные KPI по результатам, где каждый кортеж включает:
#             - `spreadsheet_url` (str): URL таблицы Google Sheets.
#             - `user_results` (list): Список списков, каждый из которых содержит:
#                 - `name` (str): Имя пользователя.
#                 - `result_value` (float): Итоговое значение результата.
#
#     Workflow:
#         1. Расчёт KPI времени работы (`get_kpi_for_work_time`).
#         2. Расчёт KPI результатов (`get_kpi_for_results`).
#         3. Сопоставление данных по URL и именам пользователей.
#         4. Вычисление разницы между результатами и временем работы для каждого пользователя.
#         5. Обновление KPI данных через `update_kpi_values`.
#
#     Example:
#         Входные данные:
#             sh_and_users = [
#                 (
#                     "https://docs.google.com/spreadsheets/d/spreadsheet_id/edit",
#                     [["User1"], ["User2"]]
#                 )
#             ]
#
#         Выходные данные:
#             [
#                 (
#                     "https://docs.google.com/spreadsheets/d/spreadsheet_id/edit",
#                     [["User1", -50.0], ["User2", 30.0]]
#                 )
#             ]
#     """
#     logging.info(f'Функция update_data_users_kpi запущена [ЩАГ № 1]')
#     kpi_work_time = get_kpi_for_work_time(sh_and_users=sh_and_users)
#     kpi_results = get_kpi_for_results(sh_and_users=sh_and_users)
#     # Преобразование списков в словари для удобства обращения
#     dict_work_time = {url: dict(data) for url, data in kpi_work_time}
#     dict_results = {url: dict(data) for url, data in kpi_results}
#
#     data_list = []
#
#     # Итерация по первому словарю
#     for url, data in dict_work_time.items():
#         # Проверка, существует ли соответствующий URL во втором словаре
#         if url in dict_results:
#             # Список для хранения обновлённых данных по текущему URL
#             updated_data = []
#             for name, value in data.items():
#                 # Вычисление разницы значений для соответствующего имени, если оно существует во втором словаре
#                 if name in dict_results[url]:
#                     # diff = value - dict_results[url][name]
#                     diff = dict_results[url][name] - value
#                     updated_data.append([name, diff])
#             # Добавление обновлённых данных в итоговый результат
#             data_list.append((url, updated_data))
#     logging.info(f'Функция update_data_users_kpi вернула ----> {data_list}\nЗапуск обновления users KPI')
#     update_kpi_values(data_list=data_list)
#     logging.info(f'kpi_results: {kpi_results} !!!!!!!!!!!!!!!!!!!!!!!')
#     return kpi_results

###########################################################################################################################
from datetime import date
from tsa_app.models import Material, WorkEntry


def get_materials_with_prices():
    """
    Возвращает словарь материалов и их стоимости из базы данных.
    Ключ: имя материала.
    Значение: стоимость материала в формате float (или None, если не указана).
    """
    materials = Material.objects.all()
    return {material.name: float(material.price) if material.price is not None else None for material in materials}


def calculate_material_earnings(entry, materials_prices):
    """
    Для каждого материала в поле materials_used сотрудника рассчитывает его заработок
    на основе расценок материалов и возвращает общую сумму заработанных денег.

    :param entry: Запись о работе сотрудника (WorkEntry).
    :param materials_prices: Словарь с расценками материалов {название_материала: цена}.
    :return: Сумма заработка за все материалы.
    """
    total_earnings = 0
    materials_used = entry.get_materials_used()

    for material, quantity in materials_used.items():
        if material in materials_prices:
            material_price = materials_prices[material]
            try:
                # Приводим quantity к float, если это необходимо
                quantity = float(quantity)
                total_earnings += quantity * material_price
            except ValueError:
                # Если не удается преобразовать quantity в float, пропускаем этот материал
                continue

    return total_earnings


def get_today_work_entries_with_material_earnings(materials_prices):
    """
    Извлекает все записи WorkEntry за сегодняшний день и вычисляет заработок
    сотрудников по всем материалам, а также рассчитывает разницу между доходом
    по материалам и дневной заработной платой.

    Возвращает список словарей с подробностями:
    - worker: имя сотрудника
    - salary_per_day: дневной заработок (salary * worked_hours)
    - total_materials_earnings: суммарный заработок по материалам
    - earnings: разница между заработком по материалам и дневной зарплатой
    - date: дата записи
    """
    today = date.today()
    work_entries = WorkEntry.objects.filter(date=today).select_related('worker', 'build_object')
    # work_entries = WorkEntry.objects.select_related('worker', 'build_object') # использую для теста брать всех а не только за текущий день

    result = []
    for entry in work_entries:
        salary = float(entry.worker.salary) if entry.worker and entry.worker.salary else 0
        worked_hours = entry.get_worked_hours_as_float()
        salary_per_day = salary * worked_hours  # Вычисление дневного заработка

        total_materials_earnings = calculate_material_earnings(entry, materials_prices)

        # Расчет разницы между заработком по материалам и дневной зарплатой
        earnings = total_materials_earnings - salary_per_day

        result.append({
            'build_object_sh_url': entry.build_object.sh_url if entry.build_object else "Unknown URL",
            'worker': entry.worker.name if entry.worker else "Unknown Worker",
            'earnings': earnings,
        })

    return result


def group_by_build_object_sh_url(result):
    """
    Группирует данные по ссылке на объект строительства (sh_url),
    возвращая список кортежей с URL и списком работников с их заработками.

    :param result: Список словарей с данными о работниках и их заработке.
    :return: Список кортежей в формате [(build_object_sh_url, [['worker_name', earnings], ...]), ...]
    """
    grouped_data = {}

    # Группируем данные по build_object_sh_url
    for entry in result:
        build_object_sh_url = entry['build_object_sh_url']
        worker_name = entry['worker']
        earnings = entry['earnings']

        if build_object_sh_url not in grouped_data:
            grouped_data[build_object_sh_url] = []

        grouped_data[build_object_sh_url].append([worker_name, earnings])

    # Преобразуем сгруппированные данные в список кортежей
    grouped_list = [(url, data) for url, data in grouped_data.items()]

    return grouped_list


def earned_today_for_installing_materials():
    """
    Извлекает данные WorkEntry за сегодняшний день, рассчитывает сумму заработка за материалы
    и группирует пользователей по sh_url объектов строительства.

    Возвращает список кортежей:
    - первый элемент: sh_url объекта строительства
    - второй элемент: список списков, где каждый внутренний список содержит имя пользователя и его сумму

    :return: Список кортежей [(sh_url, [['worker_name', total_materials_earnings], ...]), ...]
    """
    today = date.today()
    work_entries = WorkEntry.objects.filter(date=today).select_related('worker', 'build_object')
    materials_prices = get_materials_with_prices()  # Получаем расценки материалов

    grouped_data = defaultdict(list)

    for entry in work_entries:
        if not entry.worker or not entry.build_object:
            continue  # Пропускаем записи с отсутствующим работником или объектом строительства

        worker_name = entry.worker.name
        sh_url = entry.build_object.sh_url
        total_materials_earnings = calculate_material_earnings(entry, materials_prices)

        grouped_data[sh_url].append([worker_name, total_materials_earnings])

    # Преобразуем grouped_data в список кортежей
    grouped_list = [(sh_url, workers) for sh_url, workers in grouped_data.items()]

    return grouped_list

def update_data_users_kpi():
    materials_prices = get_materials_with_prices()
    logger.info(f'Функция update_data_users_kpi запущена\n materials_prices: {materials_prices}')
    print(materials_prices)
    today_entries = get_today_work_entries_with_material_earnings(materials_prices)
    for entry in today_entries:
        print(entry)
    data_to_save = group_by_build_object_sh_url(today_entries)

    print(data_to_save)
    update_kpi_values(data_to_save)
    return data_to_save

