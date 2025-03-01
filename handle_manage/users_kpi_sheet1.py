
import re
from datetime import datetime

from collections import defaultdict
from datetime import datetime
from django.utils.timezone import now
from tsa_app.models import WorkEntry

from googleapiclient.errors import HttpError
from common.get_service import get_service
import string

from common.return_results_curr_month_year import generate_results_filename


def get_values_cells(sh_url, sheet_name):

    """
    Извлекает данные из указанного листа Google Sheets.

    Args:
        sh_url (str): URL таблицы Google Sheets.
        sheet_name (str, optional): Название листа, из которого нужно получить данные.
                                    По умолчанию 'Work Time'.

    Returns:
        list: Список списков, содержащий данные из ячеек таблицы. Если данные не найдены
              или произошла ошибка, возвращается пустой список.

    Raises:
        googleapiclient.errors.HttpError: В случае ошибки запроса к Google Sheets API.
    """
    try:
        # Аутентификация и создание сервиса
        service = get_service()
        sheet_id = sh_url.split("/")[5]  # Получение ID таблицы из URL
        # sheet_name = "Work Time"

        # Получаем данные с листа
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=sheet_name
        ).execute()
        values = result.get("values", [])
        return values  # ✅ Добавляем return, чтобы вернуть данные

    except HttpError as error:
        print(f"An error occurred: {error}")
        return []  # ✅ В случае ошибки возвращаем пустой список, а не None


def extract_last_entries(values):
    """
    Извлекает последние записи после последнего заголовка "Per hour | User Name",
    исключая последнюю строку.

    Args:
        values (list): Двумерный список (таблица) с данными из Google Sheets.

    Returns:
        list: Список строк (списков), содержащий данные после последнего заголовка.
              Если заголовок не найден, возвращается пустой список.
    """
    last_header_index = None

    # Находим индекс последнего заголовка
    for i, row in enumerate(values):
        if row[:2] == ["Per hour", "User Name"]:
            last_header_index = i

    # Если заголовок найден, возвращаем строки после него, кроме последней
    if last_header_index is not None:
        return values[last_header_index + 1:-1]  # Убираем последнюю строку

    return []  # Если заголовок не найден, вернуть пустой список



def column_index_to_letter(index):
    """Преобразует числовой индекс колонки в букву (например, 0 -> 'A', 25 -> 'Z', 26 -> 'AA')."""
    letters = []
    while index >= 0:
        letters.append(string.ascii_uppercase[index % 26])
        index = index // 26 - 1
    return "".join(reversed(letters))


def get_total_cell_indices(values, original_values, sheet_name="Work Time"):
    """
    Определяет координаты (A1-нотация) ячеек для итоговых значений и
    связывает их с соответствующими именами работников.

    Args:
        values (list): Отфильтрованный список строк с данными (без заголовков).
        original_values (list): Исходные данные из Google Sheets (включая заголовки).
        sheet_name (str, optional): Название листа, используемого в A1-нотации.
                                    По умолчанию "Work Time".

    Returns:
        list: Список кортежей вида (имя работника, A1-адрес ячейки итогового значения).
    """
    total_cells = []

    # Получаем список индексов строк в оригинальных данных
    row_indices_in_original = [
        original_values.index(row) + 1 for row in values  # +1, так как Google Sheets нумерует с 1
    ]

    for row_index, row in zip(row_indices_in_original, values):
        if len(row) > 1:  # Проверяем, что строка не пустая
            name = row[1]
            total_index = len(row) - 2  # Индекс предпоследнего элемента
            col_letter = column_index_to_letter(total_index)  # Конвертируем индекс в букву колонки
            cell_address = f"{sheet_name}!{col_letter}{row_index}"  # Формируем A1-нотацию
            total_cells.append((name, cell_address))

    return total_cells


def get_total_earned_cells(values):
    """
    Извлекает координаты (A1-нотация) ячеек с суммарным заработком ('Total earned')
    и связывает их с соответствующими именами пользователей.

    Args:
        values (list): Двумерный список, содержащий данные из Google Sheets.

    Returns:
        list: Список кортежей (имя пользователя, A1-адрес ячейки 'Total earned'),
              исключая первую запись, если она содержит 'Example'.
    """
    total_earned_cells = []
    sheet_name = generate_results_filename()
    # sheet_name = 'results_January_2025'  # для проверки на результатах прошлого месяца
    print(f'Sheet name: {sheet_name}')
    for i, row in enumerate(values):
        if row and row[0] == "Name" and len(row) > 1:
            name = row[1]  # Имя пользователя
        elif row and row[0] == "Total earned" and len(row) > 1:
            col_index = 1  # 'Total earned' всегда во второй колонке (индекс 1)
            col_letter = column_index_to_letter(col_index)
            cell_address = f"{sheet_name}!{col_letter}{i + 1}"  # A1-нотация
            total_earned_cells.append((name, cell_address))

    # Отбрасываем первый кортеж, если это 'Example'
    if total_earned_cells and total_earned_cells[0][0] == "Example":
        total_earned_cells = total_earned_cells[1:]

    return total_earned_cells


def create_formula_list(work_time_list, results_list):
    """
    Создает список кортежей, где каждому имени пользователя сопоставляется формула деления
    значения из results_list на соответствующее значение из work_time_list.

    Args:
        work_time_list (list): Список кортежей (имя пользователя, A1-адрес ячейки в 'Work Time').
        results_list (list): Список кортежей (имя пользователя, A1-адрес ячейки в 'Results').

    Returns:
        list: Список кортежей (имя пользователя, формула для деления результата на время работы).
              Например: [('Ivan', '=results_January_2025!B211 / Work Time!Q11')]
    """
    formula_list = []
    print(f'create_formula_list переданы аргументы:\n'
          f'work_time_list: {work_time_list}\n'
          f'results_list: {results_list}')
    # Преобразуем второй список (results_list) в словарь для быстрого поиска
    results_dict = {name: cell for name, cell in results_list}

    # Перебираем первый список (work_time_list)
    for name, work_time_cell in work_time_list:
        if name in results_dict:
            result_cell = results_dict[name]
            formula = f"{result_cell} / {work_time_cell}"
            formula_list.append((name, formula))
    print(f'create_formula_list возвращает: {formula_list}')
    return formula_list


def get_last_non_empty_row(spreadsheet_url, sheet_name):
    """
    Функция находит последнюю записанную ячейку в листе, возвращает индекс этой ячейки.
    Так как Google API оптимизирует запросы, то выбрав ячейки диапазоном вида "{sheet_name}!A1:A",
    API вернет ячейки от 1 до последней заполненной.
    """
    service = get_service()
    spreadsheet_id = spreadsheet_url.split("/")[5]
    range_name = f"{sheet_name}!B1:B"  # Запрашиваем данные только из первого столбца

    # Получаем данные столбца
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    # Возвращаем индекс последней заполненной строки
    last_row_index = len(values) if values else 0  # Если нет данных, возвращаем 0

    print(f"Last non-empty row in column A is {last_row_index}")
    return last_row_index


def write_formulas_to_sheet(sh_url, formula_list, sheet_name):
    """Записывает формулы в Google Sheets, начиная с первой пустой строки +3.

    Аргументы:
        sh_url (str): URL таблицы Google Sheets.
        sheet_name (str): Название листа для записи.
        formula_list (list): Список кортежей (имя, формула) для записи.
    """
    try:
        # Аутентификация и создание сервиса
        service = get_service()
        sheet_id = sh_url.split("/")[5]  # Извлекаем ID таблицы
        start_row = get_last_non_empty_row(sh_url, sheet_name) + 3

        # Получаем текущий месяц
        current_month = datetime.now().strftime("%B %Y")
        # current_month = 'January 2025'

        # Функция для корректировки имен листов в формулах (добавляет одинарные кавычки)
        def fix_sheet_names(formula):
            return re.sub(r'([\w\d]+ [\w\d]+)', r"'\1'", formula)  # Одинарные кавычки

        # Формируем данные для записи
        values = [[current_month, ""]]  # Первая строка — название месяца
        values.extend([[name, f"=ROUND({fix_sheet_names(formula)}; 2)"] for name, formula in formula_list])

        # Определяем диапазон записи
        end_row = start_row + len(values) - 1
        range_to_update = f"{sheet_name}!B{start_row}:C{end_row}"

        # Записываем данные в Google Sheets
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=range_to_update,
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()

        print("Формулы успешно записаны.")

    except HttpError as error:
        print(f"Произошла ошибка: {error}")


def get_work_entries_for_current_month(sh_url: str, build_object_id: int):
    """
    Извлекает все записи WorkEntry за текущий месяц для указанного объекта строительства,
    возвращая имя пользователя и сумму всех отработанных часов за все дни.
    :param sh_url: Ссылка на Google Sheets (пока не используется в логике, но может пригодиться позже)
    :param build_object_id: ID строительного объекта
    :return: Список кортежей (имя пользователя, сумма всех отработанных часов)
    """
    today = now().date()
    first_day_of_month = today.replace(day=1)

    work_entries = WorkEntry.objects.filter(
        build_object_id=build_object_id,
        date__gte=first_day_of_month,
        date__lte=today
    ).values("worker__name", "worked_hours")

    hours_summary = defaultdict(float)
    for entry in work_entries:
        hours_summary[entry["worker__name"]] += float(entry["worked_hours"])

    result = list(hours_summary.items())
    print(f'get_work_entries_for_current_month return result: {result}')
    return result


# def get_work_entries_for_current_month(sh_url: str, build_object_id: int, test_month=None):
#     """
#     Извлекает все записи WorkEntry за указанный месяц для указанного объекта строительства.
#     Если test_month указан, использует его вместо текущего месяца.
#     """
#     today = now().date()
#
#     if test_month:  # Например, test_month = "2025-01"
#         year, month = map(int, test_month.split("-"))
#         first_day_of_month = datetime(year, month, 1).date()
#         last_day_of_month = datetime(year, month + 1, 1).date() if month < 12 else datetime(year + 1, 1, 1).date()
#     else:
#         first_day_of_month = today.replace(day=1)
#         last_day_of_month = today
#
#     work_entries = WorkEntry.objects.filter(
#         build_object_id=build_object_id,
#         date__gte=first_day_of_month,
#         date__lt=last_day_of_month  # Используем `<` для корректного охвата месяца
#     ).values("worker__name", "worked_hours")
#     print(f'get_work_entries_for_current_month return work_entries: {work_entries}')
#
#     hours_summary = defaultdict(float)
#     for entry in work_entries:
#         hours_summary[entry["worker__name"]] += float(entry["worked_hours"])
#
#     result = list(hours_summary.items())
#     print(f'get_work_entries_for_current_month return result: {result}')
#     return result


def run_write_kpi_users_in_sheet1(sh_url, build_object_id):
    # values = get_values_cells(sh_url, sheet_name='Work Time')
    #
    # filtered_values = extract_last_entries(values)
    # total_cells = get_total_cell_indices(filtered_values, values)
    # total_cells = get_work_entries_for_current_month(sh_url='dfghdf', build_object_id=20)
    # Тестируем для января 2024
    total_cells = get_work_entries_for_current_month(sh_url, build_object_id)
    # values = get_values_cells(sh_url, sheet_name='results_January_2025')
    values = get_values_cells(sh_url, sheet_name=generate_results_filename())
    total_earned_cells = get_total_earned_cells(values)
    formula_list = create_formula_list(results_list=total_earned_cells, work_time_list=total_cells)
    write_formulas_to_sheet(sh_url, formula_list, sheet_name='Лист1')

