from time import sleep
# from days_list import create_date_list
# from get_service import get_service
from common.days_list import create_date_list
from common.get_service import get_service
import datetime


def get_names_in_last_chunk(input_list):
    # Проходим список в обратном порядке и добавляем элементы в новый список, пока не встретим пустой элемент
    filtered_list = []
    for item in reversed(input_list):
        if item:  # Если элемент не пустой
            filtered_list.append(item[0])  # Добавляем первый элемент из вложенного списка
        else:
            break  # Прерываем цикл, если встретили пустой элемент
    # Возвращаем список в правильном порядке
    return list(reversed(filtered_list)) if filtered_list else [item[0] for item in input_list if item]


def find_last_index_with_user_name(data):
    """
    Перебор массива values, для поиска последнего списка со значением "User Name"
    Возвращает индекс следующего после него списка
    Будет применяться для определения индекса певой строки для сумм в столбцах Total и Salary
    """
    for i in range(len(data) - 1, -1, -1):  # Перебор массива в обратном порядке
        if 'User Name' in data[i]:
            return i + 2
    return None  # Если 'User Name' не найден


def get_names_in_work_time(spreadsheet_id, service, sheet_name):
    """
    Возвращает значения ячеек
    """
    # Определение диапазона для чтения данных из столбца B
    range_name = f"{sheet_name}!B:B"

    # Получение данных из столбца B
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    print(f'Список ячеек для поиска пустой строки\n'
          f'--------------->'
          f' {values}')
    return values


def get_start_row(values):
    start_row = len(values) + 1
    print(f'Индекс строки для записи {start_row}')
    return start_row


def append_user_to_work_time(spreadsheet_url, users, sheet_name='Work Time'):
    service = get_service()
    print('Старт функции добавления пользователей в work_time')
    print(f'Список пользователей для добавления {users}')

    if users is None:
        pass
        # TODO: Добавить проверку переданы ли пользователи или нет, если нет, азять из  базы

    print('Функция добавления пользователей в таблицы work time запустилась')

    # Извлечение spreadsheet_id из URL
    spreadsheet_id = spreadsheet_url.split("/")[5]
    values = get_names_in_work_time(spreadsheet_id, service, sheet_name)
    start_row = get_start_row(values)
    # Задаем стартовую строку для добавления формул Total и Salary
    start_row_total_and_salary = find_last_index_with_user_name(values)
    end_row_total_and_salary = start_row + len(users) - 1

    # Определение начальной строки для вставки данных
    current_row = start_row

    # Получение всех дат в виде чанков
    date_chunks = create_date_list()

    # Определение текущей даты
    today = datetime.datetime.now()
    current_month = today.strftime("%B")  # Например, "May"
    current_day = today.strftime("%d").lstrip("0")  # Удаление ведущих нулей, например, "8" вместо "08"

    # Поиск текущего чанка
    current_chunk = next((chunk for chunk in date_chunks if any(day == current_day and month == current_month for month, _, day in chunk)), None)

    if not current_chunk:
        return "No chunk found for today's date."
    names_in_last_chunk = get_names_in_last_chunk(values)
    for user_info in users:
        user_name, per_hour = user_info[0], float(user_info[1])
        if user_name not in names_in_last_chunk:
            print(f'Имя {user_name} не найдено в списке текущего чанка -> {names_in_last_chunk}'
                  f'Производим добавление')
            total_formula = f'=SUM(C{current_row}:{chr(66 + len(current_chunk))}{current_row})'
            salary_formula = f'=A{current_row}*{chr(66 + len(current_chunk) + 1)}{current_row}'
            row_data = [[per_hour, user_name] + [' '] * len(current_chunk) + [total_formula, salary_formula]]

            # Добавление данных пользователя в лист
            range_name = f"{sheet_name}!A{current_row}"
            body = {
                'values': row_data
            }
            service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id, range=range_name, body=body,
                valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS'
            ).execute()

            current_row += 1  # Переход к следующей строке для нового пользователя
        else:
            print(f'Имя {user_name} уже есть в таблице, добавлять не нужно!')

    # Добавление строки с формулами сумм Total и Salary
    total_formula_range = f"{sheet_name}!{chr(67 + len(current_chunk))}{end_row_total_and_salary + 1}"
    end_sum_total = f"{chr(67 + len(current_chunk))}{end_row_total_and_salary + 1}"
    print(f'End Total Formula Range: {end_sum_total}')
    salary_formula_range = f"{sheet_name}!{chr(68 + len(current_chunk))}{end_row_total_and_salary + 1}"
    end_sum_salary = f"{chr(68 + len(current_chunk))}{end_row_total_and_salary + 1}"
    print(f'End Salary Formula Range: {end_sum_salary}')

    total_formula_body = {'values': [[f'=SUM({chr(67 + len(current_chunk))}{start_row_total_and_salary}:{chr(67 + len(current_chunk))}{end_row_total_and_salary})']]}
    print(f'Total Formula Body: {total_formula_body}')
    salary_formula_body = {'values': [[f'=SUM({chr(68 + len(current_chunk))}{start_row_total_and_salary}:{chr(68 + len(current_chunk))}{end_row_total_and_salary})']]}
    print(f'Salary Formula Body: {salary_formula_body}')

    # Обновление ячеек с формулами
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=total_formula_range, body=total_formula_body,
        valueInputOption='USER_ENTERED'
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=salary_formula_range, body=salary_formula_body,
        valueInputOption='USER_ENTERED'
    ).execute()

    return f"User names and per hour rates appended to sheet '{sheet_name}' starting from row {start_row}"
