from time import sleep
# from days_list import create_date_list
# from get_service import get_service
from common.days_list import create_date_list
from common.get_service import get_service
import datetime


def get_start_row(spreadsheet_url, sheet_name):
    service = get_service()

    # Извлечение spreadsheet_id из URL
    spreadsheet_id = spreadsheet_url.split("/")[5]

    # Определение диапазона для чтения данных из столбца B
    range_name = f"{sheet_name}!B:B"

    try:
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])

        # Поиск первой пустой ячейки в столбце B
        for i, value in enumerate(values, start=1):
            if not value:  # Если ячейка пуста
                print(f"Empty cell found at row --------------> {i}")
                return i  # Возвращаем номер строки первой пустой ячейки

        # Если пустых ячеек не нашлось, возвращаем номер строки после последней непустой ячейки
        return len(values) + 1

    except Exception as e:
        print(f"An error occurred: {e}")
        return None  # В случае ошибки возвращаем None


def append_user_to_work_time(spreadsheet_url, users, sheet_name='Work Time'):

    if users is None:
        pass
        # TODO: Добавить проверку переданы ли пользователи или нет, если нет, азять из  базы

    print('Функция добавления пользователей в таблицы work time запустилась')
    service = get_service()
    start_row = get_start_row(spreadsheet_url, sheet_name)

    # Извлечение spreadsheet_id из URL
    spreadsheet_id = spreadsheet_url.split("/")[5]

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

    # Задаем стартовую строку для добавления формул Total и Salary
    start_row_total_and_salary = start_row
    end_row_total_and_salary = start_row + len(users) - 1

    for user_info in users:
        user_name, per_hour = user_info[0], float(user_info[1])
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

    # Добавление строки с формулами сумм Total и Salary
    total_formula_range = f"{sheet_name}!{chr(67 + len(current_chunk))}{end_row_total_and_salary + 1}"
    salary_formula_range = f"{sheet_name}!{chr(68 + len(current_chunk))}{end_row_total_and_salary + 1}"

    total_formula_body = {'values': [[f'=SUM({chr(67 + len(current_chunk))}{start_row_total_and_salary}:{chr(67 + len(current_chunk))}{end_row_total_and_salary})']]}
    salary_formula_body = {'values': [[f'=SUM({chr(68 + len(current_chunk))}{start_row_total_and_salary}:{chr(68 + len(current_chunk))}{end_row_total_and_salary})']]}

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
