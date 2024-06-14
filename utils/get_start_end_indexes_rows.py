import calendar
from time import sleep

from common.service import service


service = service.get_service()


def get_indexes_for_rows(spreadsheet_url, sheet_name):
    spreadsheet_id = spreadsheet_url.split("/")[5]
    return get_start_end_table(service, spreadsheet_id, sheet_name)


def number_of_rows_in_table(sheet_name):
    # Определяем месяц и год из названия листа
    parts = sheet_name.split('_')
    month_name = parts[1]
    year = int(parts[2])

    # Получаем количество дней в месяце
    month_number = list(calendar.month_name).index(month_name)
    days_in_month = calendar.monthrange(year, month_number)[1]

    # Определяем количество строк в таблице в зависимости от типа листа
    if 'photos' in sheet_name:
        return days_in_month + 1
    elif 'results' in sheet_name:
        return days_in_month + 8
    else:
        return None  # Если тип листа не распознан


def get_start_end_table(service, spreadsheet_id, sheet_name):
    """
    Функция определяет начальную и конечную строки таблицы для копирования,
    так же определяет начальную и конечную позицию для вставки скопированныой таблицы.
    Можно так же применять для определения начала и конца таблиц в листе.
    Например при определении границ ячеек при измененнии материалов в объекте строительства
    """

    # Получаем данные о границах ячеек листа
    result = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=sheet_name,
        fields="sheets(data(rowData(values(userEnteredFormat/borders))))"
    ).execute()
    rows = result['sheets'][0]['data'][0]['rowData']

    last_bordered_row = None
    empty_rows_count = 0

    # Итерируем по строкам с конца, ищем последнюю строку с границей
    for i in range(len(rows) - 1, -1, -1):
        row = rows[i]
        if 'values' in row:
            is_bordered = any(
                'borders' in cell.get('userEnteredFormat', {}) for cell in row['values']
            )

            if is_bordered:
                last_bordered_row = i + 1
                print(f"Последняя обведённая строка: {last_bordered_row}")
                sleep(0.1)
                break
            elif not is_bordered:
                empty_rows_count += 1
                if empty_rows_count > 3:
                    break  # Более трёх пустых строк подряд

    if last_bordered_row is not None:
        # Последняя строка + 3 для start_paste
        start_paste = last_bordered_row + 4
        # Здесь добавьте логику для определения количества строк в последней таблице и рассчитайте end_row
        end_paste = start_paste + number_of_rows_in_table(sheet_name)
        start_copy = 0
        stop_copy = end_paste - start_paste
        return start_copy, stop_copy, start_paste, end_paste
    else:
        return None, None, None, None


# # Пример использования
# spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1JFbB5TucFZn-xl4NtTVpcYAvqP89DZmOy5xxB18rBsI/edit#gid=299250498'
# sheet_name = 'photos_January_2024'

# start_copy, stop_copy, start_paste, end_paste = get_indexes_for_rows(spreadsheet_url, sheet_name)
# print(start_copy, stop_copy, start_paste, end_paste)