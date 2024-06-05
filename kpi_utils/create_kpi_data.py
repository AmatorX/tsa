from datetime import datetime
from kpi_utils.update_users_kpi_values import update_kpi_values
from common.service import service

# Здесь подготавливаются данные для записи в таблицах USers KPI

service = service.get_service()


def get_sheet_id(service, spreadsheet_id, sheet_name):
    # Получение информации о всех листах в таблице
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск идентификатора листа по имени
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']

    raise ValueError(f"Sheet with name '{sheet_name}' not found.")


def get_current_month_and_day():
    now = datetime.now()
    current_year = now.year
    current_month = now.strftime("%B")  # Получаем название текущего месяца
    current_day = now.day  # Получаем номер текущего дня
    return current_month, current_day, current_year


def find_sheet_id_by_name(spreadsheet_url, service, sheet_name='Work Time'):
    '''
    Получение ID листа 'Work Time'
    '''
    # Извлечение spreadsheet_id из URL
    spreadsheet_id = spreadsheet_url.split("/")[5]


    # Получение списка всех листов в таблице
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск ID листа по названию
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']

    # Если лист с таким названием не найден, возвращаем None
    return None


def find_cell_by_date(spreadsheet_url):

    # service = get_service()
    # Извлечение spreadsheet_id из URL
    spreadsheet_id = spreadsheet_url.split("/")[5]

    date = get_current_month_and_day()
    month_day = date[0] + ' ' + str(date[1])
    print(f'Текущая дата: {month_day}')

    # Определение диапазона для поиска
    range_name = 'Work Time!A:Z'

    # Получение данных из листа
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    # Поиск ячейки с заданным содержимым
    for row_index, row in enumerate(values):
        for col_index, value in enumerate(row):
            if month_day == value:
                # print(f"Найдено совпадение в строке {row_index + 1}, столбце {col_index + 1}")
                return spreadsheet_id, row_index + 1, col_index + 1

    return None


def calculate_salary_by_names(spreadsheet_url, user_info_list):
    # service = get_service()
    spreadsheet_id = spreadsheet_url.split("/")[5]
    salaries = []

    # Находим индекс ячейки с текущей датой
    _, date_row_index, date_col_index = find_cell_by_date(spreadsheet_url)

    # Определение диапазона для поиска имен пользователей начиная со следующей строки после даты
    range_end = date_row_index + len(user_info_list) + 1  # +1 для учета дополнительной строки после списка пользователей
    range_name = f'Work Time!B{date_row_index + 1}:B{range_end}'


    # Получение данных из листа
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    for user in user_info_list:
        user_name, user_salary = user
        user_found = False

        # Поиск имени пользователя в списке значений
        for row_index, row in enumerate(values):

            if row and row[0] == user_name:

                user_found = True
                # Получение часов из соответствующего столбца даты для найденного пользователя
                hours_range = f'Work Time!{chr(65 + date_col_index - 1)}{date_row_index + row_index + 1}'
                hours_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=hours_range).execute()
                hours = hours_result.get('values', [[0]])[0][0]

                # Вычисление заработной платы
                calculated_salary = float(hours) * float(user_salary)
                salaries.append([user_name, calculated_salary])
                break  # Прерываем цикл, так как нашли пользователя и рассчитали его зарплату

        if not user_found:
            # Если пользователь не найден, добавляем его с зарплатой 0
            salaries.append([user_name, 0.0])
    return spreadsheet_url, salaries


def get_kpi_for_work_time(sh_and_users):
    work_time_kpi = []
    for data_list in sh_and_users:
        spreadsheet_url = data_list[0]
        user_info_list = data_list[1]
        name_salary_work_time = calculate_salary_by_names(spreadsheet_url, user_info_list)
        work_time_kpi.append(name_salary_work_time)
    print()
    print(f'KPI WORk TIME\n{work_time_kpi}')
    print()
    return work_time_kpi


# Рассчет для KPI по таблицам результатов

def get_results_current_month_table(service, spreadsheet_url):
    current_date = get_current_month_and_day()
    sheet_name = f'results_{current_date[0]}_{current_date[2]}'
    # sheet_id = find_sheet_id_by_name(spreadsheet_url=spreadsheet_url, service=service, sheet_name=sheet_name)
    return sheet_name


def get_indexes_cell_name(service, spreadsheet_url, sheet_name, name):

    spreadsheet_id = spreadsheet_url.split("/")[5]
    # Определение диапазона для поиска в столбце B
    range_name = f"{sheet_name}!B:B"

    # Получение данных из столбца B
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()

    values = result.get('values', [])

    # Поиск ячейки с заданным именем
    for row_index, row in enumerate(values):
        if row and row[0] == name:
            print(f"Найдено совпадение в строке {row_index + 1}, столбце 2, для имени {name}")
            return name, row_index + 1

    # Если имя не найдено, возвращаем None
    return None


def convert_to_float(value):
    """Функция для преобразования строки в число с плавающей точкой"""
    if value:
        # Замена запятых на точки и удаление пробелов
        value = value.replace(',', '.').strip()
        try:
            # Попытка преобразовать строку в число с плавающей точкой
            return float(value)
        except ValueError:
            # В случае неудачи возвращаем 0
            return 0
    else:
        return 0


def calculate_total_material_cost(service, spreadsheet_url, sheet_name, name, start_row):
    # Определение текущего дня
    now = datetime.now()
    current_day = now.day
    spreadsheet_id=spreadsheet_url.split("/")[5]
    # Определение диапазона для material_price
    material_price_range = f"{sheet_name}!B{start_row + 2}:Z{start_row + 2}"
    # Получение данных из листа для material_price
    material_price_result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=material_price_range
    ).execute()
    material_price_values = material_price_result.get('values', [[]])[0]

    # Определение диапазона для salary
    salary_range = f"{sheet_name}!B{start_row + 4 + current_day}:Z{start_row + 4 + current_day}"
    # Получение данных из листа для salary
    salary_result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=salary_range
    ).execute()
    salary_values = salary_result.get('values', [[]])[0]

    # Вычисление суммы произведений
    total_cost = 0
    for material_price, salary in zip(material_price_values, salary_values):
        if material_price:  # Проверка, что ячейка material_price не пустая
            material_price = convert_to_float(material_price)
            salary = convert_to_float(salary)
            total_cost += material_price * salary
        else:
            break  # Прекращаем цикл, если встретили пустую ячейку в material_price
    print(f"Total cost for {name}: {total_cost}")
    return name, total_cost


def get_kpi_for_results(sh_and_users):
    # service = get_service()
    kpi_results = {}
    for data_list in sh_and_users:
        spreadsheet_url = data_list[0]
        user_info_list = data_list[1]
        sheet_name = get_results_current_month_table(service=service, spreadsheet_url=spreadsheet_url)
        user_results = []
        for user_data in user_info_list:
            name = user_data[0]
            name, start_row = get_indexes_cell_name(service=service, spreadsheet_url=spreadsheet_url, sheet_name=sheet_name, name=name)
            name_salary_results = calculate_total_material_cost(
                service=service,
                spreadsheet_url=spreadsheet_url,
                sheet_name=sheet_name,
                name=name,
                start_row=start_row)
            user_results.append(list(name_salary_results))
        kpi_results[spreadsheet_url] = user_results

    # Преобразование словаря в список кортежей
    kpi_results = [(url, data) for url, data in kpi_results.items()]
    print()
    print(f'KPI RESULTS\n{kpi_results}')
    print()
    return kpi_results


def update_data_users_kpi(sh_and_users):
    kpi_work_time = get_kpi_for_work_time(sh_and_users=sh_and_users)
    kpi_results = get_kpi_for_results(sh_and_users=sh_and_users)
    # Преобразование списков в словари для удобства обращения
    dict_work_time = {url: dict(data) for url, data in kpi_work_time}
    dict_results = {url: dict(data) for url, data in kpi_results}

    data_list = []

    # Итерация по первому словарю
    for url, data in dict_work_time.items():
        # Проверка, существует ли соответствующий URL во втором словаре
        if url in dict_results:
            # Список для хранения обновлённых данных по текущему URL
            updated_data = []
            for name, value in data.items():
                # Вычисление разницы значений для соответствующего имени, если оно существует во втором словаре
                if name in dict_results[url]:
                    # diff = value - dict_results[url][name]
                    diff = dict_results[url][name] - value
                    updated_data.append([name, diff])
            # Добавление обновлённых данных в итоговый результат
            data_list.append((url, updated_data))
    print(f'DATA FOR KPI SHEET\n{data_list}')
    update_kpi_values(data_list=data_list)
    return kpi_results

