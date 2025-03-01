import logging
from datetime import datetime
from time import sleep

from common.service import service
from kpi_utils.expenses_materials_in_objects import create_materials_table
# from kpi_utils.expenses_materials_in_objects import create_materials_table
from utils.create_object_kpi_sheet import create_monthly_kpi_table
from common.last_non_empty_row import get_last_non_empty_row


service = service.get_service()

logger = logging.getLogger(__name__)


def is_first_of_the_month():
    """
    Функция проверки числа месяца.
    Вернет True если 1-е число, иначе вернет False
    """
    today = datetime.now()
    return today.day == 1


def get_remaining_budget(service, spreadsheet_url, sheet_name):
    spreadsheet_id = spreadsheet_url.split("/")[5]
    range_name = f"{sheet_name}!E1:E"  # Запрашиваем данные только из первого столбца
    # Получаем данные столбца
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    sleep(2)
    values = result.get('values', [])
    # Определяем номер последней заполненной строки
    last_row = len(values) - 1 if values else 0
    if last_row >= 0:
        budget = values[last_row][0]  # Получаем значение из последней непустой строки
    else:
        budget = None  # Если в столбце нет значений, возвращаем None
    return budget


def get_current_day_row(spreadsheet_url, sheet_name):
    # Получаем текущий месяц и день
    now = datetime.now()
    current_month = now.strftime("%B")  # Название месяца на английском
    current_day = now.day  # Число месяца

    spreadsheet_id = spreadsheet_url.split("/")[5]
    range_name = f"{sheet_name}!A:A"

    # Получаем данные из столбца A
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    sleep(2)
    values = result.get('values', [])

    # Находим строку с текущим месяцем
    month_row = None
    for index, value in enumerate(values):
        if value and current_month in value[0]:
            month_row = index + 1  # Google Sheets использует индексацию, начиная с 1
            break
    if month_row is None:
        return None  # Месяц не найден

    # Проверяем диапазон дат с найденной строки по конец столбца
    day_range_name = f"{sheet_name}!A{month_row}:A"
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=day_range_name).execute()
    sleep(2)
    day_values = result.get('values', [])

    # Находим строку с текущим числом
    for index, value in enumerate(day_values):
        if value and str(current_day) == value[0].strip():
            return month_row + index  # Возвращаем индекс строки с текущим днем
    return None  # День не найден


def update_kpi_value(spreadsheet_url, total_sum, current_day_row):
    logger.info(f"Функция обновления данных update_kpi_value запущена")

    spreadsheet_id = spreadsheet_url.split("/")[5]
    sheet_name = 'Object KPI'
    column_letter = 'B'  # Столбец B для записи total_sum

    # Формируем range для обновления
    range_name = f"{sheet_name}!{column_letter}{current_day_row}"

    # Значения для обновления
    values = [[total_sum]]

    # Формирование тела запроса для обновления ячеек
    body = {
        'values': values,
        'range': range_name,
        'majorDimension': 'ROWS'
    }

    # Выполнение запроса на обновление данных в листе
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
    sleep(6)

    print(f'Updated total sum: {total_sum} in {range_name} on sheet {sheet_name}')
    return result



def update_build_kpi(build_kpi, sheet_name='Object KPI'):
    # !!!!! Сейчас в user_data передаются данные вида
    # [('https://docs.google.com/spreadsheets/d/1Y2Wt6fLdDkjR2QZoEui0_2LtJCihzdTjOZXS4N5QNr4/edit?gid=0#gid=0', [['Alexandr Pinsky', 256.0], ['Farhad', 1122.6]]), ('https://docs.google.com/spreadsheets/d/1JFbB5TucFZn-xl4NtTVpcYAvqP89DZmOy5xxB18rBsI/edit?gid=661746781#gid=661746781', [['Leon', -76.80000000000001]])]
    # где числа это разница заработанного и заплаченного по ставке за часы, необъодимо тут передавать фактически заработанное
    logger.info(f'Функция обновления KPI для таблицы Build KPI запущена. Передан аргумент build_kpi: {build_kpi}')
    # Перебор всех записей в build_kpi
    for spreadsheet_url, user_data in build_kpi:
        total_sum = sum(value for _, value in user_data)
        print(f'Total sum for {spreadsheet_url} is {total_sum}')
        # Получаем строку для текущего дня
        current_day_row = get_current_day_row(spreadsheet_url, sheet_name)
        print(f'Current day row: {current_day_row}')

        # Обновляем данные в KPI
        logger.info(f'Функция обновления KPI для таблицы Build KPI передает данные в update_kpi_value: '
                    f'{spreadsheet_url}, total_sum: {spreadsheet_url}, current_day_row: {current_day_row}')
        update_kpi_value(spreadsheet_url, total_sum, current_day_row)
        sleep(6)  # Пауза между обновлениями для каждой таблицы



# sh_url = 'https://docs.google.com/spreadsheets/d/1gKsdz-icvXkx7km6Dl5RIqcEkGhn9GmFp80BIt3kVco/edit#gid=1862401406'
# sh_url = 'https://docs.google.com/spreadsheets/d/1JFbB5TucFZn-xl4NtTVpcYAvqP89DZmOy5xxB18rBsI/edit#gid=661746781'

# kpi = [('https://docs.google.com/spreadsheets/d/1gKsdz-icvXkx7km6Dl5RIqcEkGhn9GmFp80BIt3kVco/edit#gid=1862401406', [['Nikolay', 491.9], ['Egor', 1370.0], ['Kris', 444.0]]), ('https://docs.google.com/spreadsheets/d/1JFbB5TucFZn-xl4NtTVpcYAvqP89DZmOy5xxB18rBsI/edit#gid=661746781', [['Denis', 531.0], ['Hron', 1065.0], ['Alex', 0]])]
# update_build_kpi(kpi)
