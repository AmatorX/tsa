from common.service import service


service = service.get_service()
def get_last_non_empty_row(service, spreadsheet_url, sheet_name):
    """
    Функция находит последнюю записанную ячейку в листе, возвращает индекс этой ячейки.
    Так как Google API оптимизирует запросы, то выбрав ячейки диапазоном вида "{sheet_name}!A1:A",
    API вернет ячейки от 1 до последней заполненной.
    """

    spreadsheet_id = spreadsheet_url.split("/")[5]
    range_name = f"{sheet_name}!A1:A"  # Запрашиваем данные только из первого столбца

    # Получаем данные столбца
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    # Возвращаем индекс последней заполненной строки
    last_row_index = len(values) if values else 0  # Если нет данных, возвращаем 0

    print(f"Last non-empty row in column A is {last_row_index}")
    return last_row_index



# spreadsheet_url = 'https://docs.google.com/spreadsheets/d/13EOrnS-F8IkBYGd4GVg3vY02E-l56mLin2BHJ6cv6Is/edit#gid=0'
# sheet_name = 'Лист1'
#
# get_last_non_empty_row(service, spreadsheet_url, sheet_name)