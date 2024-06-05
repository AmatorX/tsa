from datetime import datetime
from time import sleep
from utils.create_users_kpi_sheet_and_table import create_users_kpi_table
from common.service import service


service = service.get_service()


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
    now = datetime.now()
    current_day = now.day
    current_month_name = now.strftime("%B")

    print(f'Текущий месяц: {current_month_name}, Текущий день: {current_day}')

    for spreadsheet_url, user_data in data_list:
        if not find_users_kpi_table(service, spreadsheet_url):
            create_users_kpi_table(service, spreadsheet_url)
        spreadsheet_id = spreadsheet_url.split("/")[5]
        sheet_name = "Users KPI"

        # range_name = f"{sheet_name}!A1:A150"
        range_name = f"{sheet_name}!A:A"
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
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
                print(f'Поиск имени {name} в строках {month_row_index}:{month_row_index + 100}')
                if not user_rows:
                    empty_row_index = month_row_index + 1
                else:
                    for name_row_index, name_row in enumerate(user_rows, start=month_row_index + 1):
                        if name_row and name_row[0] == name:
                            name_found = True
                            update_value_range = f"{sheet_name}!{get_column_letter(current_day)}{name_row_index}"
                            update_value_body = {'values': [[value]]}
                            service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=update_value_range, body=update_value_body, valueInputOption='USER_ENTERED').execute()
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

                    update_name_range = f"{sheet_name}!A{empty_row_index}"
                    update_name_body = {'values': [[name]]}
                    service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=update_name_range, body=update_name_body, valueInputOption='USER_ENTERED').execute()

                    update_value_range = f"{sheet_name}!{get_column_letter(current_day)}{empty_row_index}"
                    update_value_body = {'values': [[value]]}
                    service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=update_value_range, body=update_value_body, valueInputOption='USER_ENTERED').execute()

                    # Обновляем переменную values для отражения последних изменений в таблице
                    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
                    values = result.get('values', [])
                    sleep(5)


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