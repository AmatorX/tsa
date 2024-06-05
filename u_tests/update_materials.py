import datetime
from service import get_service


def update_non_empty_rows_with_materials(service, spreadsheet_id, sheet_title, materials_info, days_in_current_month):
    row_index = 3  # Начинаем с 3-й строки
    requests = []

    while True:
        # Формируем диапазон для проверки
        range_name = f"{sheet_title}!A{row_index}:A{row_index+1}"

        # Делаем запрос к API для получения данных из указанного диапазона
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])

        # Если строки не пустые, формируем запросы на обновление данных для каждого материала
        if values and any(values):
            for i, (material_name, material_price) in enumerate(materials_info, start=1):
                # Здесь формируем запросы на обновление данных для каждого материала
                # Например, обновляем значения в следующей свободной колонке
                requests.append({
                    "updateCells": {
                        "rows": [
                            {"values": [{"userEnteredValue": {"stringValue": str(material_price)}}]},
                            {"values": [{"userEnteredValue": {"stringValue": material_name}}]}
                        ],
                        "fields": "userEnteredValue",
                        "start": {
                            "sheetId": get_sheet_id(service, spreadsheet_id, sheet_title),
                            "rowIndex": row_index-1,
                            "columnIndex": i
                        }
                    }
                })

        # Смещаемся вниз на количество дней в месяце + 12
        row_index += days_in_current_month + 12

        # Прекращаем цикл, если достигли конца проверяемой области
        if row_index > 40:
            break

    # Если есть запросы на обновление, выполняем их
    if requests:
        body = {"requests": requests}
        response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        return response

    return None


def get_sheet_id(service, spreadsheet_id, sheet_name):
    # Получение информации о всех листах в таблице
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск идентификатора листа по имени
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']


# Пример вызова функции
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1JFbB5TucFZn-xl4NtTVpcYAvqP89DZmOy5xxB18rBsI/edit#gid=299250498'
materials = [("Дерево", 100), ("Кирпич", 50)]
update_non_empty_rows_with_materials(spreadsheet_url, materials)