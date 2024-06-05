def get_sheet_id(service, spreadsheet_url, sheet_name):
    """
    Функция получения sheet_id
    """
    spreadsheet_id = spreadsheet_url.split("/")[5]
    # Получение информации о всех листах в таблице
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')

    # Поиск идентификатора листа по имени
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']

    raise ValueError(f"Sheet with name '{sheet_name}' not found.")