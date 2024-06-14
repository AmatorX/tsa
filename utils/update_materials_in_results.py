import datetime
from time import sleep

from common.service import service

service = service.get_service()

def get_names_list():
    year = datetime.datetime.now().year
    all_month_names = ["January", "February", "March", "April", "May", "June",
                       "July", "August", "September", "October", "November", "December"]
    current_month = datetime.datetime.now().month
    return [f'results_{month}_{year}' for month in all_month_names[current_month - 1:]]

def get_index_pairs(spreadsheet_id, sheet_name):
    print(f'Sheet name: {sheet_name}')
    range_name = f"{sheet_name}!A1:A"
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    indices = {'Price': [], 'Material': []}
    for i, val in enumerate(values):
        if val in [['Price'], ['Material']]:
            indices[val[0]].append(i + 1)

    return list(zip(indices['Price'], indices['Material']))

def get_value_pairs(spreadsheet_id, sheet_name, index_pairs):
    if not index_pairs:
        return {}

    price_idx, material_idx = index_pairs[0]
    price_range = f"{sheet_name}!B{price_idx}:Z{price_idx}"
    material_range = f"{sheet_name}!B{material_idx}:Z{material_idx}"

    price_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=price_range).execute()
    material_result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=material_range).execute()

    price_values = price_result.get('values', [[]])[0]
    material_values = material_result.get('values', [[]])[0]

    columns = [chr(i) for i in range(ord('B'), ord('Z') + 1)]
    return {col: [material, price] for col, material, price in zip(columns, material_values, price_values)}


def add_missing_materials(spreadsheet_id, sheet_name, index_pairs, value_pairs, materials_info):
    columns = [chr(i) for i in range(ord('B'), ord('Z') + 1)]
    used_columns = set(value_pairs.keys())

    for material, price in materials_info:
        if any(material in values for values in value_pairs.values()):
            continue  # Skip materials already in the value_pairs

        # Find the first available column that isn't used
        for col in columns:
            if col not in used_columns:
                used_columns.add(col)
                for price_idx, material_idx in index_pairs:
                    material_cell = f"{sheet_name}!{col}{material_idx}"
                    price_cell = f"{sheet_name}!{col}{price_idx}"

                    material_value_result = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=material_cell
                    ).execute()
                    sleep(1)
                    price_value_result = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=price_cell
                    ).execute()
                    sleep(1)

                    material_value = material_value_result.get('values', [[]])[0]
                    price_value = price_value_result.get('values', [[]])[0]

                    if not (material_value and material_value[0].startswith("=")):
                        # print(f"Updating material cell {material_cell} with value {material}")
                        service.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range=material_cell,
                            valueInputOption="RAW",
                            body={"values": [[material]]}
                        ).execute()
                        sleep(1)

                    if not (price_value and price_value[0].startswith("=")):
                        # print(f"Updating price cell {price_cell} with value {price}")
                        service.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range=price_cell,
                            valueInputOption="RAW",
                            body={"values": [[price]]}
                        ).execute()
                        sleep(1)
                break


def update_materials_in_sheets(sh_url, materials_info):
    print(f'URL листа {sh_url} \n материалы {materials_info}')
    spreadsheet_id = sh_url.split("/")[5]
    names_list = get_names_list()
    for sheet_name in names_list:
        index_pairs = get_index_pairs(spreadsheet_id, sheet_name)
        print(f'Index pairs: {index_pairs}\n')
        value_pairs = get_value_pairs(spreadsheet_id, sheet_name, index_pairs)
        print(f'Value pairs: {value_pairs}\n')
        add_missing_materials(spreadsheet_id, sheet_name, index_pairs, value_pairs, materials_info)
        sleep(2)

