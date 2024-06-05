from .create_tables_work_time import create_work_time_tables
from .create_results_photos import create_monthly_sheets
from .create_users_kpi_sheet_and_table import create_users_kpi_sheet
from .create_object_kpi_sheet import create_object_kpi_sheet
from common.get_service import get_service


def create_google_sheets(spreadsheet_url, materials, build_object):
    service = get_service()
    create_work_time_tables(spreadsheet_url, sheet_name='Work Time')
    create_users_kpi_sheet(service, spreadsheet_url, sheet_name='Users KPI')
    create_object_kpi_sheet(service, spreadsheet_url, build_object, sheet_name='Object KPI')
    create_monthly_sheets(service, spreadsheet_url, materials)
