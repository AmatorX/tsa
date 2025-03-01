import datetime
from time import sleep

from celery import shared_task

from django.core.exceptions import ObjectDoesNotExist

from common.get_service import get_service
from kpi_utils.check_kpi_tables_exist import if_object_kpi_tables_exist
from utils.add_users_to_results_photos import add_users_to_results_photos
from utils.add_users_to_work_time import append_user_to_work_time
from utils.create_results_photos import create_monthly_sheets
from utils.create_tables_work_time import create_work_time_tables
from common.days_list import create_date_list
from .models import Worker, BuildObject
from kpi_utils.create_kpi_data import update_data_users_kpi, earned_today_for_installing_materials
from kpi_utils.update_build_kpi_values import update_build_kpi
from kpi_utils.expenses_materials_in_objects import update_materials_on_objects
import logging


logger = logging.getLogger(__name__)


def get_sh_urls():
    sh_urls = BuildObject.objects.all().values_list('sh_url', flat=True)
    logger.info(f'Список sh_urls:{list(sh_urls)}')
    return list(sh_urls)


def get_build_object_by_sh_url(sh_url):
    try:
        build_object = BuildObject.objects.get(sh_url=sh_url)
        return build_object
    except ObjectDoesNotExist:
        logger.error(f'Объект строительства с sh_url {sh_url} не найден.')
        return None


def get_sheets_from_google_sheet(sh_url):
    service = get_service()
    # Извлечение spreadsheet_id из URL
    logger.info(f'SH_URL: {sh_url}')
    spreadsheet_id = sh_url.split("/")[5]
    logger.info(f'spreadsheetId: {spreadsheet_id}')
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    sheet_names = [sheet.get("properties", {}).get("title", "Unnamed Sheet") for sheet in sheets]
    logger.info(f'Список имен листов: {sheet_names}')
    return sheet_names


def get_materials_info(sh_url):
    build_object = get_build_object_by_sh_url(sh_url)
    materials_info = [(material.name, float(material.price)) for material in build_object.material.all()]
    return materials_info


def check_current_month_in_sheets(sheet_names):
    print('Старт функции check_current_month_in_sheets')
    # Получение текущего месяца
    current_month = datetime.datetime.now().strftime("%B")
    # Проверка наличия текущего месяца в именах листов
    for sheet_name in sheet_names:
        if current_month.lower() in sheet_name.lower():
            return True
    return False


def filter_users_names(data):
    """
    Преобразует список, оставляя только имена пользователей и отбрасывая зарплаты.

    Аргументы:
        data (list): Список списков, где каждый элемент содержит `sh_url` и список пользователей с их именами и зарплатами.
                     Пример входного значения:
                     [
                         ['https://example.com/sheet1', [['John Doe', 50000.0], ['Jane Smith', 60000.0]]],
                         ['https://example.com/sheet2', [['Alice Brown', 55000.0]]]
                     ]

    Возвращаемое значение:
        list: Список списков, где каждый элемент содержит `sh_url` и список только имен пользователей.
              Пример возвращаемого значения:
              [
                  ['https://example.com/sheet1', ['John Doe', 'Jane Smith']],
                  ['https://example.com/sheet2', ['Alice Brown']]
              ]
    """
    result = []
    for sh_url, users in data:
        # Извлекаем только имена пользователей
        user_names = [user[0] for user in users]
        # Добавляем результат в новый список
        result.append([sh_url, user_names])
    return result


@shared_task
def check_and_create_tables_results_and_photos():
    service = get_service()
    logger.info('Старт проверки необходимости создания таблиц results и photos')
    sh_urls = get_sh_urls()
    for sh_url in sh_urls:
        sheet_names = get_sheets_from_google_sheet(sh_url)
        if not check_current_month_in_sheets(sheet_names):
            logger.info(f'Текущего месяца нет в листах, создает results и photos для url: {sh_url}')
            materials_info = get_materials_info(sh_url)
            create_monthly_sheets(service, sh_url, materials_info)
            logger.info(f'Листы results и photos созданы для таблицы с URL: {sh_url}')
            build_urls_and_user_names = get_objects_and_related_users()
            data = filter_users_names(build_urls_and_user_names)
            for item in data:
                add_users_to_results_photos(item[0], item[1])
                logger.info(f'Добавляем пользователя {item[1]}')
                sleep(3)
        else:
            logger.info(f'Текущий месяц уже есть в таблицах. Проверено для url: {sh_url}')


###############################################################

@shared_task
def check_and_create_table_work_time():
    #TODO Сделать инициализацию проверки на необходимость создания всех таблиц для Object KPI и USers KPI здесь
    logger.info('Старт проверки необходимости создания новой таблицы чанка')
    date_list = create_date_list()
    current_chunk = get_current_chunk(date_list)
    if current_chunk:
        start_date_str = f"{current_chunk[0][0]} {current_chunk[0][2]} {datetime.datetime.now().year}"
        print(f'start_date_str -> {start_date_str}')

        start_date = datetime.datetime.strptime(start_date_str, "%B %d %Y")
##################
        # if datetime.datetime.now() == start_date:
        if datetime.datetime.now().date() == start_date.date():
        # if start_date: # раскоментировать для принудительного создания табл work_time
            result = get_objects_and_related_users()
            logger.info("Чанк закончился, создаем новую таблицу и добавляем пользователей.")
            # Обработка каждой пары (spreadsheet_url, users)
            for spreadsheet_url, users in result:
                logging.info(f"Обработка таблицы для URL: {spreadsheet_url}")

                # Создание таблицы для текущего URL
                create_work_time_tables(spreadsheet_url)

                # Добавление пользователей к таблице для текущего URL
                append_user_to_work_time(spreadsheet_url, users)
        else:
            logger.info("Текущий чанк еще не завершен.")
    else:
        logger.error("Текущий чанк не найден.")


def get_current_chunk(date_list):
    today = datetime.datetime.now()
    for chunk in date_list:
        start_date_str = f"{chunk[0][0]} {chunk[0][2]} {today.year}"
        end_date_str = f"{chunk[-1][0]} {chunk[-1][2]} {today.year}"
        start_date = datetime.datetime.strptime(start_date_str, "%B %d %Y")
        end_date = datetime.datetime.strptime(end_date_str, "%B %d %Y")

        if start_date <= today <= end_date:
            return chunk
    return None


def get_objects_and_related_users():
    """
        Возвращает список, связывающий URL объектов строительства (sh_url) с пользователями, работающими над этими объектами, и их зарплатами.
        Возвращаемое значение:
            list: Список, где каждый элемент представляет собой список, содержащий `sh_url` и связанный с ним список пользователей.
                  Пример возвращаемого значения:
                  [
                      ['https://example.com/sheet1', [['John Doe', 50000.0], ['Jane Smith', 60000.0]]],
                      ['https://example.com/sheet2', [['Alice Brown', 55000.0]]]
                  ]
        """
    build_objects_to_users = {}

    # Получение информации о пользователях
    workers = Worker.objects.all()
    for worker in workers:
        if worker.build_obj and worker.build_obj.sh_url:  # Проверяем, есть ли связанный BuildObject и его sh_url
            # Добавляем пользователя в список, связанный с sh_url
            if worker.build_obj.sh_url not in build_objects_to_users:
                build_objects_to_users[worker.build_obj.sh_url] = [[worker.name, float(worker.salary)]]
            else:
                build_objects_to_users[worker.build_obj.sh_url].append([worker.name, float(worker.salary)])

    # Преобразование словаря в список списков
    result = [[sh_url, users] for sh_url, users in build_objects_to_users.items()]
    return result


# @shared_task
# def update_kpi():
#     """
#     Функция обновления значений KPI для users и building
#     Запускается через планироващик
#     """
#     result = get_objects_and_related_users()
#     logger.info('Старт рассчета KPI из планировщика')
#     kpi_results = update_data_users_kpi(result)
#     update_build_kpi(kpi_results)
#     logger.info('Планировщик отработал успешно, данные обновлены')
#     return result



@shared_task
def update_kpi_task():
    update_kpi()


def update_kpi():
    result = get_objects_and_related_users()
    logger.info('Старт рассчета KPI')
    # kpi_results = update_data_users_kpi(result)
    kpi_results = update_data_users_kpi()
    build_total_sum_earned_from_installed = earned_today_for_installing_materials()
    # update_build_kpi(kpi_results)
    update_build_kpi(build_total_sum_earned_from_installed)
    logger.info('KPI успешно обновлены')
    return result


@shared_task
def update_objects_with_materials_task():
    update_objects_with_materials()


def update_objects_with_materials():
    building_objects = BuildObject.objects.all()
    building_objects_dict = {}

    for building_object in building_objects:
        materials = building_object.material.all()
        if materials:  # Проверка наличия связанных объектов материалов
            material_names = [material.name for material in materials]
            building_objects_dict[building_object.sh_url] = material_names

    update_materials_on_objects(building_objects_dict)

    return building_objects_dict


@shared_task
def check_object_kpi_tables_exist():
    building_objects = BuildObject.objects.all()
    objects_dict = {}

    for building_object in building_objects:
        materials = building_object.material.all()
        if materials:  # Проверка наличия связанных объектов материалов
            material_names = [material.name for material in materials]
            objects_dict[building_object.sh_url] = material_names
    if_object_kpi_tables_exist(objects_dict)
    return objects_dict
