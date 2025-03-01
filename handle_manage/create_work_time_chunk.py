import datetime
import logging

from common.days_list import create_date_list
from tsa_app.models import Worker
from utils.add_users_to_work_time import append_user_to_work_time
from utils.create_tables_work_time import create_work_time_tables

logger = logging.getLogger(__name__)


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


def get_object_and_related_users(build_obj):
    """
    Возвращает данные, связывающие URL объекта строительства (sh_url) с пользователями, работающими над этим объектом, и их зарплатами.

    Аргументы:
        build_obj (BuildObject): Объект строительства, для которого необходимо получить данные.

    Возвращаемое значение:
        list: Список, содержащий `sh_url` и список пользователей, работающих над этим объектом.
              Пример возвращаемого значения:
              ['https://example.com/sheet1', [['John Doe', 50000.0], ['Jane Smith', 60000.0]]]
              Если объект не имеет связанных пользователей, возвращается пустой список пользователей.
    """
    if not build_obj or not build_obj.sh_url:  # Проверяем наличие объекта и его sh_url
        return None

    # Получение пользователей, связанных с данным объектом
    workers = Worker.objects.filter(build_obj=build_obj)

    # Формирование списка пользователей с их зарплатами
    users = [[worker.name, float(worker.salary)] for worker in workers]

    # Формирование результата
    result = [build_obj.sh_url, users]
    return result


# def check_and_create_table_work_time(build_obj):
#     logger.info(f'Старт проверки необходимости создания новой таблицы чанка для объекта {build_obj.name}')
#     date_list = create_date_list()
#     current_chunk = get_current_chunk(date_list)
#     if current_chunk:
#         start_date_str = f"{current_chunk[0][0]} {current_chunk[0][2]} {datetime.datetime.now().year}"
#         print(f'start_date_str -> {start_date_str}')
#
#         start_date = datetime.datetime.strptime(start_date_str, "%B %d %Y")
#         if start_date: # раскоментировать для принудительного создания табл work_time
#             result = get_object_and_related_users(build_obj)
#             logger.info("Чанк закончился, создаем новую таблицу и добавляем пользователей.")
#             # Обработка каждой пары (spreadsheet_url, users)
#             for spreadsheet_url, users in result:
#                 logging.info(f"Обработка таблицы для URL: {spreadsheet_url}, объект: {build_obj.name}")
#
#                 # Создание таблицы для текущего URL
#                 create_work_time_tables(spreadsheet_url)
#
#                 # Добавление пользователей к таблице для текущего URL
#                 append_user_to_work_time(spreadsheet_url, users)
#         else:
#             logger.info("Текущий чанк еще не завершен.")
#     else:
#         logger.error("Текущий чанк не найден.")


def check_and_create_table_work_time(build_obj):
    logger.info(f'Старт создания новой таблицы чанка для объекта {build_obj.name}')
    date_list = create_date_list()
    current_chunk = get_current_chunk(date_list)
    if current_chunk:
        start_date_str = f"{current_chunk[0][0]} {current_chunk[0][2]} {datetime.datetime.now().year}"
        logger.info(f'start_date_str -> {start_date_str}')

        start_date = datetime.datetime.strptime(start_date_str, "%B %d %Y")
        if start_date:  # раскоментировать для принудительного создания табл work_time
            result = get_object_and_related_users(build_obj)
            if result:  # Проверяем, что данные существуют
                spreadsheet_url, users = result  # Распаковываем корректно, т.к. результат — один список
                logger.info(f"Обработка таблицы для URL: {spreadsheet_url}, объект: {build_obj.name}")

                # Создание таблицы для текущего URL
                create_work_time_tables(spreadsheet_url)

                # Добавление пользователей к таблице для текущего URL
                append_user_to_work_time(spreadsheet_url, users)
            else:
                logger.warning(f"No related users found for {build_obj.name}")
        else:
            logger.info("Текущий чанк еще не завершен.")
    else:
        logger.error("Текущий чанк не найден.")



