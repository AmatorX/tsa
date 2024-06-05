import os
import datetime
from collections import defaultdict

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz

from common.days_list import create_date_list
from utils.add_users_to_work_time import append_user_to_work_time
from utils.create_tables_work_time import create_work_time_tables
from .models import Worker, BuildObject
from kpi_utils.create_kpi_data import update_data_users_kpi
from kpi_utils.update_build_kpi_values import update_build_kpi
from kpi_utils.expenses_materials_in_objects import update_materials_on_objects


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


def get_objects_and_related_users():
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


def check_and_create_tables():
    print('Старт проверки необходимости создания новой таблицы чанка')
    date_list = create_date_list()
    current_chunk = get_current_chunk(date_list)
    if current_chunk:
        end_date_str = f"{current_chunk[-1][0]} {current_chunk[-1][2]} {datetime.datetime.now().year}"
        end_date = datetime.datetime.strptime(end_date_str, "%B %d %Y")

        if datetime.datetime.now() > end_date:
            result = get_objects_and_related_users()
            print("Чанк закончился, создаем новую таблицу и добавляем пользователей.")
            # Обработка каждой пары (spreadsheet_url, users)
            for spreadsheet_url, users in result:
                print(f"Обработка таблицы для URL: {spreadsheet_url}")

                # Создание таблицы для текущего URL
                create_work_time_tables(spreadsheet_url)

                # Добавление пользователей к таблице для текущего URL
                append_user_to_work_time(spreadsheet_url, users)
        else:
            print("Текущий чанк еще не завершен.")
    else:
        print("Текущий чанк не найден.")


def update_kpi():
    """
    Функция обновления значений KPI для users и building
    Запускается через планироващик
    """
    # Словарь для хранения связи между sh_url и списком пользователей
    # build_objects_to_users = {}
    #
    # # Получение информации о пользователях
    # workers = Worker.objects.all()
    # for worker in workers:
    #     if worker.build_obj and worker.build_obj.sh_url:  # Проверяем, есть ли связанный BuildObject и его sh_url
    #         # Добавляем пользователя в список, связанный с sh_url
    #         if worker.build_obj.sh_url not in build_objects_to_users:
    #             build_objects_to_users[worker.build_obj.sh_url] = [[worker.name, float(worker.salary)]]
    #         else:
    #             build_objects_to_users[worker.build_obj.sh_url].append([worker.name, float(worker.salary)])
    #
    # # Преобразование словаря в список списков
    # result = [[sh_url, users] for sh_url, users in build_objects_to_users.items()]

    result = get_objects_and_related_users()
    print('Старт рассчета KPI из планировщика')
    kpi_results = update_data_users_kpi(result)
    update_build_kpi(kpi_results)
    print('Планировщик отработал успешно, данные обновлены')
    return result


# Настройка планировщика
scheduler = BackgroundScheduler()

# Настройка временной зоны
calgary_tz = pytz.timezone('America/Edmonton')  # Калгари находится в часовом поясе America/Edmonton

# Добавление задачи в планировщик. Задача будет выполняться каждый день в 19:00 по времени Калгари
scheduler.add_job(update_kpi, trigger=CronTrigger(hour=8, minute=36, second=1, timezone=calgary_tz))
scheduler.add_job(check_and_create_tables, trigger=CronTrigger(hour=0, minute=1, second=0, timezone=calgary_tz))
scheduler.add_job(update_objects_with_materials, trigger=CronTrigger(hour=12, minute=52, second=0, timezone=calgary_tz))


def start_scheduler():
    if os.environ.get('RUN_MAIN') or os.environ.get('WERKZEUG_RUN_MAIN'):
        scheduler.start()
