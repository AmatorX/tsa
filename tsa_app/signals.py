from django.db.models.signals import m2m_changed, post_save, pre_save, post_delete
from django.dispatch import receiver
import logging

from logging_config import setup_logging
from tools.create_headers_tools_table import create_headers_tools_table
from tools.remove_tool_from_sheet import remove_tool_from_sheet
from tools.add_tool_to_sheet import add_tool_to_sheet

from tools.update_tool_user import update_tool_user, update_building_in_users_rows
from utils.update_materials_in_results import update_materials_in_sheets
from u_tests.get_start_end_indexes_rows import number_of_rows_in_table
from .models import BuildObject, Worker, Tool, ToolsSheet
from utils.create_object_tables import create_google_sheets
from utils.add_users_to_work_time import append_user_to_work_time
from utils.add_users_to_results_photos import add_users_to_results_photos
from common.get_service import get_service


setup_logging()
logger = logging.getLogger(__name__)


# Используем m2m2_changed что бы отслеживать изменения связывания объекта с материалами многие ко многим
# Так как объект создается только при добавлении материалов
@receiver(m2m_changed, sender=BuildObject.material.through)
def materials_changed(sender, instance, action, reverse, model, pk_set, **kwargs):

    if action == "post_add":  # Проверяем, что материалы были добавлены к объекту строительства
        sh_url = instance.sh_url
        service = get_service()

        # Проверяем наличие листа "Work Time"
        sheet_metadata = service.spreadsheets().get(spreadsheetId=sh_url.split("/")[5]).execute()
        sheets = sheet_metadata.get('sheets', '')
        sheet_names = [sheet['properties']['title'] for sheet in sheets]

        name = instance.name
        total_budget = instance.total_budget
        sh_url = instance.sh_url
        # materials_info = [(material.name, float(material.price)) for material in model.objects.filter(pk__in=pk_set)]

        # Получаем все материалы, уже привязанные к этому объекту
        materials_info = [(material.name, float(material.price)) for material in instance.material.all()]

        if "Work Time" not in sheet_names:
            # Если листа "Work Time" нет, значит объект создается впервые
            logger.info(f"Название объекта строительства: {name}")
            logger.info(f"Добавленные материалы: {materials_info}")
            logger.info(f"Общий бюджет: {total_budget}")
            logger.info(f"URL: {sh_url}")
            create_google_sheets(sh_url, materials_info, instance)
            logger.info(f'Google sheets created')

        else:
            # Лист "Work Time" существует, значит происходит обновление материалов существующего объекта
            logger.info(f'Материалы {materials_info} изменили связь с объектом {name}')
            update_materials_in_sheets(sh_url, materials_info)
            logger.info(f'Добавить данные в Object KPI')

    elif action == "post_remove":
        # Материалы были удалены из BuildObject
        logger.info(f"Материалы удалены из объекта строительства '{instance.name}'.")
        # Здесь можно добавить дополнительную логику

    elif action == "post_clear":
        # Все материалы были удалены из BuildObject
        logger.info(f"Все материалы были удалены из объекта строительства '{instance.name}'.")
        # Здесь можно добавить дополнительную логику


# Отслеживаем создание воркера и изменение связи к объекту строительства
@receiver(post_save, sender=Worker)
def track_worker_changes(sender, instance, created, **kwargs):

    if created:
        # Объект был только что создан
        logger.debug(f"Работник {instance.name} создан впервые.")
        # Вызов функции для обработки создания нового объекта Worker и добавление его в work time таблицу
        if instance.build_obj_id:
            logger.debug(f'Данные для append_user_to_work_time {instance.build_obj.sh_url, [(instance.name, instance.salary)]}')
            append_user_to_work_time(
                instance.build_obj.sh_url,
                [(instance.name, instance.salary)])
            add_users_to_results_photos(
                instance.build_obj.sh_url,
                [instance.name])

    else:
        # Объект обновляется, проверяем изменение привязки к объекту строительства
        original_build_obj_id = getattr(instance, '__original_build_obj_id', None)
        current_build_obj_id = instance.build_obj_id

        if original_build_obj_id != current_build_obj_id or not original_build_obj_id:

            # вставляем проверку есть ли пользователь в ворк тайм
            # Привязка к объекту строительства изменилась
            logger.info(f"Работник {instance.name} изменил свою связь с объектом строительства с ID {original_build_obj_id} на ID {current_build_obj_id}")
            # Вызов функции для обработки изменения привязки
            if instance.build_obj:  # Проверяем что build_obj существует
                append_user_to_work_time(
                    instance.build_obj.sh_url,
                    [(instance.name, instance.salary)])
                add_users_to_results_photos(
                    instance.build_obj.sh_url,
                    [instance.name])
                sh_url = ToolsSheet.objects.first().sh_url
                if sh_url:
                    name = instance.name
                    building = instance.build_obj.name
                    # Обновляем записи с таблице инструментов, так как изменился объект у работника
                    update_building_in_users_rows(sh_url, name, building)


######### Обработка записи, удаления и изменения Tools
@receiver(pre_save, sender=Tool)
def tool_pre_save(sender, instance, **kwargs):
    if instance.pk:
        previous_instance = Tool.objects.get(pk=instance.pk)
        instance._previous_tools_sheet = previous_instance.tools_sheet
    else:
        instance._previous_tools_sheet = None


@receiver(post_save, sender=Tool)
def tool_post_save(sender, instance, **kwargs):
    if instance.tools_sheet and (instance._previous_tools_sheet != instance.tools_sheet):
        # Assuming tools_sheet.sh_url is the URL to the Google Sheet
        add_tool_to_sheet(instance.tools_sheet.sh_url, instance)


@receiver(post_delete, sender=Tool)
def tool_post_delete(sender, instance, **kwargs):
    if instance.tools_sheet:
        # Удаление записи из Google Sheets
        remove_tool_from_sheet(instance.tools_sheet.sh_url, instance)


@receiver(pre_save, sender=Tool)
def tool_pre_save(sender, instance, **kwargs):
    if instance.pk:
        previous_instance = Tool.objects.get(pk=instance.pk)
        instance._previous_assigned_to = previous_instance.assigned_to
    else:
        instance._previous_assigned_to = None


@receiver(post_save, sender=Tool)
def tool_post_save(sender, instance, **kwargs):
    previous_assigned_to = instance._previous_assigned_to
    current_assigned_to = instance.assigned_to

    if previous_assigned_to != current_assigned_to:
        if current_assigned_to is not None:
            # assigned_to изменилось поле с None на имя пользователя или с одного имени пользователя на другое
            update_tool_user(instance.tools_sheet.sh_url, instance)
        elif current_assigned_to is None:
            # assigned_to изменилось с имени пользователя на None
            update_tool_user(instance.tools_sheet.sh_url, instance)


@receiver(post_save, sender=ToolsSheet)
def tools_sheet_post_save(sender, instance, created, **kwargs):
    if created:
        # Вызов функции для создания заголовков только при создании нового объекта
        create_headers_tools_table(instance)
