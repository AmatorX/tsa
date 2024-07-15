from django.db.models.signals import m2m_changed, post_save
from django.dispatch import receiver

from utils.update_materials_in_results import update_materials_in_sheets
from .models import BuildObject, Worker
from utils.create_object_tables import create_google_sheets
from utils.add_users_to_work_time import append_user_to_work_time
from utils.add_users_to_results_photos import add_users_to_results_photos
from common.get_service import get_service


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
            print(f"Название объекта строительства: {name}")
            print(f"Добавленные материалы: {materials_info}")
            print(f"Общий бюджет: {total_budget}")
            print(f"URL: {sh_url}")
            create_google_sheets(sh_url, materials_info, instance)
            print(f'Google sheets created')

        else:
            # Лист "Work Time" существует, значит происходит обновление материалов существующего объекта
            print(f'Материалы {materials_info} изменили связь с объектом {name}')
            update_materials_in_sheets(sh_url, materials_info)
            print(f'Добавить данные в Object KPI')
        
    elif action == "post_remove":
        # Материалы были удалены из BuildObject
        print(f"Материалы удалены из объекта строительства '{instance.name}'.")
        # Здесь можно добавить дополнительную логику

    elif action == "post_clear":
        # Все материалы были удалены из BuildObject
        print(f"Все материалы были удалены из объекта строительства '{instance.name}'.")
        # Здесь можно добавить дополнительную логику


# Отслеживаем создание воркера и изменение связи к объекту строительства
@receiver(post_save, sender=Worker)
def track_worker_changes(sender, instance, created, **kwargs):

    if created:
        # Объект был только что создан
        print(f"--------------------------->Работник {instance.name} создан впервые.")
        # Вызов функции для обработки создания нового объекта Worker и добавление его в work time таблицу
        if instance.build_obj_id:
            print('--------------------------------->Сработал сигнал created')
            print(f'Данные для append_user_to_work_time -------------------> {instance.build_obj.sh_url, [(instance.name, instance.salary)]}')
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
            # Привязка к объекту строительства изменилась
            print(f"Работник {instance.name} изменил свою связь с объектом строительства с ID {original_build_obj_id} на ID {current_build_obj_id}")
            # Вызов функции для обработки изменения привязки
            if instance.build_obj:  # Проверяем что build_obj существует
                append_user_to_work_time(
                    instance.build_obj.sh_url,
                    [(instance.name, instance.salary)])
                add_users_to_results_photos(
                    instance.build_obj.sh_url,
                    [instance.name])


# @receiver(m2m_changed, sender=BuildObject.material.through)
# def track_build_object_material_changes(sender, instance, action, pk_set, **kwargs):
#     print('track_build_object_material_changes')
#     if action == "post_add":
#         # Материалы были добавлены к BuildObject
#         print(f"Материалы добавлены к объекту строительства '{instance.name}'.")
#         # Получение объектов материалов по их первичным ключам
#         materials = Material.objects.filter(pk__in=pk_set)

#         # Извлечение name и price из каждого объекта материала
#         materials_info = [(material.name, material.price) for material in materials]

#     elif action == "post_remove":
#         # Материалы были удалены из BuildObject
#         print(f"Материалы удалены из объекта строительства '{instance.name}'.")
#         # Здесь можно добавить дополнительную логику

#     elif action == "post_clear":
#         # Все материалы были удалены из BuildObject
#         print(f"Все материалы были удалены из объекта строительства '{instance.name}'.")
#         # Здесь можно добавить дополнительную логику