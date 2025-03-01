import logging

from django.contrib import admin, messages

from common.get_service import get_service
from handle_manage.create_work_time_chunk import check_and_create_table_work_time
from handle_manage.users_kpi_sheet1 import run_write_kpi_users_in_sheet1
from utils.add_users_to_results_photos import add_many_users_to_results_photos
from utils.create_results_photos import create_monthly_sheets

from .models import Worker, BuildObject, Material, Tool, ToolsSheet
# from ..utils.add_users_to_work_time import append_user_names_to_tables
from django.contrib.admin import SimpleListFilter

from .tasks import update_kpi, update_objects_with_materials


logger = logging.getLogger(__name__)


class WorkerAdmin(admin.ModelAdmin):
    list_display = ('name', 'tg_id', 'build_obj', 'salary', 'title', 'foreman')
    list_display_links = ('name',)
    list_filter = ('build_obj', 'salary', 'title')
    search_fields = ('name', 'tg_id', 'build_obj')
    list_editable = ('build_obj',)
    fieldsets = (
        (None, {
            'fields': ('name', 'tg_id', 'salary', 'foreman', 'build_obj', 'phone_number', 'email')
        }),
        ('Advanced options', {
            'classes': ('collapse',),
            'fields': ('employment_agreement', 'over_time', 'start_to_work', 'title', 'payroll_eligible', 'payroll', 'resign_agreement', 'benefits_eligible', 'birthday', 'issued', 'expiry', 'address', 'tickets_available'),
        }),
    )
    list_per_page = 10
    # inlines = (WorkerProfileInline,)

    def save_model(self, request, obj, form, change):
        super().save_model(request, obj, form, change)
        if not change:  # Проверяем, что объект был создан, а не обновлен
            pass
            # append_user_names_to_tables(obj)  # Ваша функция для записи данных
            # Отображаем сообщение в админке
            # messages.success(request, 'Данные записаны успешно!')


class MaterialAdmin(admin.ModelAdmin):
    fields = ('name', 'price')
    list_display = ('name', 'price')


class BuildObjectFilter(SimpleListFilter):
    title = 'Build Object'  # Название фильтра в админке
    parameter_name = 'build_obj'  # Параметр фильтра

    def lookups(self, request, model_admin):
        # Возвращаем список опций для фильтра
        build_objs = set(obj.assigned_to.build_obj for obj in model_admin.model.objects.all() if obj.assigned_to and obj.assigned_to.build_obj)
        return [(obj.id, obj.name) for obj in build_objs]

    def queryset(self, request, queryset):
        # Фильтруем queryset в зависимости от выбранного значения
        if self.value():
            return queryset.filter(assigned_to__build_obj__id=self.value())
        return queryset


class ToolAdmin(admin.ModelAdmin):
    fields = ('name', 'tool_id', 'price', 'date_of_issue', 'assigned_to', 'tools_sheet')
    list_display = ('name', 'tool_id', 'price', 'date_of_issue', 'assigned_to', 'get_build_obj_name')
    list_editable = ('assigned_to', 'price', 'date_of_issue',)
    list_filter = (BuildObjectFilter, 'name', 'tool_id', 'date_of_issue', 'assigned_to', )
    list_per_page = 50

    def get_build_obj_name(self, obj):
        return obj.assigned_to.build_obj.name if obj.assigned_to and obj.assigned_to.build_obj else 'N/A'

    get_build_obj_name.short_description = 'Build Object'


class ToolsSheetAdmin(admin.ModelAdmin):
    fields = ('name', 'sh_url')


@admin.action(description="Start updating users KPI data")
def run_update_users_kpi(modeladmin, request, queryset):
    # Вызов функции update_kpi
    update_kpi()

    modeladmin.message_user(request, "The data has been updated")


@admin.action(description="Start updating building KPI data")
def run_update_building_kpi(modeladmin, request, queryset):
    # Вызов функции update_kpi
    update_objects_with_materials()
    modeladmin.message_user(request, "The data has been updated")


@admin.action(description="Create a work time table for the selected object")
def create_work_time_table(modeladmin, request, queryset):
    for build_object in queryset:
        try:
            check_and_create_table_work_time(build_object)
            modeladmin.message_user(
                request,
                f"The working time table for the object {build_object.name} has been created!"
            )
        except Exception as e:
            logger.error(f"Error while creating table for {build_object.name}: {e}")
            modeladmin.message_user(
                request,
                f"Error creating table for {build_object.name}: {e}",
                level=logging.ERROR
            )


@admin.action(description="Create a photos and results table for the selected object")
def create_photos_table_handle(modeladmin, request, queryset):
    for build_object in queryset:
        try:
            service = get_service()
            materials = [(material.name, float(material.price)) for material in build_object.material.all()]
            create_monthly_sheets(service, build_object.sh_url, materials)
            modeladmin.message_user(
                request,
                f"The photos and results table for the object {build_object.name} has been created!"
            )
        except Exception as e:
            logger.error(f"Error while creating table for {build_object.name}: {e}")
            modeladmin.message_user(
                request,
                f"Error creating table for {build_object.name}: {e}",
                level=logging.ERROR
            )
        worker_names = build_object.workers.values_list('name', flat=True)
        add_many_users_to_results_photos(build_object.sh_url, worker_names)


@admin.action(description="<<< Create table for calculating users KPIs >>>")
def create_kpi_tables_in_sheet1(modeladmin, request, queryset):
    for obj in queryset:
        run_write_kpi_users_in_sheet1(obj.sh_url, obj.id)


class BuildObjectAdmin(admin.ModelAdmin):
    actions = [run_update_users_kpi, create_kpi_tables_in_sheet1, run_update_building_kpi, create_work_time_table, create_photos_table_handle]
    fields = ('name', 'total_budget', 'material', 'sh_url')
    list_display = ('name', 'total_budget', 'display_materials')

    def display_materials(self, obj):
        return ", ".join([material.name for material in obj.material.all()])

    display_materials.short_description = 'Materials'


# @admin.register(BuildObject)
# class BuildObjectAdmin(admin.ModelAdmin):
#     actions = [run_update_users_kpi, run_update_building_kpi, create_work_time_table]


admin.site.register(Worker, WorkerAdmin)
admin.site.register(BuildObject, BuildObjectAdmin)
admin.site.register(Material, MaterialAdmin)
admin.site.register(ToolsSheet, ToolsSheetAdmin)
admin.site.register(Tool, ToolAdmin)

