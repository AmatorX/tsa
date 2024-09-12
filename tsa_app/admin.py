from django.contrib import admin, messages
from .models import Worker, BuildObject, Material, Tool, ToolsSheet
# from ..utils.add_users_to_work_time import append_user_names_to_tables
from django.contrib.admin import SimpleListFilter


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
    list_per_page = 20
    # inlines = (WorkerProfileInline,)

    def save_model(self, request, obj, form, change):
        super().save_model(request, obj, form, change)
        if not change:  # Проверяем, что объект был создан, а не обновлен
            pass
            # append_user_names_to_tables(obj)  # Ваша функция для записи данных
            # Отображаем сообщение в админке
            # messages.success(request, 'Данные записаны успешно!')


class BuildObjectAdmin(admin.ModelAdmin):
    fields = ('name', 'total_budget', 'material', 'sh_url')
    list_display = ('name', 'total_budget', 'display_materials')

    def display_materials(self, obj):
        return ", ".join([material.name for material in obj.material.all()])

    display_materials.short_description = 'Materials'


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


admin.site.register(Worker, WorkerAdmin)
admin.site.register(BuildObject, BuildObjectAdmin)
admin.site.register(Material, MaterialAdmin)
admin.site.register(ToolsSheet, ToolsSheetAdmin)
admin.site.register(Tool, ToolAdmin)

