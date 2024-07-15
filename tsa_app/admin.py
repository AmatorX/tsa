from django.contrib import admin, messages
from .models import Worker, BuildObject, Material, Tool, ToolsSheet
# from ..utils.add_users_to_work_time import append_user_names_to_tables


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


class BuildObjectAdmin(admin.ModelAdmin):
    fields = ('name', 'total_budget', 'material', 'sh_url')
    list_display = ('name', 'total_budget', 'display_materials')

    def display_materials(self, obj):
        return ", ".join([material.name for material in obj.material.all()])

    display_materials.short_description = 'Materials'


class MaterialAdmin(admin.ModelAdmin):
    fields = ('name', 'price')
    list_display = ('name', 'price')


class ToolAdmin(admin.ModelAdmin):
    fields = ('name', 'tool_id', 'date_of_issue', 'assigned_to', 'tools_sheet')
    list_display = ('name', 'tool_id', 'date_of_issue', 'assigned_to', 'tools_sheet')
    list_editable = ('assigned_to', 'tools_sheet')


class ToolsSheetAdmin(admin.ModelAdmin):
    fields = ('name', 'sh_url')


admin.site.register(Worker, WorkerAdmin)
admin.site.register(BuildObject, BuildObjectAdmin)
admin.site.register(Material, MaterialAdmin)
admin.site.register(ToolsSheet, ToolsSheetAdmin)
admin.site.register(Tool, ToolAdmin)

