import json
from django.db import models
from django.core.validators import MinValueValidator, MaxValueValidator
from django.utils import timezone


class Worker(models.Model):
    CHOICES = [
        ('YES', 'Yes'),
        ('NO', 'No'),
        ('NA', 'N/A'),
    ]

    EMPLOYMENT_AGREEMENT_CHOICES = CHOICES
    OVER_TIME = CHOICES
    PAYROLL = CHOICES

    name = models.CharField(max_length=255)
    tg_id = models.IntegerField(unique=True, null=True, blank=True)
    foreman = models.BooleanField(default=False)
    build_obj = models.ForeignKey('BuildObject', on_delete=models.SET_NULL, related_name='workers', null=True,
                                  blank=True)
    email = models.EmailField(max_length=255, blank=True, null=True)
    employment_agreement = models.CharField(max_length=3, choices=EMPLOYMENT_AGREEMENT_CHOICES, default='NA')
    over_time = models.CharField(max_length=3, choices=OVER_TIME, default='NA')
    start_to_work = models.DateField(null=True, blank=True)
    title = models.CharField(max_length=255, null=True, blank=True)
    salary = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    payroll_eligible = models.DateField(null=True, blank=True)
    payroll = models.CharField(max_length=3, choices=EMPLOYMENT_AGREEMENT_CHOICES, default='NA')
    resign_agreement = models.DateField(null=True, blank=True)
    benefits_eligible = models.DateField(null=True, blank=True)
    birthday = models.DateField(null=True, blank=True)
    age = models.IntegerField(validators=[MinValueValidator(14), MaxValueValidator(99)], null=True, blank=True)
    issued = models.DateField(null=True, blank=True)
    expiry = models.DateField(null=True, blank=True)
    address = models.CharField(max_length=255, null=True, blank=True)
    phone_number = models.CharField(max_length=35, null=True, blank=True)
    tickets_available = models.CharField(max_length=255, null=True, blank=True)

    # сохранения оригинального значения build_obj_id до сохранения объекта в базе данных
    # необходимо для провеки изменения связи воркера с объектом стройки в сигналах
    def save(self, *args, **kwargs):
        if self.pk:
            self.__original_build_obj_id = Worker.objects.get(pk=self.pk).build_obj_id
        super().save(*args, **kwargs)

    def __str__(self):
        return self.name


class Material(models.Model):
    name = models.CharField(max_length=255)
    price = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    build_obj = models.ManyToManyField('BuildObject', related_name='materials')

    def __str__(self):
        return self.name


class BuildObject(models.Model):
    name = models.CharField(max_length=255)
    total_budget = models.FloatField(null=True, blank=True)
    material = models.ManyToManyField('Material', related_name='build_objects')
    sh_url = models.URLField(null=True, blank=True)

    def __str__(self):
        return self.name


class ToolsSheet(models.Model):
    name = models.CharField(max_length=255)
    sh_url = models.URLField(null=True, blank=True)

    def __str__(self):
        return self.name


class Tool(models.Model):
    name = models.CharField(max_length=255, blank=True)
    tool_id = models.CharField(max_length=255, unique=True)
    # tool_id = models.CharField(max_length=255, blank=True, unique=True)
    price = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, default=0)
    # price = models.FloatField(max_length=16, blank=True, null=True)
    date_of_issue = models.DateField(blank=True, null=True)
    assigned_to = models.ForeignKey('Worker', related_name='tools', on_delete=models.SET_NULL, null=True, blank=True)
    tools_sheet = models.ForeignKey('ToolsSheet', related_name='tools_sheet', on_delete=models.CASCADE, null=True, blank=True)

    def save(self, *args, **kwargs):
        if self.pk:  # Если объект уже существует в базе данных
            previous = Tool.objects.get(pk=self.pk)

            if previous.assigned_to != self.assigned_to:
                if self.assigned_to is None:
                    # Если поле assigned_to изменено с пользователя на None, очищаем date_of_issue
                    self.date_of_issue = None
                else:
                    # Если поле assigned_to изменено с None или другого пользователя на нового
                    # и если дата не установлена, устанавливаем текущую дату
                    if not self.date_of_issue:
                        self.date_of_issue = timezone.now().date()
        else:
            # Если объект новый и assigned_to уже заполнено
            if self.assigned_to and not self.date_of_issue:
                self.date_of_issue = timezone.now().date()

        # Автоматическое присвоение tools_sheet, если оно не указано
        if not self.tools_sheet:
            self.tools_sheet = ToolsSheet.objects.first()

        super().save(*args, **kwargs)

    def __str__(self):
        return self.name


class WorkEntry(models.Model):
    worker = models.ForeignKey('Worker', on_delete=models.CASCADE, related_name='work_entries')
    build_object = models.ForeignKey('BuildObject', on_delete=models.CASCADE, related_name='work_entries')
    worked_hours = models.DecimalField(max_digits=5, decimal_places=2)  # Для хранения отработанных часов
    materials_used = models.TextField()  # Сохраняем данные в формате JSON как текст
    date = models.DateField(auto_now_add=True)  # Дата, когда запись была создана

    def set_materials_used(self, materials_dict):
        """Сохраняет словарь материалов в поле materials_used."""
        self.materials_used = json.dumps(materials_dict)

    def get_materials_used(self):
        """Возвращает словарь материалов из поля materials_used."""
        return json.loads(self.materials_used)

    def get_worked_hours_as_float(self):
        """Возвращает worked_hours как float."""
        return float(self.worked_hours)

    def __str__(self):
        # Используем имя работника и объекта стройки вместо их ID
        worker_name = self.worker.name if self.worker else "Unknown Worker"
        build_object_name = self.build_object.name if self.build_object else "Unknown Object"
        return f"Worker: {worker_name}, Object: {build_object_name}, Date: {self.date}"

