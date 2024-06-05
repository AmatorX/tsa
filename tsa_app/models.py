from django.db import models
from django.core.validators import MinValueValidator, MaxValueValidator


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


class Tool(models.Model):
    name = models.CharField(max_length=255, blank=True)
    tool_id = models.CharField(max_length=255, blank=True, unique=True)
    date_of_issue = models.DateField(blank=True, null=True)
    assigned_to = models.ForeignKey('Worker', related_name='tools', on_delete=models.SET_NULL, null=True, blank=True)

    def __str__(self):
        return self.name
