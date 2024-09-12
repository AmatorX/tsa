from django.apps import AppConfig


class TsaAppConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'tsa_app'

    def ready(self):
        import tsa_app.signals  # Импортируем сигналы для их подключения
        # from .scheduler import start_scheduler
        # print('TsaAppConfig запуск')
        # start_scheduler()

