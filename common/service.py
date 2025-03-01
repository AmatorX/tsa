import googleapiclient.discovery
from google.oauth2 import service_account


class GoogleSheetsService:
    """
    credentials_file='./google_sheets_key.json' использовать при запуске django через python manage.py runserver
    credentials_file='../google_sheets_key.json' при тестировании отдельных модулей, что бы запустить их
    """
    # def __init__(self, credentials_file='../google_sheets_key.json'):
    def __init__(self, credentials_file='./google_sheets_key.json'):
        self.credentials_file = credentials_file
        self.service = self._build_service()

    def _build_service(self):
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        credentials = service_account.Credentials.from_service_account_file(
            self.credentials_file, scopes=scopes)
        service = googleapiclient.discovery.build('sheets', 'v4', credentials=credentials, cache_discovery=False)
        return service

    def get_service(self):
        return self.service


# Создание глобального объекта сервиса
service = GoogleSheetsService()
