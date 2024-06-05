import googleapiclient.discovery
from google.oauth2 import service_account


class GoogleSheetsService:
    def __init__(self, credentials_file='./google_sheets_key.json'):
        self.credentials_file = credentials_file
        self.service = self._build_service()

    def _build_service(self):
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        credentials = service_account.Credentials.from_service_account_file(
            self.credentials_file, scopes=scopes)
        service = googleapiclient.discovery.build('sheets', 'v4', credentials=credentials)
        return service

    def get_service(self):
        return self.service


# Создание глобального объекта сервиса
service = GoogleSheetsService()
