import googleapiclient.discovery
from google.oauth2 import service_account


def get_service():
    # Создание сервиса
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    service_account_file = './google_sheets_key.json'
    credentials = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=scopes)
    service = googleapiclient.discovery.build('sheets', 'v4', credentials=credentials)

    return service
