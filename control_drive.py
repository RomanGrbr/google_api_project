# import os
#
# from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient import discovery
from pprint import pprint

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# load_dotenv()
#
# info = {
#     'type':  os.environ['TYPE'],
#     'project_id':  os.environ['PROJECT_ID'],
#     'private_key_id':  os.environ['PRIVATE_KEY_ID'],
#     'private_key':  os.environ['PRIVATE_KEY'],
#     'client_email':  os.environ['CLIENT_EMAIL'],
#     'client_id':  os.environ['CLIENT_ID'],
#     'auth_uri':  os.environ['AUTH_URI'],
#     'token_uri':  os.environ['TOKEN_URI'],
#     'auth_provider_x509_cert_url':  os.environ['AUTH_PROVIDER_X509_CERT_URL'],
#     'client_x509_cert_url':  os.environ['CLIENT_X509_CERT_URL']
# }

CREDENTIALS_FILE = 'academic-oath-414305-dab9a8b92d1b.json'


def auth():
    """Авторизация."""
    credentials = Credentials.from_service_account_file(
        filename=CREDENTIALS_FILE, scopes=SCOPES)
    service = discovery.build('drive', 'v3', credentials=credentials)
    return service


def get_list_obj(service):
    """Получить файлы."""
    # response = service.files().list()  # Вернет все документы и папки
    response = service.files().list(q="name='Бюджет путешествий'")  # Поиск по полю
    return response.execute()


def clear_disk(service, spreadsheets):
    """Удалить файлы."""
    for spreadsheet in spreadsheets:
        response = service.files().delete(fileId=spreadsheet['id'])
        response.execute()


service = auth()
pprint(get_list_obj(service))
# spreadsheets = get_list_obj(service)['files']
# clear_disk(service, spreadsheets)
