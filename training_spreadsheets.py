from google.oauth2.service_account import Credentials
from googleapiclient import discovery


SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

CREDENTIALS_FILE = 'academic-oath-414305-dab9a8b92d1b.json'


def auth():
    """Авторизация."""
    credentials = Credentials.from_service_account_file(
        filename=CREDENTIALS_FILE, scopes=SCOPES)
    # Методы для работы с Google API
    service = discovery.build('sheets', 'v4', credentials=credentials)
    return service, credentials


def create_spreadsheet(service):
    """Создание документа."""
    spreadsheet_body = {
        # Свойства документа
        'properties': {
            'title': 'Бюджет путешествий',
            'locale': 'ru_RU'
        },
        # Свойства листов документа
        'sheets': [{
            'properties': {
                'sheetType': 'GRID',
                'sheetId': 0,
                'title': 'Отпуск 2077',
                'gridProperties': {
                    'rowCount': 100,
                    'columnCount': 100
                    }
                }
        }]
    }
    request = service.spreadsheets().create(body=spreadsheet_body)
    response = request.execute()
    spreadsheet_id = response['spreadsheetId']
    print('https://docs.google.com/spreadsheets/d/' + spreadsheet_id)
    return spreadsheet_id


def set_user_permissions(spreadsheet_id, credentials):
    """Предоставление доступа к google таблице."""
    permissions_body = {'type': 'user',  # Тип учетных данных.
                        'role': 'writer',  # Права доступа для учётной записи.
                        'emailAddress': 'romangrbr@gmail.com'}  # Гугл-аккаунт.

    # Создаётся экземпляр класса Resource для Google Drive API.
    drive_service = discovery.build('drive', 'v3', credentials=credentials)

    # Формируется и сразу выполняется запрос на выдачу прав вашему аккаунту.
    drive_service.permissions().create(
        fileId=spreadsheet_id,
        body=permissions_body,
        fields='id'
    ).execute()


def spreadsheet_update_values(spreadsheet_id, service):
    # Данные для заполнения: выводятся в таблице сверху вниз, слева направо.
    table_values = [
        ['Бюджет путешествий'],
        ['Весь бюджет', '5000'],
        ['Все расходы', '=SUM(E7:E30)'],
        ['Остаток', '=B2-B3'],
        ['Расходы'],
        ['Описание', 'Тип', 'Кол-во', 'Цена', 'Стоимость'],
        ['Перелет', 'Транспорт', '2', '400', '=C7*D7']
    ]

    c8 = [['Виски', 'Расслабон', '3', '500', '=C8*D8']]
    c10 = [['Ром', 'Расслабон', '2', '340', '=C10*D10']]

    # Тело запроса.
    request_body = {
        'majorDimension': 'ROWS',
        'values': table_values,
        'range': 'Отпуск 2077!A1:F20'
    }
    request_body2 = {
        'majorDimension': 'ROWS',
        'values': c8,
        'range': 'Отпуск 2077!A8:E8'
    }
    request_body3 = {
        'majorDimension': 'ROWS',
        'values': c10,
        'range': 'Отпуск 2077!A10:E10'
    }
    # Формирование запроса к Google Sheets API.
    # request = service.spreadsheets().values().update(
    #     spreadsheetId=spreadsheet_id,
    #     range='Отпуск 2077!A1:F20',
    #     valueInputOption='USER_ENTERED',
    #     body=request_body
    # )
    request = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            'valueInputOption': 'USER_ENTERED',
            'data': [request_body, request_body2, request_body3]
        }
    )

    # Выполнение запроса.
    request.execute()

def read_values(service, spreadsheet_id):
    range = 'Отпуск 2077!A1:F20'
    response = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range,
    ).execute()
    return response['values']

service, credentials = auth()
spreadsheetId = create_spreadsheet(service)
set_user_permissions(spreadsheetId, credentials)
spreadsheet_update_values(spreadsheetId, service)
# data = read_values(service, spreadsheetId)
# print(data)
