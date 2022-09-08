import os
import random

# from google.auth.transport.requests import Request
# from google_auth_oauthlib.flow import InstalledAppFlow
# from google.oauth2.credentials import Credentials
# from googleapiclient.discovery import build
# from googleapiclient.errors import HttpError
import gspread
from gspread.exceptions import APIError

# SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
# SPREADSHEET_ID = '1EZjxIIWaHHCroIrxq08yPjynBuVKOxyg36WpCv3qvfA'
# RANGE_NAME = '1985519506'

# creds = None
# if os.path.exists('token.json'):
#     creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# if not creds or not creds.valid:
#     if creds and creds.expired and creds.refresh_token:
#         creds.refresh(Request())
#     else:
#         flow = InstalledAppFlow.from_client_secrets_file('client_secret_2.json', SCOPES)
#         creds = flow.run_local_server(port=0)
#     with open('token.json', 'w') as f:
#         f.write(creds.to_json())

# try:
#     service = build('sheets', 'v4', credentials=creds)
#     sheet = service.spreadsheets()
#     result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
#     values = result.get('values', [])

#     if not values:
#         print('No data found')
    
#     print('Name, Major:')
#     for row in values:
#         print('%s, %s' % (row[0], row[1]))
# except HttpError as e:
#     print(e)

sheet_name = 'anlp-assgn1-evaluations'
tas = ['Sagar', 'Suyash', 'Tanvi', 'Veeral']

sa = gspread.service_account('sagar-sa-key.json')
sh = sa.open(sheet_name)
ws = sh.worksheet('names')

print('Rows:', ws.row_count)
print('Cols:', ws.col_count)

records = ws.get_all_values()
header = records[0]
records = records[1:]

print(type(records), len(records))

print(records[0])
n_tas = len(tas)
n_records = len(records)
record_set = set(range(n_records))

ta_records = dict()
for i in range(n_tas):
    if i == (n_tas-1):
        indices = list(record_set)
        random.shuffle(indices)
        indices = set(indices)
    else:
        indices = random.sample(list(record_set), n_records // 4)
    ta_records[tas[i]] = [records[i] for i in indices]
    record_set -= set(indices)

print(len(record_set))

for ta in ta_records:
    print(ta, len(ta_records[ta]))

try:
    ws2 = sh.add_worksheet(title='slots', rows=ws.row_count, cols=ws.col_count)
except APIError as e:
    print(e)
    ws2 = sh.worksheet('slots')

print(header)
ws2.update('A1:F1', [header])
s = 1

for ta, records in ta_records.items():
    e = s + len(records)
    s += 1
    update_records = [r[:3] for r in records]

    ws2.update(f'D{s}', ta)
    ws2.update(f'A{s}:C{e}', update_records)

    s = e

    # for record in records:
    #     c += 1
    #     ws2.update(f'A{c}:C{c}', [record[:3]])
