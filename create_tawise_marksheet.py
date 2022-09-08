from collections import defaultdict

import gspread
from gspread.exceptions import APIError


slots_sheet_name = 'anlp-assgn1-evaluations'
marksheet_name = 'assgn1-marksheet'
tas = ['Sagar', 'Suyash', 'Tanvi', 'Veeral']

sa = gspread.service_account('sagar-sa-key.json')
sh = sa.open(slots_sheet_name)
ws = sh.worksheet('slots')
n_rows = ws.row_count
n_cols = ws.col_count

records = ws.get_all_values()[1:]
print(ws.acell('D2').value)
print(ws.acell('D4').value)
print(len(records))

ta_name_col = [c.value for c in ws.range(f'D2:D{len(records)+1}')]
ta_records = defaultdict(list)
current_ta = None
for i, (record, ta_name) in enumerate(zip(records, ta_name_col)):
    if ta_name:
        current_ta = ta_name
    ta_records[current_ta].append(record)

print(len(ta_records))

sh2 = sa.open(marksheet_name)

# for ta in ta_records:
#     records = ta_records[ta]
#     sort(records, key=lambda r: r [0])
#     ta_records[ta] = sorted(records, lambda r: r[1])

def create_or_get_worksheet(sh, sheet_name, nrows=100, ncols=100):
    try:
        ws = sh.add_worksheet(title=sheet_name, rows=nrows, cols=ncols)
    except APIError as e:
        print(e)
        ws = sh.worksheet(sheet_name)
    
    return ws

header = ['First Name', 'Last Name', 'Roll No.', 'Status']

for ta, records in ta_records.items():
    ta_sheet = create_or_get_worksheet(sh2, ta.lower(), nrows=n_rows, ncols=n_cols)
    ta_sheet.update('A1:D1', [header])
    update_records = [r[:3] for r in records]
    e = len(update_records) + 1
    ta_sheet.update(f'A2:C{e}', update_records)
