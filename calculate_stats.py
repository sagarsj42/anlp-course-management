import gspread
import numpy as np


marks_worksheet_name = 'anlp-marksheet'
sheet_name = 'assgn3'
marks_col = 'Normalized'

sa = gspread.service_account('../sagar-sa-key.json')
sh = sa.open(marks_worksheet_name)
ws = sh.worksheet(sheet_name)

records = ws.get_all_values()
header = records[0]
records = records[1:]
marks_index = header.index(marks_col)

print(header)

all_marks = np.array([float(r[marks_index]) for r in records])

print(sheet_name)
print(f'Mean: {all_marks.mean():.2f}')
print(f'Std: {all_marks.std():.2f}')
print(f'Median: {np.median(all_marks):.2f}')

print(f'({all_marks.mean():.2f}, {all_marks.std():.2f}, {np.median(all_marks):.2f}) [{int(all_marks.max())}]')
