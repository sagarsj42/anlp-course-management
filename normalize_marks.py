from collections import defaultdict

import gspread
import numpy as np


marks_worksheet_name = 'anlp-marksheet'
sheet_name = 'assgn3'
tas = ['Sagar', 'Suyash', 'Tanvi', 'Veeral']

sa = gspread.service_account('../sagar-sa-key.json')
sh = sa.open(marks_worksheet_name)
ws = sh.worksheet(sheet_name)

records = ws.get_all_values()
header = records[0]
records = records[1:]

print(header)

ta_index = header.index('TA')
marks_index = header.index('Marks')

ta_marks = defaultdict(list)
ta_marks_nz = dict()

for row in records:
    ta = row[ta_index]
    if not ta:
        continue
    ta_marks[ta].append(float(row[marks_index]))
for ta in tas:
    ta_marks[ta] = np.array(ta_marks[ta])

stats_col = list()
stats_col.append('Stats')
stats_col.append('')
stats_col.append('Full range per TA')
print('Full range per TA')

for ta, marks in ta_marks.items():
    stats_col.append(ta)
    stats_col.append(marks.mean())
    stats_col.append(marks.std())
    stats_col.append('')
    print(ta, marks.mean(), marks.std())
stats_col.append('')

stats_col.append('Normalizable stats')
stats_col.append('')
print('Normalizable stats:')

lb = 5
ub = 95

for ta, marks in ta_marks.items():
    ta_marks_nz[ta] = np.array([m for m in marks if m > lb and m < ub])
    stats_col.append(ta)
    stats_col.append(ta_marks_nz[ta].mean())
    stats_col.append(ta_marks_nz[ta].std())
    stats_col.append('')
    print(ta, ta_marks_nz[ta].mean(), ta_marks_nz[ta].std())
stats_col.append('')

common_mean = 79
common_std = 9

stats_col.append('Normalizing to:')
stats_col.append(common_mean)
stats_col.append(common_std)
stats_col.append('')
print('Normalizing to:')
print(common_mean, common_std)

ta_norm_marks = dict()
for ta, marks in ta_marks.items():
    norm_marks = list()
    for m in marks:
        if m <= lb or m >= ub:
            norm_marks.append(round(m))
            continue
        nm = round(min(common_mean + (m - ta_marks_nz[ta].mean()) * common_std / ta_marks_nz[ta].std(), ub))
        norm_marks.append(nm)
    ta_norm_marks[ta] = np.array(norm_marks)

stats_col.append('Normalized stats, full:')
print('Normalized stats, full:')
for ta, marks in ta_norm_marks.items():
    stats_col.append(ta)
    stats_col.append(marks.mean())
    stats_col.append(marks.std())
    stats_col.append('')
    print(ta, marks.mean(), marks.std())
stats_col.append('')

stats_col.append('Normalized stats, normalized values only:')
print('Normalized stats, normalized values only:')
for ta, marks in ta_norm_marks.items():
    marks_nz = np.array([m for m in marks if m > lb and m < ub])

    stats_col.append(ta)
    stats_col.append(marks_nz.mean())
    stats_col.append(marks_nz.std())
    stats_col.append('')
    print(ta, marks_nz.mean(), marks_nz.std())

normalized = ['Normalized']
for record in records:
    ta = record[ta_index]
    m = float(record[marks_index])

    if not ta:
        normalized.append(0)
        continue
    
    tmnz = ta_marks_nz[ta]
    t_mean = tmnz.mean()
    t_std = tmnz.std()

    if m <= lb or m >= ub:
        nm = round(m)
    else:
        nm = round(min(common_mean + (m - t_mean) * common_std / t_std, ub))
    normalized.append(nm)

ws.update(f'G1:G{len(stats_col)}', [[s] for s in stats_col])
ws.update(f'H1:H{len(normalized)}', [[nm] for nm in normalized])
