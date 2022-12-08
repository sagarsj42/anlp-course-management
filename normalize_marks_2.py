from collections import defaultdict

import gspread
import numpy as np


marks_worksheet_name = 'anlp-marksheet'
sheet_name = 'assgn3'
tas = ['Sagar', 'Suyash', 'Tanvi', 'Veeral']

lb = 5
ub = 95

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
all_marks = list()
all_marks_nz = list()

for row in records:
    ta = row[ta_index]
    if not ta:
        continue
    marks = row[marks_index]
    ta_marks[ta].append(float(marks))
    all_marks.append(float(marks))
for ta in tas:
    ta_marks[ta] = np.array(ta_marks[ta])

all_marks = np.array(all_marks)

stats_col = list()
stats_col.append('Stats')
stats_col.append('Lower bound')
stats_col.append(lb)
stats_col.append('Upper bound')
stats_col.append(ub)
stats_col.append('')
stats_col.append('Full range per TA')
print('Full range per TA')

for ta, marks in ta_marks.items():
    stats_col.append(ta)
    stats_col.append(marks.mean())
    stats_col.append(marks.std())
    print(ta, marks.mean(), marks.std())
stats_col.append('')

stats_col.append('Full range mean:')
stats_col.append(all_marks.mean())
stats_col.append('Full range std:')
stats_col.append(all_marks.std())
stats_col.append('')
stats_col.append('Normalizable stats')
stats_col.append('')
print('Normalizable stats:')

for ta, marks in ta_marks.items():
    nz = [float(m) for m in marks if m > lb and m < ub]
    ta_marks_nz[ta] = np.array(nz)
    all_marks_nz.extend(nz)
    stats_col.append(ta)
    stats_col.append(ta_marks_nz[ta].mean())
    stats_col.append(ta_marks_nz[ta].std())
    print(ta, ta_marks_nz[ta].mean(), ta_marks_nz[ta].std())
stats_col.append('')

all_marks_nz = np.array(all_marks_nz)
common_mean = all_marks_nz.mean()
common_std = all_marks_nz.std()

stats_col.append('Common mean: normalizable')
stats_col.append(common_mean)
stats_col.append('Common std: normalizable')
stats_col.append(common_std)
stats_col.append('')

print('Common mean:', common_mean)
print('Common std:', common_std)
print('Normalizing to common mean of', common_mean)

ta_norm_marks = defaultdict(list)
normalized = ['Normalized']
for record in records:
    ta = record[ta_index]
    m = float(record[marks_index])

    if not ta:
        normalized.append(0)
        continue
    
    tmnz = ta_marks_nz[ta]
    t_mean = tmnz.mean()

    if m <= lb or m >= ub:
        nm = round(m)
    elif t_mean >= common_mean:
        nm = round(m)
    else:
        diff = common_mean - t_mean
        nm = round(min(m + diff, ub))
    normalized.append(nm)
    ta_norm_marks[ta].append(nm)

stats_col.append('Normalized stats, full:')
print('Normalized stats, full:')
for ta, marks in ta_norm_marks.items():
    marks = np.array(marks)
    ta_marks[ta] = marks
    stats_col.append(ta)
    stats_col.append(marks.mean())
    stats_col.append(marks.std())
    print(ta, marks.mean(), marks.std())
stats_col.append('')

ws.update(f'G1:G{len(stats_col)}', [[s] for s in stats_col])
ws.update(f'H1:H{len(normalized)}', [[nm] for nm in normalized])
