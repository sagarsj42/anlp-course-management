import csv
import gspread


marks_worksheet_name = 'anlp-marksheet'
wb_name = 'assgn1'
marks_col = 'Marks'

gradesheet_path = '../assignments/assignment-1/assgn-1-gradesheet.csv'
out_col = 'Assignment: Assignment 1: Language Modelling (Real)'
id_key = 'Username'

sa = gspread.service_account('../sagar-sa-key.json')
sh2 = sa.open(marks_worksheet_name)
mws = sh2.worksheet(wb_name)

records = mws.get_all_values()
header = records[0]

id_index = header.index(id_key)
marks_index = header.index(marks_col)
records = records[1:]

print(header)

csv_rows = list()
with open(gradesheet_path, 'r') as f:
    reader = csv.DictReader(f)

    for row in reader:
        idx = row[id_key]
        record = [r for r in records if r[id_index] == idx][0]
        row[out_col] = record[marks_index]
        csv_rows.append(row)

print(len(csv_rows))

with open(gradesheet_path, 'w') as f:
    writer = csv.DictWriter(f, fieldnames=csv_rows[0].keys())

    writer.writeheader()
    for row in csv_rows:
        writer.writerow(row)
