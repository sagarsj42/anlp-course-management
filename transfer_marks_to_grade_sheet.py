import csv
import gspread


sa = gspread.service_account('../sagar-sa-key.json')

input_output_sheet_map = {
    'assgn1': 'Assignment: Assignment 1: Language Modelling (Real)',
    'assgn1-bonus': 'Assignment: Assignment 1 - Bonus (Real)',
    'assgn2': 'Assignment: Assignment 2: ELMo (Real)',
    'assgn2-bonus': 'Assignment: Assignment 2 - Bonus (Real)',
    'assgn3': 'Assignment: Assignment 3: Pointer-Generator Networks (Real)',
    'assgn3-bonus': 'Assignment: Assignment 3 - Bonus (Real)',
    'outline': 'Assignment: Project Outline (Real)',
    'interim': 'Assignment: Interim Submission (Real)',
    'project-final': 'Assignment: Final Submission (Real)'
}

input_col_map = {
    'assgn1': 'Marks',
    'assgn1-bonus': 'Marks',
    'assgn2': 'Marks',
    'assgn2-bonus': 'Marks',
    'assgn3': 'Marks',
    'assgn3-bonus': 'Marks',
    'outline': 'Marks',
    'interim': 'Marks',
    'project-final': 'Normalized'
}

marks_worksheet_name = 'anlp-marksheet'

gradesheet_path = '../gradesheet-v3.0.csv'
id_key = 'Username'
outsheet_path = '../graded.csv'

input_wb = sa.open(marks_worksheet_name)

csv_rows = list()
with open(gradesheet_path, 'r') as f:
    reader = csv.DictReader(f)
    for row in reader:
        csv_rows.append(row)

print('# output rows:', len(csv_rows))
    
for isheet_name, ocol_name in input_output_sheet_map.items():
    input_ws = input_wb.worksheet(isheet_name)

    records = input_ws.get_all_values()
    header = records[0]

    id_index = header.index(id_key)
    marks_index = header.index(input_col_map[isheet_name])
    records = records[1:]

    print(isheet_name, header)

    for row in csv_rows:
        idx = row[id_key]
        record = [r for r in records if r[id_index] == idx][0]
        row[ocol_name] = record[marks_index]

with open(outsheet_path, 'w') as f:
    writer = csv.DictWriter(f, fieldnames=csv_rows[0].keys())

    writer.writeheader()
    for row in csv_rows:
        writer.writerow(row)
