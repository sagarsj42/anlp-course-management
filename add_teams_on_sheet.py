import openpyxl
import gspread


sa = gspread.service_account('../sagar-sa-key.json')
output_wb_name = 'anlp-project-teams'
output_sheet_name = 'through-form'
output_wb = sa.open(output_wb_name)
output_ws = output_wb.worksheet(output_sheet_name)
output_data = list()

form_wb_name = '../project/registration-form-results.xlsx'
students_wb_name = '../project/students.xlsx'

teams_ws = openpyxl.load_workbook(form_wb_name)['Sheet1']
students_ws = openpyxl.load_workbook(students_wb_name)['Grades']

teams_ws_data = list(teams_ws.values)
students_ws_data = list(students_ws.values)

tname_header = 'Team Name'
m1_rn_header = 'Member 1: Roll No'
m2_rn_header = 'Member 2: Roll No'
m3_rn_header = 'Member 3: Roll No'
students_fn_header = 'First name'
students_ln_header = 'Surname'
students_rn_header = 'ID number'
students_mail_header = 'Username'

tname_index = teams_ws_data[0].index(tname_header)
m1_index = teams_ws_data[0].index(m1_rn_header)
m2_index = teams_ws_data[0].index(m2_rn_header)
m3_index = teams_ws_data[0].index(m3_rn_header)
m_indices = [m1_index, m2_index, m3_index]

students_fn_index = students_ws_data[0].index(students_fn_header)
students_ln_index = students_ws_data[0].index(students_ln_header)
students_rn_index = students_ws_data[0].index(students_rn_header)
students_mail_index = students_ws_data[0].index(students_mail_header)

teams_data_sorted = sorted(teams_ws_data[1:], key=lambda t: t[tname_index].lower())
student_ids = [s[students_rn_index] for s in students_ws_data]
student_count = 0

for i, teams_row in enumerate(teams_data_sorted):
    tname = teams_row[tname_index]
    n_members = 3 if bool(teams_row[m3_index]) else 2
    team_data = [i+1, tname]

    for n in range(n_members):
        m_rno = teams_row[m_indices[n]]
        student = students_ws_data[student_ids.index(m_rno)]
        student_fname = student[students_fn_index]
        student_lname = student[students_ln_index]
        student_mail = student[students_mail_index]
        student_count += 1

        team_data.extend([student_fname, student_lname, m_rno, student_mail])
    if n_members == 2:
        team_data.extend(['', '', '', ''])
    output_data.append(team_data)
    print(team_data)
    print()

print('# students registered:', student_count)

end_col = (chr(ord('A') + len(output_data[0]) - 1))
end_row = len(output_data)
output_ws.update(f'A1:{end_col}{end_row}', output_data)
