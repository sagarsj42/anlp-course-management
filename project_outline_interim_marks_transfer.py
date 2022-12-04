import gspread


sa = gspread.service_account('../sagar-sa-key.json')

input_wb_name = 'anlp-project-outline-interim-marking'
input_sheet_name = 'interim-marks'
input_ta_col = 'TA'
input_marks_col = 'Total'
input_team_no_col = 'Team No.'

output_wb_name = 'anlp-marksheet'
output_sheet_name = 'interim'
output_rno_col = 'ID number'
output_ta_col = 'TA'
output_marks_col = 'Marks'

teams_wb_name = 'anlp-project-teams'
teams_sheet_name = 'final'

teams_team_no_indx = 0
teams_m1_rno_indx = 4
teams_m2_rno_indx = 8
teams_m3_rno_indx = 12

input_wb = sa.open(input_wb_name)
input_ws = input_wb.worksheet(input_sheet_name)
input_data = input_ws.get_all_values()
input_header = input_data[0]
input_records = input_data[4:]

output_wb = sa.open(output_wb_name)
output_ws = output_wb.worksheet(output_sheet_name)
output_data = output_ws.get_all_values()
output_header = output_data[0]
outut_records = output_data[1:]

teams_wb = sa.open(teams_wb_name)
teams_ws = teams_wb.worksheet(teams_sheet_name)
teams_data = teams_ws.get_all_values()
teams_header = teams_data[0]
teams_records = teams_data[2:]

input_ta_indx = input_header.index(input_ta_col)
input_marks_indx = input_header.index(input_marks_col)
input_team_no_indx = input_header.index(input_team_no_col)
output_ta_indx = output_header.index(output_ta_col)
output_rno_indx = output_header.index(output_rno_col)
output_marks_indx = output_header.index(output_marks_col)

print('Input sheet header:', input_header)
print('Output sheet header:', output_header)
print('Teams sheet header:', teams_header)

teamno_rno_dict = dict()
for record in teams_records:
    tno = record[teams_team_no_indx]
    rno_1 = record[teams_m1_rno_indx]
    rno_2 = record[teams_m2_rno_indx]
    rno_3 = record[teams_m3_rno_indx]
    rnos = [r for r in (rno_1, rno_2, rno_3) if r]
    teamno_rno_dict[tno] = rnos

print('Teams found:', len(teamno_rno_dict))

for record in input_records:
    ta = record[input_ta_indx]
    tno = record[input_team_no_indx]
    marks = record[input_marks_indx]
    rnos = teamno_rno_dict[tno]
    output_indxs = [i for i, o in enumerate(output_data) if o[output_rno_indx] in rnos]
    for i in output_indxs:
        output_data[i][output_ta_indx] = ta
        output_data[i][output_marks_indx] = marks

output_ws.update(f'A1:F{len(output_data)}', output_data)
