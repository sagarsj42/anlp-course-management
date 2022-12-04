import gspread


sa = gspread.service_account('../sagar-sa-key.json')
tas = ['sagar', 'suyash', 'tanvi', 'veeral']

input_wb_name = 'anlp-project-final-marking'
input_ta_col = 'TA'
input_marks_col = 'Total'
input_rno_col = 'Roll No.'

input_wb = sa.open(input_wb_name)

output_wb_name = 'anlp-marksheet'
output_sheet_name = 'project-final'
output_rno_col = 'ID number'
output_ta_col = 'TA'
output_marks_col = 'Marks'

output_wb = sa.open(output_wb_name)
output_ws = output_wb.worksheet(output_sheet_name)
output_data = output_ws.get_all_values()
output_header = output_data[0]
output_ta_indx = output_header.index(output_ta_col)
output_rno_indx = output_header.index(output_rno_col)
output_marks_indx = output_header.index(output_marks_col)
for o in output_data[1:]:
    o[output_marks_indx] = '0'

for ta in tas:
    ta_input_ws = input_wb.worksheet(ta)
    input_data = ta_input_ws.get_all_values()
    input_header = input_data[0]
    input_records = input_data[3:]
    input_marks_indx = input_header.index(input_marks_col)
    input_rno_indx = input_header.index(input_rno_col)

    for record in input_records:
        if not record[input_rno_indx]:
            continue
        if record[input_marks_indx] != '0':
            marks = record[input_marks_indx]

        rno = record[input_rno_indx]
        output_indxs = [i for i, o in enumerate(output_data) if o[output_rno_indx] == rno]
        for output_indx in output_indxs:
            output_data[output_indx][output_ta_indx] = ta.capitalize()
            output_data[output_indx][output_marks_indx] = marks

output_ws.update(f'A1:H{len(output_data)}', output_data)
