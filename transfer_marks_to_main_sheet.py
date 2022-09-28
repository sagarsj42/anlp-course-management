import gspread


tawise_sheet_name = 'assgn1-marksheet'
ip_col = 'Assgn. Total'
tas = ['Sagar', 'Suyash', 'Tanvi', 'Veeral']

main_sheet_name = 'anlp-marksheet'
marks_wb_name = 'assgn1'

sa = gspread.service_account('../sagar-sa-key.json')
sh1 = sa.open(tawise_sheet_name)
sh2 = sa.open(main_sheet_name)
mws = sh2.worksheet(marks_wb_name)

out_records = mws.get_all_values()[1:]
rnos = [r[2] for r in out_records]
n_vals = len(rnos)
marks = [0] * n_vals
ta_list = [''] * n_vals

for ta in tas:
    iws = sh1.worksheet(ta.lower())
    ip_records = iws.get_all_values()
    marks_index = ip_records[0].index(ip_col)

    data = ip_records[2:]
    for row in data:
        find = row[2]
        enter = row[marks_index]
        for i in range(n_vals):
            if rnos[i] == find:
                marks[i] = enter
                ta_list[i] = ta

for i in range(n_vals):
    out_records[i][4] = ta_list[i]
    out_records[i][5] = marks[i]

mws.update(f'A2:F{n_vals+1}', out_records)
