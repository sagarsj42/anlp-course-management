import random
from collections import defaultdict

import gspread
from gspread.exceptions import APIError


def create_or_get_worksheet(sh, sheet_name, nrows=100, ncols=100):
    try:
        ws = sh.add_worksheet(title=sheet_name, rows=nrows, cols=ncols)
    except APIError as e:
        ws = sh.worksheet(sheet_name)
    
    return ws


sa = gspread.service_account('../sagar-sa-key.json')

tas = ['sagar', 'suyash', 'tanvi', 'veeral']
nonta_projects = [5, 8, 9, 10, 11, 12, 14, 16, 17]
all_projects = list(range(1, 18))
random.shuffle(nonta_projects)
ta_nprojects = {k: v for k, v in list(zip(tas, [5, 4, 4, 4]))}

project_mentors_wb_name = 'anlp-project-mentors'
project_mentors_sheet_name = 'mentors'
project_mentors_pno_col = 'Project No.'
project_mentors_mentor_col = 'Name'

project_teams_wb_name = 'anlp-project-allocation'
project_teams_sheet_name = 'projects-alloted'
project_teams_tno_col = 'Team No.'
project_teams_project_col = 'Project'

teams_wb_name = 'anlp-project-teams'
teams_sheet_name = 'final'

teams_team_no_indx = 0
teams_name_indx = 1
mspan_size = 4
teams_m1_s_indx = 2
teams_m2_s_indx = 6
teams_m3_s_indx = 10

evals_wb_name = 'anlp-assgn3-project-evaluations'
evals_sheet_name = 'slots'

project_marks_wb_name = 'anlp-project-final-marking'
assgn_marks_wb_name = 'anlp-assgn3-marksheet'

project_mentors_wb = sa.open(project_mentors_wb_name)
project_mentors_sheet = project_mentors_wb.worksheet(project_mentors_sheet_name)
project_mentors_data = project_mentors_sheet.get_all_values()
project_mentors_headers = project_mentors_data[:2]
project_mentors_records = project_mentors_data[2:]

project_teams_wb = sa.open(project_teams_wb_name)
project_teams_sheet = project_teams_wb.worksheet(project_teams_sheet_name)
project_teams_data = project_teams_sheet.get_all_values()
project_teams_header = project_teams_data[0]
project_teams_records = project_teams_data[1:]

teams_wb = sa.open(teams_wb_name)
teams_sheet = teams_wb.worksheet(teams_sheet_name)
teams_data = teams_sheet.get_all_values()
teams_records = teams_data[2:]

evals_wb = sa.open(evals_wb_name)
evals_sheet = evals_wb.worksheet(evals_sheet_name)

project_marks_wb = sa.open(project_marks_wb_name)
assgn_marks_wb = sa.open(assgn_marks_wb_name)

project_mentors_pno_index = project_mentors_headers[0].index(project_mentors_pno_col)
project_mentors_mentor_index = project_mentors_headers[1].index(project_mentors_mentor_col)

project_teams_tno_index = project_teams_header.index(project_teams_tno_col)
project_teams_project_index = project_teams_header.index(project_teams_project_col)

ta_project_dict = defaultdict(list)
for ta in tas:
    for pno in all_projects[:15]:
        mentor = [r for r in project_mentors_records if int(r[project_mentors_pno_index]) == pno][0][project_mentors_mentor_index]
        if ta in mentor.lower():
            ta_project_dict[ta].append(pno)

i = 0
for ta in tas:
    c = ta_nprojects[ta] - len(ta_project_dict[ta])
    while c > 0:
        ta_project_dict[ta].append(nonta_projects[i])
        i += 1
        c -= 1

print('tawise-projects:', ta_project_dict)

project_ta_dict = dict()
for ta, pnos in ta_project_dict.items():
    for pno in pnos:
        project_ta_dict[pno] = ta

tno_ta_dict = defaultdict(list)
for record in project_teams_records:
    tno = record[project_teams_tno_index]
    project = record[project_teams_project_index]
    
    pno = int(project.split('-')[0].strip())
    ta = project_ta_dict[pno]
    tno_ta_dict[tno] = (ta, pno)

ta_records = defaultdict(list)
for record in teams_records:
    tno = record[teams_team_no_indx]
    name = record[teams_name_indx]
    ta, pno = tno_ta_dict[tno]
    members = list()

    for member_start in [teams_m1_s_indx, teams_m2_s_indx, teams_m3_s_indx]:
        info = list()
        for i in range(member_start, member_start+4):
            val = record[i]
            if not val:
                break
            info.append(val)
        if info:
            members.append(info)
    
    ta_records[ta].append((tno, name, pno, members))

for ta, records in ta_records.items():
    random.shuffle(records)
    ta_records[ta] = records

all_records = list()
for ta, teams in ta_records.items():
    for team in teams:
        ta = ta.capitalize()
        first_record = [team[0], team[1], team[2]] + team[3][0] + [ta]
        all_records.append(first_record)
        for member in team[3][1:]:
            record = ['']*3 + member + ['']
            all_records.append(record)
        all_records.append([''] * 8)

output_header = ['Team No.', 'Team Name', 'Project No.', 'Member', '', '', '', 'TA']
all_records = [output_header] + all_records

evals_sheet.update(f'A1:H{len(all_records)}', all_records)

project_marks_header = ['Team No.', 'Team Name', 'Project No.', 'First Name', 'Last Name', 'Roll No.']
for ta, teams in ta_records.items():
    sheet_records = [project_marks_header]
    for team in teams:
        first_record = [team[0], team[1], team[2]] + team[3][0][:3]
        sheet_records.append(first_record)
        for member in team[3][1:]:
            record = ['']*3 + member[:3]
            sheet_records.append(record)
        sheet_records.append([''] * 6)
    ta_sheet = create_or_get_worksheet(project_marks_wb, ta)
    ta_sheet.update(f'A1:F{len(sheet_records)}', sheet_records)

assgn_marks_header = ['First Name', 'Last Name', 'Roll No.']
for ta, teams in ta_records.items():
    sheet_records = [assgn_marks_header]
    for team in teams:
        for member in team[3]:
            sheet_records.append(member[:3])
    ta_sheet = create_or_get_worksheet(assgn_marks_wb, ta)
    ta_sheet.update(f'A1:C{len(sheet_records)}', sheet_records)
