"""
AFX Project Team Assignment Script

Required packages: pandas, openpyxl
Run "pip install pandas" and "pip install openpyxl"

Cd into the directory containing the project team preference Excel file.
Run this script as "python proj-assignment.py EXCEL_FILENAME PROJECT_TEAM_SIZE"

Creates an Excel file containing the following sheets:
sheet 0: remaining auditionees after assignment (audition number, name)
sheets 1-n: project team assignments (audition number, name)

Common Errors:
- IndexError: pop from an empty deque --> insufficient team picks or too much overlap, request additional preferences
"""

import sys
import pandas as pd
import numpy as np
from collections import deque
import random
random.seed(42) # this can be changed if desired

"""
Read in data via Excel file formatted as follows:
***NO COLUMN HEADERS!***
sheet 0: master auditionee list (audition number, name) <-- numbers must be unique!
sheets 1-n: project team preferences (audition number, name)
"""

# input file and parameters
if len(sys.argv) != 3:
    print('Expected 3 arguments, received ' + str(len(sys.argv)), file=sys.stderr)
    sys.exit(1)
excel_file = sys.argv[1]
try:
    proj_team_size = int(sys.argv[2])
except ValueError as e:
    print('Invalid team size ' + sys.argv[2] + ' provided, must be an integer', file=sys.stderr)
    sys.exit(1)

# read in data from excel file: creates dictionary linking sheet name to dataframe
df_collection = pd.read_excel(excel_file, sheet_name=None, header=None)

roster_title = list(df_collection.keys())[0]
proj_team_names = list(df_collection.keys())[1:]

remaining = set()
deque_dict = {k:deque() for k in proj_team_names}
for key in df_collection.keys():
    # check that audition numbers are all integers
    df = df_collection[key]
    if df.iloc[:, 0].dtype != np.int64:
        raise Exception('Audition numbers not all integers')
    if key == roster_title:
        if len(df.iloc[:, 0]) != len(set(list(df.iloc[:, 0]))):
            raise Exception('Audition numbers not unique')
        remaining.update(set(list(df.iloc[:, 0])))
    else:
        deque_dict[key].extend(df.iloc[:, 0])

def exactly_once(lst):
    result = []
    for elem in lst:
        if lst.count(elem) == 2:
            result.append(elem)
    return result

team_assignments = {k:[] for k in proj_team_names}
curr_top_picks = {k:v for (k,v) in zip(proj_team_names, [deque_dict[j].popleft() for j in proj_team_names])}

def assign():
    final_picks = {k:-1 for k in proj_team_names}
    unique_nums = exactly_once(list(curr_top_picks.values()))
    for key in proj_team_names:
        if (curr_top_picks[key] in unique_nums) and (curr_top_picks[key] in remaining):
            remaining.remove(curr_top_picks[key])
            final_picks[key] = curr_top_picks[key]
    while -1 in final_picks.values():
        unfinished = [k for k,v in final_picks.items() if v == -1]
        selected = random.choice(unfinished)
        if curr_top_picks[selected] in remaining:
            remaining.remove(curr_top_picks[selected])
            final_picks[selected] = curr_top_picks[selected]
        else:
            curr_top_picks[selected] = deque_dict[selected].popleft()
    for key in proj_team_names:
        team_assignments[key].append(final_picks[key])
        curr_top_picks[key] = deque_dict[key].popleft()

i = 0
while i < proj_team_size:
    assign()
    i += 1

print(team_assignments)
print(deque_dict)
print(remaining)

remaining_roster = df_collection[roster_title][df_collection[roster_title].iloc[:, 0].isin(remaining)]
remaining_roster = remaining_roster.reset_index(drop=True)
print(remaining_roster)

with pd.ExcelWriter('output.xlsx') as writer:
    remaining_roster.to_excel(writer, sheet_name='remaining-roster', index=False, header=False)
    for key in proj_team_names:
        df = df_collection[key][df_collection[key].iloc[:, 0].isin(team_assignments[key])]
        df.to_excel(writer, sheet_name=key, index=False, header=False)

"""
remaining = []
deque_dict = {k:deque() for k in proj_team_names}
for key in df_collection.keys():
    # check that audition numbers are all integers
    df = df_collection[key]
    if df.iloc[:, 0].dtype != np.int64:
        raise Exception('Audition numbers not all integers')
    if key == roster_title:
        if list(df.iloc[:, 0]) != list(set(list(df.iloc[:, 0]))):
            raise Exception('Audition numbers not sorted or not unique')
        remaining.extend(df.iloc[:, 0])
    else:
        deque_dict[key].extend(df.iloc[:, 0])

team_assignments = {k:[] for k in proj_team_names}
# curr_top_picks = {k:v for (k, v) in zip(proj_team_names, [df.iloc[0, 0] for df in df_collection.values()][1:])}
curr_top_picks = {k:v for (k,v) in zip(proj_team_names, [deque_dict[j].popleft() for j in proj_team_names])}

print(team_assignments)
print(remaining)
print(curr_top_picks)
print(deque_dict)

def assign():
    final_picks = {k:-1 for k in proj_team_names}
    while -1 in final_picks.values():
        unique_nums = set(list(curr_top_picks.values()))
        # check if a top pick has already been assigned
        for key in [k for k,v in final_picks.items() if v == -1]:
            if curr_top_picks[key] in unique_nums:
                final_picks[key] = curr_top_picks[key]
            else:
                
                curr_top_picks[key] = 


    while -1 in final_picks.values():
        unique_nums = set(list(curr_top_picks.values()))
        for key in curr_top_picks.keys():
            if curr_top_picks[key] in unique_nums:
                final_picks[key] = curr_top_picks[key]
            else:


    for key in curr_top_picks.keys():
        selected = curr_top_picks[key]
        team_assignments[key].append(selected)



# def assign(curr_top_dict):
#     curr_top_list = list(curr_top_dict.values())
#     unique_nums = set(curr_top_list)
#     for key in curr_top_dict:
#         # number is unique, i.e. only one team picked that person for that ranking
#         if curr_top_dict[key] in unique_nums:
#             team_assignments[key].append(curr_top_dict[key])
#         # number is not unique, i.e. multiple teams picked that person for that ranking
#         else:
#             # break ties randomly

        
#     # for item in curr_top_list:
#     #     # if number is unique, add them to the team
#     #     if item in unique_nums:
#     #         team_assignments
#     #     # if there is a tie, randomly pick


# while any(len(l) < proj_team_size for l in team_assignments.values()):
#     assign(curr_top_picks)
"""