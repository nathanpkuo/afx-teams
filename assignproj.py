"""
AFX Project Team Assignment Script

Required packages: pandas, openpyxl
Run "pip install pandas" and "pip install openpyxl"

Cd into the directory containing the project team preference Excel file.
Run this script as "python assignproj.py EXCEL_FILENAME PROJECT_TEAM_SIZE"

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

def ranked_assignment(team_pref_dict, team_size):
    team_names = list(team_pref_dict.keys())[1:]

    remaining = set()
    deque_dict = {k:deque() for k in team_names}
    for key in team_pref_dict.keys():
        # check that audition numbers are all integers
        df = team_pref_dict[key]
        if df.iloc[:, 0].dtype != np.int64:
            raise Exception('Audition numbers not all integers')
        if key == list(team_pref_dict.keys())[0]: # roster_title
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

    team_assignments = {k:[] for k in team_names}
    curr_top_picks = {k:v for (k,v) in zip(team_names, [deque_dict[j].popleft() for j in team_names])}

    def assign():
        final_picks = {k:-1 for k in team_names}
        unique_nums = exactly_once(list(curr_top_picks.values()))
        for key in team_names:
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
        for key in team_names:
            team_assignments[key].append(final_picks[key])
            curr_top_picks[key] = deque_dict[key].popleft()

    i = 0
    while i < team_size:
        assign()
        i += 1

    return team_assignments, list(remaining)

if __name__ == "__main__":
    # input file and parameters
    if len(sys.argv) != 3:
        print('Expected 3 arguments total, received ' + str(len(sys.argv)), file=sys.stderr)
        sys.exit(1)
    excel_file = sys.argv[1]
    try:
        proj_team_size = int(sys.argv[2])
    except ValueError as e:
        print('Invalid team size ' + sys.argv[2] + ' provided, must be an integer', file=sys.stderr)
        sys.exit(1)

    # read in data from Excel file: creates dictionary linking sheet name to dataframe
    df_collection = pd.read_excel(excel_file, sheet_name=None, header=None)

    roster_title = list(df_collection.keys())[0]

    assignment_dict, remaining_list = ranked_assignment(df_collection, proj_team_size)

    remaining_roster = df_collection[roster_title][df_collection[roster_title].iloc[:, 0].isin(remaining_list)]

    with pd.ExcelWriter('proj_output.xlsx') as writer:
        remaining_roster.to_excel(writer, sheet_name='remaining_roster', index=False, header=False)
        for key in list(df_collection.keys())[1:]:
            df = df_collection[key][df_collection[key].iloc[:, 0].isin(assignment_dict[key])]
            df.to_excel(writer, sheet_name=key, index=False, header=False)