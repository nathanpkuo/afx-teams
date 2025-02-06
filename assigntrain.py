"""
AFX Training Team Assignment Script

Required packages: pandas, openpyxl
Run "pip install pandas" and "pip install openpyxl"

Cd into the directory containing the training team preference Excel file.
This directory must also contain the proj-assignment.py script.
Run this script as "python assigntrain.py ROSTER_EXCEL_FILENAME BOARD_EXCEL_FILENAME TRAINING_SELECT_SIZE TRAINING_TEAM_SIZE"

Creates an Excel file containing the following sheets:
sheet 0: remaining auditionees after assignment (audition number, name, email) <-- make sure to remove board members!
sheets 1-n: training team assignments (audition number, name, email)

Common Errors:
- IndexError: pop from an empty deque --> insufficient team picks or too much overlap, request additional preferences
"""

import sys
import pandas as pd
import numpy as np
from collections import deque
import random
from assignproj import ranked_assignment
random.seed(42)

"""
Read in roster data via Excel file formatted as follows:
***NO COLUMN HEADERS!***
sheet 0: master auditionee list after project team assignment (audition number, name, email) <-- numbers must be unique!
sheets 1-n: training team preferences (audition number, name, email)

Read in board pick data via Excel file formatted as follows:
***NO COLUMN HEADERS!***
sheets 0-(n-1): board member picks for each team (audition number, name, email) <-- teams in same order as in roster!

SHEETS MUST HAVE THE SAME NAMES (i.e., the team name) IN BOTH EXCEL FILES
"""

# input file and parameters
if len(sys.argv) != 5:
    print('Expected 5 arguments total, received ' + str(len(sys.argv)), file=sys.stderr)
    sys.exit(1)
roster_excel = sys.argv[1]
board_excel = sys.argv[2]
try:
    train_select_size = int(sys.argv[3])
except ValueError as e:
    print('Invalid select size ' + sys.argv[3] + ' provided, must be an integer', file=sys.stderr)
    sys.exit(1)
try:
    train_team_size = int(sys.argv[4])
except ValueError as e:
    print('Invalid team size ' + sys.argv[4] + ' provided, must be an integer', file=sys.stderr)
    sys.exit(1)

# read in roster data from Excel file: creates dictionary linking sheet name to dataframe
df_collection = pd.read_excel(roster_excel, sheet_name=None, header=None)

# read in board picks from Excel file: creates dictionary linking team to board members
board_picks = pd.read_excel(board_excel, sheet_name=None, header=None)

roster_title = list(df_collection.keys())[0]
train_team_names = list(df_collection.keys())[1:]

# assignment_dict contains pairs of form TEAM_NAME:AUDITION_NUMBER_LIST
# remaining_list contains all audition numbers not already assigned to a training team 
assignment_dict, remaining_list = ranked_assignment(df_collection, train_select_size)

# add in board picks
max_team_size_before_rand = 0
for key in train_team_names:
    if not (board_picks[key].empty):
        assignment_dict[key].extend(board_picks[key].iloc[:, 0].tolist())
    max_team_size_before_rand = max(max_team_size_before_rand, len(assignment_dict[key]))

# even out team sizes with random assignment
for key in train_team_names:
    while len(assignment_dict[key]) < max_team_size_before_rand:
        selected = random.choice(remaining_list)
        assignment_dict[key].append(selected)
        remaining_list.remove(selected)

rand_assign_size = train_team_size - max_team_size_before_rand

def next_rand_assign(remaining_nums):
    next_elems = {k:[] for k in train_team_names}
    for key in next_elems:
        if len(remaining_nums) == 0:
            break
        selected = random.choice(remaining_nums)
        next_elems[key].append(selected)
        remaining_nums.remove(selected)
    return next_elems

randomly_selected_dict = {k:[] for k in train_team_names}

i = 0
while (i < rand_assign_size) and (len(remaining_list) > 0):
    next_to_add = next_rand_assign(remaining_list)
    for key in train_team_names:
        randomly_selected_dict[key].extend(next_to_add[key])
    i += 1

for key in train_team_names:
    assignment_dict[key].extend(randomly_selected_dict[key])

remaining_roster = df_collection[roster_title][df_collection[roster_title].iloc[:, 0].isin(remaining_list)]
remaining_roster = remaining_roster.reset_index(drop=True)

with pd.ExcelWriter('train_output.xlsx') as writer:
    remaining_roster.to_excel(writer, sheet_name='waitlist_roster', index=False, header=False)
    for key in train_team_names:
        df = df_collection[roster_title][df_collection[roster_title].iloc[:, 0].isin(assignment_dict[key])]
        # df2 = df_collection[roster_title][df_collection[roster_title].iloc[:, 0].isin(randomly_selected_dict[key])]
        # df = df_collection[roster_title][df_collection[roster_title].iloc[:, 0].isin(assignment_dict[key])]
        # df = df.reset_index(drop=True)
        # df = pd.concat([df1, df2], axis=0, ignore_index=True)
        df.to_excel(writer, sheet_name=key, index=False, header=False)