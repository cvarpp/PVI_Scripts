import pandas as pd
import numpy as np
import datetime
from datetime import date
from dateutil import parser
import csv
import sys
import util

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("usage: python monthly_schedule.py MM-DD-YY")
        exit(1)
    else:
        try:
            date_of_interest = pd.to_datetime(sys.argv[1])
        except:
            exit(1)
    gaea_tracker = pd.ExcelFile(util.gaea_folder + 'GAEA Tracker.xlsx')
    participant_schedules = pd.read_excel(gaea_tracker, sheet_name='2024-25 Vaccine Samples')
    participant_info = pd.read_excel(gaea_tracker, sheet_name='Summary', index_col='Participant ID')
    visit_cols = ['Day 30 Due', 'Day 60 Due', 'Day 90 Due', 'Day 180 Due', 'Day 360 Due']
    all_visit_due = participant_schedules.set_index('Participant ID').loc[:, visit_cols].stack().reset_index()
    all_visit_due.columns = ['Participant ID', 'Visit Kind', 'Date']
    all_visit_due['Last in Window'] = pd.to_datetime(all_visit_due['Date'], errors='coerce') + datetime.timedelta(days=7)
    not_too_late_filter = all_visit_due['Last in Window'] >= date_of_interest
    next_visits = all_visit_due[not_too_late_filter].drop_duplicates(subset='Participant ID').set_index('Participant ID')

    for day in next_visits.index:
        day_of_week = next_visits.loc[day, 'Date']
        if day_of_week.weekday() == 5:
            next_visits.loc[day,'Date'] = next_visits.loc[day,'Date'] - datetime.timedelta(days=1)
        elif day_of_week.weekday() == 6:
            next_visits.loc[day,'Date'] = next_visits.loc[day,'Date'] + datetime.timedelta(days=1) 
    visit_types = ['Day 30 Post COVID Vaccine', 'Day 60 Post COVID Vaccine', 'Day 90 Post COVID Vaccine', 'Day 180 Post COVID Vaccine', 'Day 360 Post COVID Vaccine']
    timepoints = ['1 Month', '2 Month', '3 Month', '6 Month', '1 Year']
    visit_key = {'Name':visit_cols, 'Type': visit_types, 'Timepoint': timepoints}
    key_frame = pd.DataFrame.from_dict(visit_key).set_index('Name')
    
    scheduling_stuff = participant_info.join(next_visits)
    due_date = pd.to_datetime(scheduling_stuff['Date'])

    pts_due = []
    for pid in scheduling_stuff.index:
        for day in range (0,7):
            day_check = date_of_interest + datetime.timedelta(days=day)
            if due_date[pid] == day_check:
                pts_due.append(pid)
   
    table = {'Name':[], 'Study':[], 'Visit Details':[], 'NoF':[], 'ID':[], 'Email':[], 'Timepoint':[], 'Due Date':[]}
    for pt in pts_due:
        table['Name'].append(scheduling_stuff.loc[pt, 'Name'])
        table['Study'].append('GAEA')
        table['Visit Details'].append(key_frame.loc[scheduling_stuff.loc[pt, 'Visit Kind'], 'Type'])
        table['NoF'].append('Follow-Up')
        table['ID'].append(pt)
        table['Email'].append(scheduling_stuff.loc[pt, 'Email'])
        table['Timepoint'].append(key_frame.loc[scheduling_stuff.loc[pt, 'Visit Kind'], 'Timepoint'])
        table['Due Date'].append(scheduling_stuff.loc[pt, 'Date'].strftime("%A, %B %d, %Y"))
    output = pd.DataFrame.from_dict(table)
    print(output.head())

    destination = util.gaea_folder + 'GAEA Scheduling Week of {}.xlsx'.format(date_of_interest.date())
    with pd.ExcelWriter(destination) as writer:
        output.to_excel(writer, sheet_name='{}-{}'.format(date_of_interest.strftime("%m.%d.%y"), (date_of_interest + datetime.timedelta(days=4)).strftime("%m.%d.%y")), index=False)