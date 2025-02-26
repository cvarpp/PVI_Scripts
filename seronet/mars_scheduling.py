import pandas as pd
import numpy as np
import sys
import util

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("usage: python -m seronet.mars_scheduling MM-DD-YY")
        exit(1)
    else:
        try:
            date_of_interest = pd.to_datetime(sys.argv[1])
        except:
            exit(1)
    mars_tracker = pd.ExcelFile(util.mars_folder + 'MARS Tracker.xlsx')
    umb_tracker = pd.read_excel(util.umbrella_tracker, sheet_name='Summary')
    umb_mars = umb_tracker[umb_tracker['Cohort'] == 'MARS'].set_index('Subject ID')
    participant_schedules = pd.read_excel(mars_tracker, sheet_name='kp.2 boost tracker')
    participant_info = pd.read_excel(mars_tracker, sheet_name='Pt List', index_col='Participant ID')
    visit_cols = ['30 day timepoint', '60 day timepoint', '90 day timepoint', '120 day timepoint', '180 day timepoint']
    all_visit_due = participant_schedules.set_index('Participant ID').loc[:, visit_cols].stack().reset_index()
    all_visit_due.columns = ['Participant ID', 'Visit Kind', 'Date']
    all_visit_due['Last in Window'] = pd.to_datetime(all_visit_due['Date'], errors='coerce') + pd.Timedelta(days=11)
    all_visit_due['First in Window'] = pd.to_datetime(all_visit_due['Date'], errors='coerce') - pd.Timedelta(days=11)
    not_too_late_filter = all_visit_due['Last in Window'] >= date_of_interest
    next_visits = all_visit_due[not_too_late_filter].drop_duplicates(subset='Participant ID').set_index('Participant ID')

    for day in next_visits.index:
        day_of_week = next_visits.loc[day, 'Date']
        if day_of_week.weekday() == 5:
            next_visits.loc[day,'Date'] = next_visits.loc[day,'Date'] - pd.Timedelta(days=1)
        elif day_of_week.weekday() == 6:
            next_visits.loc[day,'Date'] = next_visits.loc[day,'Date'] + pd.Timedelta(days=1) 
    visit_types = ['Day 30 Post COVID Vaccine', 'Day 60 Post COVID Vaccine', 'Day 90 Post COVID Vaccine', 'Day 120 Post COVID Vaccine', 'Day 180 Post COVID Vaccine']
    timepoints = ['1 Month', '2 Month', '3 Month', '4 Month', '6 Month']
    visit_key = {'Name':visit_cols, 'Type': visit_types, 'Timepoint': timepoints}
    key_frame = pd.DataFrame.from_dict(visit_key).set_index('Name')
    
    scheduling_stuff = (participant_info.join(next_visits).join(umb_mars, rsuffix='1', how='inner'))
    due_date = pd.to_datetime(scheduling_stuff['Date'])

    pts_due = []
    for pid in scheduling_stuff.index:
        for day in range (0,7):
            day_check = date_of_interest + pd.Timedelta(days=day)
            if due_date[pid] == day_check:
                pts_due.append(pid)
   
    table = {'Name':[], 'Study':[], 'Visit Details':[], 'NoF':[], 'ID':[], 'MRN':[], 'Email':[], 'Timepoint':[], 'Visit Date':[], 'Can See Starting':[], 'Must See By':[]}
    for pt in pts_due:
        table['Name'].append(scheduling_stuff.loc[pt, 'Name'])
        table['Study'].append('MARS')
        table['Visit Details'].append(key_frame.loc[scheduling_stuff.loc[pt, 'Visit Kind'], 'Type'])
        table['NoF'].append('Follow-Up')
        table['ID'].append(pt)
        table['MRN'].append(scheduling_stuff.loc[pt, 'MRN'])
        table['Email'].append(scheduling_stuff.loc[pt, 'Email'])
        table['Timepoint'].append(key_frame.loc[scheduling_stuff.loc[pt, 'Visit Kind'], 'Timepoint'])
        table['Visit Date'].append(scheduling_stuff.loc[pt, 'Date'].strftime("%A, %B %d, %Y"))
        table['Can See Starting'].append(scheduling_stuff.loc[pt, 'First in Window'].strftime("%A, %B %d, %Y"))
        table['Must See By'].append(scheduling_stuff.loc[pt, 'Last in Window'].strftime("%A, %B %d, %Y"))
    output = pd.DataFrame.from_dict(table)

    destination = util.mars_folder + 'MARS Participants Due Week of {}.xlsx'.format(date_of_interest.strftime("%m.%d.%Y"))
    with pd.ExcelWriter(destination) as writer:
        output.to_excel(writer, sheet_name='{}-{}'.format(date_of_interest.strftime("%m.%d.%y"), (date_of_interest + pd.Timedelta(days=4)).strftime("%m.%d.%y")), index=False)
        print("Written to", destination)
