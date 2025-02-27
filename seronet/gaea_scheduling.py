import pandas as pd
import numpy as np
import argparse
import util

def adjust_weekend_dates(date):
    if date.weekday() == 5:  # Saturday
        return date - pd.Timedelta(days=1)
    elif date.weekday() == 6:  # Sunday
        return date + pd.Timedelta(days=1)
    return date

def main(sunday_of_interest):
    gaea_tracker = pd.ExcelFile(util.gaea_folder + 'GAEA Tracker.xlsx')
    participant_schedules = pd.read_excel(gaea_tracker, sheet_name='2024-25 Vaccine Samples')
    participant_info = pd.read_excel(gaea_tracker, sheet_name='Summary', index_col='Participant ID')
    visit_cols = ['Day 30 Due', 'Day 60 Due', 'Day 90 Due', 'Day 120 Due', 'Day 180 Due', 'Day 360 Due']
    all_visit_due = participant_schedules.set_index('Participant ID').loc[:, visit_cols].stack(future_stack=True).reset_index()
    all_visit_due.columns = ['Participant ID', 'Visit Kind', 'Date']
    all_visit_due['Date'] = pd.to_datetime(all_visit_due['Date'], errors='coerce')
    all_visit_due['Last in Window'] = all_visit_due['Date'] + pd.Timedelta(days=11)
    all_visit_due['First in Window'] = all_visit_due['Date'] - pd.Timedelta(days=11)
    not_too_late_filter = all_visit_due['Last in Window'] >= sunday_of_interest
    next_visits = all_visit_due[not_too_late_filter].drop_duplicates(subset='Participant ID').set_index('Participant ID')

    next_visits['Date'] = next_visits['Date'].apply(adjust_weekend_dates)
    visit_types = ['Day 30 Post COVID Vaccine', 'Day 60 Post COVID Vaccine', 'Day 90 Post COVID Vaccine',
                   'Day 120 Post COVID Vaccine', 'Day 180 Post COVID Vaccine', 'Day 360 Post COVID Vaccine']
    timepoints = ['1 Month', '2 Month', '3 Month', '4 Month', '6 Month', '1 Year']
    visit_key = pd.DataFrame({'Name':visit_cols, 'Type': visit_types, 'Timepoint': timepoints}).set_index('Name')
    
    scheduling_stuff = participant_info.join(next_visits)
    due_date = scheduling_stuff['Date']
    date_range = [sunday_of_interest + pd.Timedelta(days=day) for day in range(7)]
    table = scheduling_stuff.loc[due_date.isin(date_range)].reset_index()
    table['Study'] = 'GAEA'
    table['Visit Details'] = table['Visit Kind'].map(visit_key['Type'])
    table['NoF'] = 'Follow-Up'
    table['Timepoint'] = table['Visit Kind'].map(visit_key['Timepoint'])
    table['Visit Date'] = table['Date'].dt.strftime("%A, %B %d, %Y")
    table['Can See Starting'] = table['First in Window'].dt.strftime("%A, %B %d, %Y")
    table['Must See By'] = table['Last in Window'].dt.strftime("%A, %B %d, %Y")
    table = table[['Name', 'Study', 'Visit Details', 'NoF', 'Participant ID', 'Email',
                   'Timepoint', 'Visit Date', 'Can See Starting', 'Must See By']]

    destination = util.gaea_folder + 'GAEA Scheduling Week of {}.xlsx'.format(sunday_of_interest.strftime("%m.%d.%Y"))
    with pd.ExcelWriter(destination) as writer:
        table.to_excel(writer, sheet_name='{}-{}'.format(sunday_of_interest.strftime("%m.%d.%y"),
                                                         (sunday_of_interest + pd.Timedelta(days=4)).strftime("%m.%d.%y")), index=False)
        print("Written to", destination)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Scheduling for GAEA participants.')
    parser.add_argument('date', type=lambda d: pd.to_datetime(d, format='%m-%d-%y'), help='Sunday of Week of Interest in MM-DD-YY format')
    args = parser.parse_args()
    
    main(args.date)
