import pandas as pd
import numpy as np
import datetime
from datetime import date
from dateutil import parser
import csv
import sys
import util

def functor(first_day, val): #rename val to baseline
    if type(val) == datetime.time:
        print("Please make sure every participant has a baseline visit date in the tracker!")
        print("If they have not yet had their baseline visit, please disregard.")
        return datetime.timedelta(days=-1)
    else:
        return first_day - val.date()

if __name__ == '__main__':

    participants = pd.read_excel(util.paris_tracker, sheet_name='Participant details')
    last_visit = pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=util.header_intake).drop_duplicates(subset='Participant ID', keep='last').set_index('Participant ID')
    participants.set_index('Subject ID', inplace=True)
    participants['Recently Seen'] = participants.apply(lambda row: last_visit.loc[row.name, 'Date Collected'] if row.name in last_visit.index.to_numpy() else datetime.date.today(), axis=1)
    completed = participants[participants['Study Status'] == 'Completed']
    participants = participants[participants['Study Status'] == 'Active']
    latest_week = 48
    raw_weeks = range(latest_week, latest_week + 300, 4) # could future-proof this, really goes from 48-128 right now say
    weeks = [datetime.timedelta(days=int(x) * 7) for x in raw_weeks]
    contents = ['blood (5), saliva'] * 70
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    emails = ['morgan.vankesteren@mssm.edu', 'jacob.mauldin@mssm.edu']
    crcs = ['Jessica', 'Jake']
    max_days = 28
    crc_schedule = {} 

    if len(sys.argv) != 2:
        print("usage: python monthly_schedule.py MM-DD-YYYY")
        exit(1)
    else:
        try:
            start_date = parser.parse(sys.argv[1]).date()
        except:
            print("usage: python monthly_schedule.py MM-DD-YYYY")
            exit(1)
    
    for val in range(0, max_days, 7):
        first_day = start_date + datetime.timedelta(days=val)
        crc_schedule[val] = [[] for _ in range(len(emails))]
        ref = participants['Baseline / Day 0 / Week 0'].apply(lambda val: functor(first_day, val))
        for week_num, day_diff, content in zip(raw_weeks, weeks, contents):
            for subj_id in participants.index:
                delta = day_diff - ref.loc[subj_id]
                for idx in range(5): #rename idx weekday for sake of understanding
                    try:
                        if delta == datetime.timedelta(days=idx) and (participants.loc[subj_id, 'Schedule'] == '4 weeks'):
                            crc_email = participants.loc[subj_id, 'CRC Email']
                            if crc_email in emails:
                                crc_schedule[val][emails.index(crc_email)].append((subj_id, week_num, content, idx))
                            print("{} was not added. This is controlled by the 'CRC Email' column in the 'Participant Details' tab".format(subj_id))
                    except:
                        print(subj_id)
                        print(delta)
                        exit(1)

    support_loc = util.clin_ops + 'PARIS/' + 'Scheduling Support PARIS 1 Month from {}.xlsx'.format(start_date)
    writer = pd.ExcelWriter(support_loc)
    ls_col = 'Last Seen as of {}'.format(datetime.date.today())
    completed.loc[:, ['Recently Seen']].to_excel(writer, sheet_name='May Need Reconsent')

    for val in range(0, max_days, 7):
        secret_table = {ls_col: [], 'Email': [], 'Name': [], 'Study': [], 'Week Details': [], 'NoF': [], 'ID': [], 'Location': [], 'Week Number': [], 'Day': [], 'Frequency': [], 'Notes': [] }
        first_day = start_date + datetime.timedelta(days=val)
        for i, subjs in enumerate(crc_schedule[val]):
            subjs = sorted(subjs, key=lambda tup: (tup[3], tup[0]))
            for k in secret_table.keys():
                if k == 'Name':
                    secret_table[k].append(crcs[i])
                else:
                    secret_table[k].append('')
            for k in secret_table.keys():
                secret_table[k].append(k)
            for subj_id, week_num, contents, idx in subjs:
                secret_table[ls_col].append(participants.loc[subj_id, 'Recently Seen'])
                secret_table['Email'].append(participants.loc[subj_id, 'E-mail'])
                secret_table['Name'].append(participants.loc[subj_id, 'Subject Name'])
                secret_table['Study'].append('PARIS')
                secret_table['Week Details'].append('Week {} - {}'.format(week_num, contents))
                secret_table['NoF'].append('follow-up')
                secret_table['ID'].append(subj_id)
                secret_table['Location'].append(participants.loc[subj_id, 'Blood / Saliva Location'])
                secret_table['Week Number'].append(week_num)
                secret_table['Day'].append('{}, {}'.format(days[idx], (first_day + datetime.timedelta(days=idx)).strftime("%m.%d.%y")))
                secret_table['Frequency'].append(participants.loc[subj_id, 'Schedule'])
                secret_table['Notes'].append(participants.loc[subj_id, 'Scheduling Notes'])
            for k in secret_table.keys():
                secret_table[k].append('')
        print(secret_table)
        pd.DataFrame.from_dict(secret_table).to_excel(writer, sheet_name='{}-{}'.format(first_day.strftime("%m.%d.%y"), (first_day + datetime.timedelta(days=4)).strftime("%m.%d.%y")), index=False)
    else:
        print(f"No data for val={val}") 
        
    writer.save()
    print("Written to", support_loc)
    writer.close()