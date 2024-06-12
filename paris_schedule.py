import pandas as pd
import numpy as np
import datetime
import sys
import util

def day_difference(first_date, baseline_datetime):
    if type(baseline_datetime) == datetime.time:
        print("Please make sure every participant has a baseline visit date in the tracker!")
        print("If they have not yet had their baseline visit, please disregard.")
        return datetime.timedelta(days=-1)
    else:
        return first_date - baseline_datetime.date()

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("usage: python monthly_schedule.py MM-DD-YYYY")
        exit(1)
    else:
        try:
            start_monday = pd.to_datetime(sys.argv[1]).date() # TODO: Add check verifying that this is a Monday
        except:
            print("usage: python monthly_schedule.py MM-DD-YYYY")
            exit(1)
    participants = pd.read_excel(util.paris_tracker, sheet_name='Participant details')
    last_visit = pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=util.header_intake).drop_duplicates(subset='Participant ID', keep='last').set_index('Participant ID')
    participants.set_index('Subject ID', inplace=True)
    participants['Recently Seen'] = participants.apply(lambda row: last_visit.loc[row.name, 'Date Collected'] if row.name in last_visit.index.to_numpy() else datetime.date.today(), axis=1)
    completed = participants[participants['Study Status'] == 'Completed']
    participants = participants[participants['Study Status'] == 'Active']
    latest_week = 48 # This is actually the earliest visit we are worrying about (Week 48) since everyone has reached it
    visit_week_numbers = range(latest_week, latest_week + 300, 4) # goes until ~week 300, further than we need to worry about
    visit_week_numbers = [x for x in visit_week_numbers if x <= 208]
    visit_day_offsets = [datetime.timedelta(days=int(x) * 7) for x in visit_week_numbers]
    contents = ['blood (5), saliva'] * 70 # Only looking at up to 70 visits (which is 6 years, seems fine)
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    emails = ['morgan.vankesteren@mssm.edu', 'jacob.mauldin@mssm.edu']
    crcs = ['Jessica', 'Jake']
    max_days = 28
    crc_schedule = {}
    invalid_emails = set()
    for week in range(0, max_days, 7):
        first_day_in_week = start_monday + datetime.timedelta(days=week)
        crc_schedule[week] = [[] for _ in range(len(crcs))]
        days_from_baseline = participants['Baseline / Day 0 / Week 0'].apply(lambda baseline: day_difference(first_day_in_week, baseline))
        for week_num, day_diff, content in zip(visit_week_numbers, visit_day_offsets, contents):
            for subj_id in participants.index:
                delta = day_diff - days_from_baseline.loc[subj_id] # difference between participant day offset and day offset for current visit in loop
                for weekday in range(5): #rename idx weekday for sake of understanding
                    try:
                        if delta == datetime.timedelta(days=weekday) and (participants.loc[subj_id, 'Schedule'] == '4 weeks'):
                            crc_email = participants.loc[subj_id, 'CRC Email']
                            if crc_email in emails:
                                crc_schedule[week][emails.index(crc_email)].append((subj_id, week_num, content, weekday))
                            else:
                                invalid_emails.add(crc_email)
                    except Exception as err:
                        print("Errored at", subj_id, "with time difference", delta, "for reason", err)
                        exit(1)

    print("Invalid CRC emails encountered:", invalid_emails)
    support_loc = util.clin_ops + 'PARIS/' + 'Scheduling Support PARIS 1 Month from {}.xlsx'.format(start_monday)
    ls_col = 'Last Seen as of {}'.format(datetime.date.today())
    to_write = []
    for week in range(0, max_days, 7):
        secret_table = {ls_col: [], 'Email': [], 'Name': [], 'Study': [], 'Week Details': [], 'NoF': [], 'ID': [], 'Location': [], 'Week Number': [], 'Day': [], 'Frequency': [], 'Notes': [] }
        first_day = start_monday + datetime.timedelta(days=week)
        for i, subjs in enumerate(crc_schedule[week]):
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
        sheet_name = '{}-{}'.format(first_day.strftime("%m.%d.%y"), (first_day + datetime.timedelta(days=4)).strftime("%m.%d.%y"))
        to_write.append((sheet_name,
                         pd.DataFrame.from_dict(secret_table)))
    with pd.ExcelWriter(support_loc) as writer:
        completed.loc[:, ['Recently Seen']].rename(columns={'Recently Seen': ls_col}).to_excel(writer, sheet_name='May Need Reconsent')
        for sname, df in to_write:
            df.to_excel(writer, sheet_name=sname, index=False)
    print("Written to", support_loc)