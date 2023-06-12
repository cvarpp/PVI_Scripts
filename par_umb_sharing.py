from email.policy import default
import numpy as np
import pandas as pd
import util
from datetime import date
import datetime
from dateutil import parser
from helpers import query_intake
import argparse

def not_shared(val):
    return pd.isna(val) or val == 'No' or val == ''

def make_report(use_cache=False):
    paris_info = pd.read_excel(util.paris_tracker, sheet_name='Main', header=8)
    umbrella_info = pd.read_excel(util.umbrella_tracker, sheet_name='Summary', header=0)
    was_shared_col = 'Clinical Ab Result Shared?'

    samples = query_intake(use_cache=use_cache)
    share_filter = samples[was_shared_col].apply(not_shared)
    result_filter = (samples['Qualitative'] != '') & ~samples['Qualitative'].apply(pd.isna)
    samplesClean = samples[share_filter & result_filter].copy()
    samplesCleanAll = samplesClean.copy()
    older_results = []
    pemail = paris_info[['Subject ID', 'E-mail']].rename(columns={'E-mail': 'Email'})
    uemail = umbrella_info[['Subject ID', 'Email']]
    emails = (pd.concat([pemail, uemail])
                .assign(pid=lambda df: df['Subject ID'].str.strip())
                .drop_duplicates(subset='pid').set_index('pid'))

    participants = samplesClean['participant_id'].unique()
    participant_results = {participant: [] for participant in participants}
    for sid, sample in samplesClean.iterrows():
        participant = sample['participant_id']
        if pd.isna(sample[was_shared_col]) or sample[was_shared_col] == "No":
            if str(sample['Qualitative']).strip().upper()[:2] == "NE":
                sample['Qualitative'] = "Negative"
                sample['Quantitative'] = "Negative"
            try:
                participant_results[participant].append((pd.to_datetime(sample['Date Collected']).date(), sid, sample[util.visit_type], sample['Qualitative'], sample['Quantitative']))
            except:
                print("Sample", sid, "improperly recorded as collected on", sample['Date Collected'])
                print("Not included in result sharing")
    data = {'Participant ID': [], 'Sample ID': [], 'Date': [], 'Email': [], 'Visit Type': [], 'Qualitative': [], 'Quantitative': []}
    for participant, results in participant_results.items():
        for result in results:
            date_collected, sample_id, util.visit_type, result_stat, result_value = result
            data['Participant ID'].append(participant)
            data['Sample ID'].append(sample_id)
            data['Date'].append(date_collected)
            data['Visit Type'].append(util.visit_type)
            data['Qualitative'].append(result_stat)
            data['Quantitative'].append(result_value)
            if participant in emails.index:
                email = emails.loc[participant, 'Email']
            else:
                email = ''
            data['Email'].append(email)
    report = pd.DataFrame(data)
    return report

if __name__ == '__main__':
    argparser = argparse.ArgumentParser()
    argparser.add_argument('-d', '--debug', action='store_true')
    argparser.add_argument('-c', '--use_cache', action='store_true')
    args = argparser.parse_args()

    report = make_report(args.use_cache)
    last_60_days = date.today() - datetime.timedelta(days=60)
    old_report = report[report['Date'] < last_60_days]
    report_filtered = report[report['Date'] >= last_60_days]
    output_filename = util.sharing + 'result_reporting_test_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    if not args.debug:
        with pd.ExcelWriter(output_filename) as writer:
            report_filtered.to_excel(writer, sheet_name='Needs Results - Recent', index=False)
            old_report.to_excel(writer, sheet_name='Needs Results - Old(60d+)', index=False)
        print("Report written to {}".format(output_filename))
    print(report.shape[0], old_report.shape[0], report_filtered.shape[0])
