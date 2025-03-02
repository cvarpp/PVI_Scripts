import numpy as np
import pandas as pd
import util
from helpers import query_intake, ValuesToClass
import argparse
import os
import sys
import PySimpleGUI as sg

def not_shared(val):
    return pd.isna(val) or val == 'No' or val == ''

def make_report(use_cache=False):
    paris_info = pd.read_excel(util.paris_tracker, sheet_name='Main', header=8)
    umbrella_info = pd.read_excel(util.umbrella_tracker, sheet_name='Summary', header=0)
    titan_info = pd.read_excel(util.titan_tracker, sheet_name='Tracker', header=4)
    apollo_info = pd.read_excel(util.apollo_tracker, sheet_name='Summary')
    hd_info = pd.read_excel(util.clin_ops + 'Healthy donors/Healthy Donors Participant Tracker.xlsx', sheet_name = 'Participants')

    was_shared_col = 'Clinical Ab Result Shared?'

    samples = query_intake(use_cache=use_cache, include_research=True)
    share_filter = samples[was_shared_col].apply(not_shared)
    old_result_filter = ~samples['Qualitative'].apply(pd.isna)
    new_result_filter = ~samples['COV22'].apply(pd.isna) | (samples['COV22_str'].str.len() > 0)
    samplesClean = samples[share_filter & (old_result_filter | new_result_filter)].copy()
    samplesCleanAll = samplesClean.copy()
    older_results = []

    pemail = paris_info[['Subject ID', 'E-mail']].rename(columns={'E-mail': 'Email'})
    uemail = umbrella_info[['Subject ID', 'Email']]
    temail = titan_info[['Umbrella Corresponding Participant ID', 'Email (From EPIC)']].rename(columns={'Umbrella Corresponding Participant ID': 'Subject ID', 'Email (From EPIC)': 'Email'})
    aemail = apollo_info[['Participant ID', 'Email']].rename(columns={'Participant ID': 'Subject ID'})
    hdemail = hd_info[['Participant ID', 'Email']].rename(columns={'Participant ID': 'Subject ID'})
    lk_sheet = util.cross_project + 'Lock & Key/Lock and Key - KDS.xlsx'
    if os.path.exists(lk_sheet):
        lkemail = pd.read_excel(lk_sheet, sheet_name='Link L&K').rename(columns={'Participant ID': 'Subject ID'})[['Subject ID', 'Email']]
    else:
        lkemail = uemail
    emails = (pd.concat([lkemail, pemail, temail, uemail, aemail, hdemail])
                .assign(pid=lambda df: df['Subject ID'].str.strip())
                .drop_duplicates(subset='pid').set_index('pid'))

    keep_cols = ['participant_id', 'sample_id', 'Date Collected', util.visit_type, 'Email', 'Qualitative', 'Quant_str', 'COV22_str', 'Quantitative', 'COV22', 'Spike endpoint', 'AUC']
    report = samplesClean.join(emails, on='participant_id').reset_index().loc[:, keep_cols]
    valid_prefixes = ['16791', '03374', '23873', '16772']
    report = report[report['participant_id'].str[:5].isin(valid_prefixes)].copy()
    report['COV22 / Quant'] = np.exp2(np.log2(report['COV22']) - np.log2(report['Quantitative']))
    report['COV22 / Research'] = np.exp2(np.log2(report['COV22']) - np.log2(report['AUC']))
    return report

if __name__ == '__main__':
    if len(sys.argv) != 1:
        argparser = argparse.ArgumentParser()
        argparser.add_argument('-d', '--debug', action='store_true')
        argparser.add_argument('-c', '--use_cache', action='store_true')
        argparser.add_argument('-r', '--recency', type=int, default=60, help='Number of days in the past to consider recent')
        args = argparser.parse_args()
    else:
        sg.theme('Dark Blue 17')

        layout = [[sg.Text('Result Sharing Script')],
                  [sg.Checkbox("Use Cache", key='use_cache', default=False), \
                    sg.Checkbox("Debug?", key='debug', default=False)],
                    [sg.Text('How Recent?') ,sg.Input(key="recency", default_text="60")],
                    [sg.Submit(), sg.Cancel()]]

        window = sg.Window("Result Sharing Script", layout)

        event, values = window.read()
        window.close()

        if event =='Cancel':
            quit()
        else:
            values['recency'] = int(values['recency'])
            args = ValuesToClass(values)        

    report = make_report(args.use_cache)
    recency_cutoff = pd.Timestamp.now() - pd.Timedelta(days=args.recency)
    report_old = report[report['Date Collected'] < recency_cutoff]
    report_new = report[report['Date Collected'] >= recency_cutoff]
    report_new_emails = report_new[~report_new['Email'].isna()]
    report_new_no_emails = report_new[report_new['Email'].isna()]

    assert(report_old.shape[0] + report_new.shape[0] == report.shape[0])
    assert(report_new_emails.shape[0] + report_new_no_emails.shape[0] == report_new.shape[0])
    output_filename = util.sharing + 'result_reporting_{}.xlsx'.format(pd.Timestamp.now().strftime("%m.%d.%y"))
    if not args.debug:
        with pd.ExcelWriter(output_filename) as writer:
            report_new_emails.to_excel(writer, sheet_name='Results to Share', index=False)
            report_new_no_emails.to_excel(writer, sheet_name='Recent Unshared - Email Missing', index=False)
            report_old.to_excel(writer, sheet_name='Older Unshared Results', index=False)
        print("Report written to {}".format(output_filename))
    print("Total to report:", report.shape[0], "Older count:", report_old.shape[0], "Recent Results to Share:", report_new_emails.shape[0], "Recent Results Missing Email:", report_new_no_emails.shape[0])
