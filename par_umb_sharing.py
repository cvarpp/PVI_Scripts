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
    old_result_filter = ~samples['Qualitative'].apply(pd.isna)
    new_result_filter = ~samples['COV22'].apply(pd.isna)
    samplesClean = samples[share_filter & (old_result_filter | new_result_filter)].copy()
    samplesCleanAll = samplesClean.copy()
    older_results = []
    pemail = paris_info[['Subject ID', 'E-mail']].rename(columns={'E-mail': 'Email'})
    uemail = umbrella_info[['Subject ID', 'Email']]
    emails = (pd.concat([pemail, uemail])
                .assign(pid=lambda df: df['Subject ID'].str.strip())
                .drop_duplicates(subset='pid').set_index('pid'))

    keep_cols = ['participant_id', 'sample_id', 'Date Collected', util.visit_type, 'Email', 'Qualitative', 'Quantitative', 'COV22']
    report = samplesClean.join(emails, on='participant_id').reset_index().loc[:, keep_cols]
    return report

if __name__ == '__main__':
    argparser = argparse.ArgumentParser()
    argparser.add_argument('-d', '--debug', action='store_true')
    argparser.add_argument('-c', '--use_cache', action='store_true')
    argparser.add_argument('-r', '--recency', type=int, default=60, help='Number of days in the past to consider recent')
    args = argparser.parse_args()

    report = make_report(args.use_cache)
    recency_cutoff = date.today() - datetime.timedelta(days=args.recency)
    report_old = report[report['Date Collected'] < recency_cutoff]
    report_new = report[report['Date Collected'] >= recency_cutoff]
    report_new_emails = report_new[~report_new['Email'].isna()]
    report_new_no_emails = report_new[report_new['Email'].isna()]
    assert(report_old.shape[0] + report_new.shape[0] == report.shape[0])
    assert(report_new_emails.shape[0] + report_new_no_emails.shape[0] == report_new.shape[0])
    output_filename = util.sharing + 'result_reporting_test_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    if not args.debug:
        with pd.ExcelWriter(output_filename) as writer:
            report_new_emails.to_excel(writer, sheet_name='Results to Share', index=False)
            report_new_no_emails.to_excel(writer, sheet_name='Recent Unshared Results - Email Missing', index=False)
            report_old.to_excel(writer, sheet_name='Older Unshared Results', index=False)
        print("Report written to {}".format(output_filename))
    print("Total to report:", report.shape[0], "Older count:", report_old.shape[0], "Recent Results to Share:", report_new_emails.shape[0], "Recent Results Missing Email:", report_new_no_emails.shape[0])
