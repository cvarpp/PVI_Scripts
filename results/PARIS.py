import pandas as pd
import numpy as np
from datetime import date
from dateutil import parser
import util
import argparse
import sys
import PySimpleGUI as sg
from helpers import try_datediff, permissive_datemax, query_intake, map_dates, immune_history, ValuesToClass

def auc_logger(val):
    if val == '-' or str(val).strip() == '' or len(str(val)) > 15:
        return '-'
    elif float(val) == 0.:
        return 0.
    else:
        return np.log2(float(val))

def paris_results():
    paris_data = pd.read_excel(util.paris + 'Patient Tracking - PARIS.xlsx', sheet_name='Subgroups', header=4).dropna(subset=['Participant ID'])
    paris_data['Participant ID'] = paris_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = paris_data['Participant ID'].unique()
    paris_data.set_index('Participant ID', inplace=True)
    dems = pd.read_excel(util.projects + 'PARIS/Demographics.xlsx').set_index('Subject ID')
    sample_info = query_intake(participants=participants, include_research=True)

    col_order = ['Participant ID', 'Date', 'Sample ID', 'Immune History', 'Days to 1st Vaccine Dose',
                'Days to Boost', 'Days to Last Infection', 'Days to Last Vax', 'Infection Timing',
                'Qualitative', 'Quantitative', 'Spike endpoint', 'AUC', 'Log2AUC', 'Log2Quant', 'Log2COV22', 'Vaccine Type',
                'Boost Type', 'Days to Infection 1', 'Days to Infection 2', 'Days to Infection 3', 'Days to Infection 4',
                'Infection Pre-Vaccine?', 'Number of SARS-CoV-2 Infections', 'Infection on Study',
                'First Dose Date', 'Second Dose Date', 'Days to 2nd', 'Boost Date', 'Boost 2 Date',
                'Days to Boost 2', 'Boost 2 Type', 'Boost 3 Date', 'Days to Boost 3', 'Boost 3 Type',
                'Boost 4 Date', 'Days to Boost 4', 'Boost 4 Type', 'Boost 5 Date', 'Days to Boost 5', 'Boost 5 Type',
                'Boost 6 Date', 'Days to Boost 6', 'Boost 6 Type','Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date', 
                'Most Recent Infection','Most Recent Vax', 'Post-Baseline', 'Visit Type', 'Gender', 'Age', 'Race', 'Ethnicity: Hispanic or Latino']
    dem_cols = ['Gender', 'Age', 'Race', 'Ethnicity: Hispanic or Latino']
    shared_cols = ['Infection Pre-Vaccine?', 'Vaccine Type', 'Number of SARS-CoV-2 Infections', 'Infection Timing', 'Boost Type', 'Boost 2 Type', 'Boost 3 Type', 'Boost 4 Type', 'Boost 5 Type', 'Boost 6 Type']
    date_cols = ['First Dose Date', 'Second Dose Date', 'Boost Date', 'Boost 2 Date', 'Boost 3 Date', 'Boost 4 Date', 'Boost 5 Date', 'Boost 6 Date', 'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    day_cols = ['Days to 1st Vaccine Dose', 'Days to 2nd', 'Days to Boost', 'Days to Boost 2', 'Days to Boost 3', 'Days to Boost 4', 'Days to Boost 5', 'Days to Boost 6', 'Days to Infection 1', 'Days to Infection 2', 'Days to Infection 3', 'Days to Infection 4']
    dose_dates = ['First Dose Date', 'Second Dose Date', 'Boost Date', 'Boost 2 Date', 'Boost 3 Date', 'Boost 4 Date', 'Boost 5 Date', 'Boost 6 Date']
    vaccine_type_cols = ['Vaccine Type', 'Vaccine Type', 'Boost Type', 'Boost 2 Type', 'Boost 3 Type', 'Boost 4 Type', 'Boost 5 Type', 'Boost 6 Type']
    inf_dates = ['Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    sample_info['Date'] = sample_info['Date Collected']
    sample_info = (sample_info.join(dems.loc[:, dem_cols], on='participant_id')
                              .join(paris_data.loc[:, shared_cols + date_cols], on='participant_id')
                              .reset_index().copy())
    baseline_data = sample_info.sort_values(by='Date').drop_duplicates(subset='participant_id').set_index('participant_id')
    sample_info['Post-Baseline'] = (pd.to_datetime(sample_info['Date']) - pd.to_datetime(sample_info['participant_id'].apply(lambda val: baseline_data.loc[val, 'Date']))).dt.days
    sample_info['Infection on Study'] = sample_info['participant_id'].apply(lambda ppt: 'yes' in (str(paris_data.loc[ppt, 'Infection 1 On Study?']) + str(paris_data.loc[ppt, 'Infection 2 On Study?']) + str(paris_data.loc[ppt, 'Infection 3 On Study?'])).lower())
    sample_info = sample_info.pipe(map_dates, date_cols)
    for date_col, day_col in zip(date_cols, day_cols):
        sample_info[day_col] = sample_info.apply(lambda row: try_datediff(row[date_col], row['Date']), axis=1)
    sample_info['Most Recent Infection'] = sample_info.apply(lambda row: permissive_datemax([row[inf_date] for inf_date in inf_dates], row['Date']), axis=1)
    sample_info['Most Recent Vax'] = sample_info.apply(lambda row: permissive_datemax([row[dose_date] for dose_date in dose_dates], row['Date']), axis=1)
    sample_info['Days to Last Infection'] = sample_info.apply(lambda row: try_datediff(row['Most Recent Infection'], row['Date']), axis=1)
    sample_info['Days to Last Vax'] = sample_info.apply(lambda row: try_datediff(row['Most Recent Vax'], row['Date']), axis=1)
    sample_info['Immune History'] = sample_info.apply(lambda row: immune_history(row[dose_dates], row[vaccine_type_cols], row[inf_dates], row['Date']), axis=1)
    sample_info['Participant ID'] = sample_info['participant_id']
    sample_info['Sample ID'] = sample_info['sample_id']
    sample_info['Visit Type'] = sample_info[util.visit_type]
    sample_info['Log2AUC'] = sample_info['AUC'].apply(auc_logger)
    (sample_info.loc[:, ['Participant ID', 'Date', 'Sample ID']]
           .drop_duplicates(subset=['Participant ID'], keep='last')
           .to_excel(util.paris + 'LastSeen.xlsx', index=False))
    return sample_info.loc[:, col_order]

if __name__ == '__main__':
    if len(sys.argv) != 1:

        argparser = argparse.ArgumentParser(description='Generate report for all PARIS samples')
        argparser.add_argument('-o', '--output_file', action='store', default='tmp', help="Prefix for the output file (current date appended")
        argparser.add_argument('-d', '--debug', action='store_true', help="Print to the command line but do not write to file")
        args = argparser.parse_args()

    else:
        sg.theme('Green')

        layout = [[sg.Text('PARIS Result Generation Script')],
                  [sg.Checkbox('Debug?', key='debug', default=False)],
                  [sg.Text('Output File Name:'),sg.Input(key='output_file')],
                  [sg.Submit(), sg.Cancel()]]
        
        window = sg.Window('MARS Results Script', layout)

        event,  values = window.read()
        window.close()

        if event=='Cancel':
            quit()
        else:
            args = ValuesToClass(values)

    report = paris_results()
    if not args.debug:
        output_filename = util.paris + 'datasets/{}_{}.xlsx'.format(args.output_file, date.today().strftime("%m.%d.%y"))
        report.to_excel(output_filename, index=False)
        print("PARIS report written to {}".format(output_filename))
    print("{} samples from {} participants".format(
        report.shape[0],
        report['Participant ID'].unique().size
    ))