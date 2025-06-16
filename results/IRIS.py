import pandas as pd
import numpy as np
from datetime import date
from dateutil import parser
import util
import argparse
import sys
import PySimpleGUI as sg
from helpers import try_datediff, permissive_datemax, query_intake, map_dates, immune_history, ValuesToClass

def bin_event(days):
  if days > 400:
    return ">400"
  elif days > 150:
    bin = 50 * ((days - 1) // 50)
    return f"{bin+1:g}-{bin+50:g}"
  else:
    bin = 30 * ((days - 1) // 30)
    return f"{bin+1:g}-{bin+30:g}"  

def iris_results():
    iris_data = pd.read_excel(util.iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID'])
    iris_data['Participant ID'] = iris_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = iris_data['Participant ID'].unique()
    iris_data.set_index('Participant ID', inplace=True)
    sample_info = query_intake(participants=participants, include_research=True)

    immune_event_cols = ['First Dose Date', 'Second Dose Date', 'Third Dose Date', 'Fourth Dose Date', 'Fifth Dose Date',
                         'Sixth Dose Date', 'Seventh Dose Date', 'Eighth Dose Date', 'Ninth Dose Date',
                         'Symptom Onset', 'Symptom Onset 2', 'Symptom Onset 3']
    vaccine_date_cols = ['First Dose Date', 'Second Dose Date', 'Third Dose Date', 'Fourth Dose Date', 'Fifth Dose Date',
                         'Sixth Dose Date', 'Seventh Dose Date', 'Eighth Dose Date', 'Ninth Dose Date']
    vaccine_type_cols = ['Which Vaccine?', 'Which Vaccine?', 'Third Dose Type', 'Fourth Dose Type', 'Fifth Dose Type',
                         'Sixth Dose Type', 'Seventh Dose Type', 'Eighth Dose Type', 'Ninth Dose Type']
    infection_date_cols = ['Symptom Onset', 'Symptom Onset 2', 'Symptom Onset 3']
    sample_info = (sample_info.join(iris_data, on='participant_id', rsuffix='1')
                              .reset_index().copy())
    sample_info.rename(columns={'sample_id': 'Sample ID', 'participant_id':'Participant ID', 'IBD: CD vs UC':'Condition',
                                'Spike endpoint':'Spike Endpoint', 'Visit Type / Samples Needed':'Visit Type'}, inplace=True)
    sample_info = sample_info.pipe(map_dates, immune_event_cols)
    sample_info['Date of Last Immune Event'] = sample_info.apply(lambda row: permissive_datemax([row[immune_event] for immune_event in immune_event_cols], row['Date Collected']), axis=1)
    sample_info['Days from Last Immune Event'] = sample_info.apply(lambda row: try_datediff(row['Date of Last Immune Event'], row['Date Collected']), axis=1)
    sample_info['Immune History'] = sample_info.apply(lambda row: immune_history(row[vaccine_date_cols], row[vaccine_type_cols], row[infection_date_cols], row['Date Collected']), axis=1)
    sample_info['Total Immune Events'] = sample_info['Immune History'].str.count('-') + 1
    sample_info['Total Infections'] = sample_info['Immune History'].str.count('I')
    sample_info['Total Vaccine Doses'] = sample_info['Total Immune Events'].astype(int) - sample_info['Total Infections'].astype(int)
    sample_info['Last Immune Event'] = sample_info['Immune History'].str.endswith('I').apply(lambda val: 'Infection' if val else 'Vaccination')
    sample_info['Last Immune Event Bin'] = sample_info['Days from Last Immune Event'].apply(bin_event)
    col_order = ['Sample ID', 'Participant ID', 'Date Collected', 'Visit Type', 'Condition', 'Last Immune Event', 
                 'Date of Last Immune Event', 'Days from Last Immune Event', 'Last Immune Event Bin', 
                 'Immune History', 'Total Immune Events', 'Total Vaccine Doses', 'Total Infections', 
                 'COV22 Results', 'Spike Endpoint', 'AUC', 'Age', 'Sex', 'Race', 
                 'Ethnicity']
    return sample_info.reset_index().loc[:, col_order]

def iris_pts():
    iris_data = pd.read_excel(util.iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header = 4).dropna(subset=['Participant ID'])
    iris_data['Participant ID'] = iris_data['Participant ID'].apply(lambda val: val.strip().upper())
    iris_data.set_index('Participant ID', inplace=True)

    immune_event_cols = ['First Dose Date', 'Second Dose Date', 'Third Dose Date', 'Fourth Dose Date', 'Fifth Dose Date',
                         'Sixth Dose Date', 'Seventh Dose Date', 'Eighth Dose Date', 'Ninth Dose Date',
                         'Symptom Onset', 'Symptom Onset 2', 'Symptom Onset 3']
    vaccine_date_cols = ['First Dose Date', 'Second Dose Date', 'Third Dose Date', 'Fourth Dose Date', 'Fifth Dose Date',
                         'Sixth Dose Date', 'Seventh Dose Date', 'Eighth Dose Date', 'Ninth Dose Date']
    vaccine_type_cols = ['Which Vaccine?', 'Which Vaccine?', 'Third Dose Type', 'Fourth Dose Type', 'Fifth Dose Type',
                         'Sixth Dose Type', 'Seventh Dose Type', 'Eighth Dose Type', 'Ninth Dose Type']
    infection_date_cols = ['Symptom Onset', 'Symptom Onset 2', 'Symptom Onset 3']
    pt_info = iris_data.reset_index().copy()
    pt_info.rename(columns={'participant_id':'Participant ID', 'IBD: CD vs UC':'Condition'}, inplace=True)
    pt_info = pt_info.pipe(map_dates, immune_event_cols)
    pt_info['Immune History'] = pt_info.apply(lambda row: immune_history(row[vaccine_date_cols], row[vaccine_type_cols], row[infection_date_cols], date.today()), axis=1)
    pt_info['Total Immune Events'] = pt_info['Immune History'].str.count('-') + 1
    pt_info['Total Infections'] = pt_info['Immune History'].str.count('I')
    pt_info['Total Vaccine Doses'] = pt_info['Total Immune Events'].astype(int) - pt_info['Total Infections'].astype(int)
    pt_info['Last Immune Event'] = pt_info['Immune History'].str.endswith('I').apply(lambda val: 'Infection' if val else 'Vaccination')
    col_order = ['Participant ID', 'Condition', 'Immune History', 'Total Immune Events', 'Total Vaccine Doses', 
                 'Total Infections', 'Age', 'Sex', 'Race', 'Ethnicity','First Dose Date', 'Second Dose Date', 'Third Dose Date', 'Fourth Dose Date', 'Fifth Dose Date',
                 'Sixth Dose Date', 'Seventh Dose Date', 'Eighth Dose Date', 'Ninth Dose Date',
                 'Symptom Onset', 'Symptom Onset 2', 'Symptom Onset 3']
    return pt_info.reset_index().loc[:, col_order].sort_values(by=['Participant ID'])

if __name__ == '__main__':
    if len(sys.argv) != 1:

        argparser = argparse.ArgumentParser(description='Generate report for all IRIS samples')
        argparser.add_argument('-d', '--debug', action='store_true', help="Print to the command line but do not write to file")
        argparser.add_argument('-n', '--nogui', action='store_true', help="Don't pull up the GUI")
        args = argparser.parse_args()
    else:
        sg.theme('Light Brown 5')

        layout = [[sg.Text('IRIS Result Generation Script')],
                  [sg.Checkbox('Debug?', key='debug', default=False)],
                  [sg.Submit(), sg.Cancel()]]
        
        window = sg.Window('IRIS Results Script', layout)

        event,  values = window.read()
        window.close()

        if event=='Cancel':
            quit()
        else:
            args = ValuesToClass(values)
    samples_report = iris_results()
    pts_report = iris_pts()
    if not args.debug:
         with pd.ExcelWriter(util.iris_folder + 'IRIS Sampling Dashboard {}.xlsx'.format(date.today().strftime("%m.%d.%y"))) as writer:
            samples_report.to_excel(writer, index=False, sheet_name = 'Samples')
            pts_report.to_excel(writer, index=False, sheet_name = 'Participants')
            print("Written to {}".format(writer))
    print("{} samples from {} participants".format(
        samples_report.shape[0],
        samples_report['Participant ID'].unique().size
    ))
