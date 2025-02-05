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

def apollo_results():
    apollo_data = pd.read_excel(util.apollo_tracker, sheet_name='Summary').dropna(subset=['Participant ID'])
    apollo_data['Participant ID'] = apollo_data['Participant ID'].apply(lambda val: val.strip().upper())
    apollo_vaccines = pd.read_excel(util.apollo_tracker, sheet_name='Vaccinations').dropna(subset=['Participant ID'])
    apollo_infections = pd.read_excel(util.apollo_tracker, sheet_name='Infections').dropna(subset=['Participant ID'])
    participants = apollo_data['Participant ID'].unique()
    participants_vax = apollo_vaccines['Participant ID'].unique()
    participants_inf = apollo_infections['Participant ID'].unique()
    apollo_data.set_index('Participant ID', inplace=True)
    apollo_vaccines.set_index('Participant ID', inplace=True)
    apollo_infections.set_index('Participant ID', inplace=True)
    sample_info = query_intake(participants=participants, include_research=True)

    schedule = ['2/22/24', '5/13/24', '8/20/24', '12/2/24', '3/3/25']
    def get_quarter(date):
       for quarter in range(0,4):
          if pd.to_datetime(date) >= pd.to_datetime(schedule[quarter]) and pd.to_datetime(date) < pd.to_datetime(schedule[quarter+1]):
             return quarter+1
          else:
             quarter+1
    immune_event_cols = ['Dose #1 Date', 'Dose #2 Date', 'Dose #3 Date', 'Dose #4 Date', 'Dose #5 Date', 
                     'Dose #6 Date', 'Dose #7 Date', 'Dose #8 Date', 'Dose #9 Date', 
                     'Symptom Onset Date 1', 'Symptom Onset Date 2', 'Symptom Onset Date 3', 
                     'Symptom Onset Date 4']
    vaccine_date_cols = ['Dose #1 Date', 'Dose #2 Date', 'Dose #3 Date', 'Dose #4 Date', 'Dose #5 Date', 
                     'Dose #6 Date', 'Dose #7 Date', 'Dose #8 Date', 'Dose #9 Date']
    vaccine_type_cols = ['Dose #1 Type', 'Dose #2 Type', 'Dose #3 Type', 'Dose #4 Type', 'Dose #5 Type', 
                     'Dose #6 Type', 'Dose #7 Type', 'Dose #8 Type', 'Dose #9 Type']
    infection_date_cols = ['Symptom Onset Date 1', 'Symptom Onset Date 2', 'Symptom Onset Date 3', 
                     'Symptom Onset Date 4']
    sample_info = (sample_info.join(apollo_data, on='participant_id')
                              .join(apollo_vaccines, on='participant_id', how = 'inner')
                              .join(apollo_infections, on='participant_id', how = 'inner')
                              .reset_index().copy())
    sample_info.rename(columns={'sample_id': 'Sample ID', 'participant_id':'Participant ID', 'Enrollment Wave':'Quarter Enrolled', 
                                'Spike endpoint':'Spike Endpoint', 'Age @ Enrollment':'Age', 'Visit Type / Samples Needed':'Visit Type'}, inplace=True)
    sample_info = sample_info.pipe(map_dates, immune_event_cols)
    sample_info['Quarter Collected'] = sample_info.apply(lambda row: get_quarter(row['Date Collected']), axis=1)
    sample_info['Date of Last Immune Event'] = sample_info.apply(lambda row: permissive_datemax([row[immune_event] for immune_event in immune_event_cols], row['Date Collected']), axis=1)
    sample_info['Days from Last Immune Event'] = sample_info.apply(lambda row: try_datediff(row['Date of Last Immune Event'], row['Date Collected']), axis=1)
    sample_info['Immune History'] = sample_info.apply(lambda row: immune_history(row[vaccine_date_cols], row[vaccine_type_cols], row[infection_date_cols], row['Date Collected']), axis=1)
    sample_info['Total Immune Events'] = sample_info['Immune History'].str.count('-') + 1
    sample_info['Total Infections'] = sample_info['Immune History'].str.count('I')
    sample_info['Total Vaccine Doses'] = sample_info['Total Immune Events'].astype(int) - sample_info['Total Infections'].astype(int)
    sample_info['Last Immune Event'] = sample_info['Immune History'].str.endswith('I').apply(lambda val: 'Infection' if val else 'Vaccination')
    sample_info['Last Immune Event Bin'] = sample_info['Days from Last Immune Event'].apply(bin_event)
    col_order = ['Sample ID', 'Participant ID', 'Date Collected', 'Visit Type', 'Quarter Collected', 'Quarter Enrolled', 'Last Immune Event', 
                 'Date of Last Immune Event', 'Days from Last Immune Event', 'Last Immune Event Bin', 
                 'Immune History', 'Total Immune Events', 'Total Vaccine Doses', 'Total Infections', 
                 'COV22 Results', 'Spike Endpoint', 'AUC', 'Age', 'Sex', 'Gender', 'Race', 
                 'Ethnicity']
    no_vax_info = set(participants) - set(participants_vax)
    if len(no_vax_info) != 0:
        print("Vaccine info not found for", no_vax_info)
    no_inf_info = set(participants) - set(participants_inf)
    if len(no_inf_info) != 0:
        print("Infection info not found for", no_inf_info)
    return sample_info.reset_index().loc[:, col_order]

if __name__ == '__main__':
    if len(sys.argv) != 1:

        argparser = argparse.ArgumentParser(description='Generate report for all APOLLO samples')
        argparser.add_argument('-d', '--debug', action='store_true', help="Print to the command line but do not write to file")
        argparser.add_argument('-n', '--nogui', action='store_true', help="Don't pull up the GUI")
        args = argparser.parse_args()
    else:
        sg.theme('Light Brown 5')

        layout = [[sg.Text('APOLLO Result Generation Script')],
                  [sg.Checkbox('Debug?', key='debug', default=False)],
                  [sg.Submit(), sg.Cancel()]]
        
        window = sg.Window('APOLLO Results Script', layout)

        event,  values = window.read()
        window.close()

        if event=='Cancel':
            quit()
        else:
            args = ValuesToClass(values)
    report = apollo_results()
    if not args.debug:
        output_filename = util.apollo_folder + 'APOLLO Sampling Dashboard_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
        report.to_excel(output_filename, index=False)
        print("Written to {}".format(output_filename))
    print("{} samples from {} participants".format(
        report.shape[0],
        report['Participant ID'].unique().size
    ))