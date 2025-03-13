import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import argparse
import util
import sys
import PySimpleGUI as sg
from seronet.accrual import accrue
import os
from helpers import ValuesToClass
import warnings

def baseline_sunday(visit_num, date):
    if 'Baseline' in str(visit_num):
        day_of_week = date.weekday()
        sunday_prior = date - pd.Timedelta(days=(day_of_week+1))
        return sunday_prior
    else:
        return ('N/A')
    
def days_from_vax(idxtocollection, idxtovaccination):
    if pd.isna(idxtovaccination):
        return 'Pre-Vaccine'
    else:
        return idxtocollection-idxtovaccination

def in_window(key, rpid, days):
    tps = [30, 60, 90, 120, 180, 360]
    if days == 'Pre-Vaccine':
        if rpid in key['Research_Participant_ID'].unique():
            return 'No'
        else:
            return 'Yes'
    else:
        for tp in tps:
            if days >= tp-11 and days <= tp +11:
                return 'Yes'
        else:
            return 'No'        

def make_intake(df_accrual, ecrabs, dfs_clin, seronet_key):
    collection_month = args.report_end.strftime('%B')
    output_folder = util.cross_d4 + 'Data/2025 Data Submissions/{} Intake Set'.format(collection_month) + os.sep
    #Calculate List of Samples to Include
    accrued_in_period = df_accrual[df_accrual['Collected_In_This_Reporting_Period'] == 'Yes']
    samples_in_period = accrued_in_period['Sample ID']
    intake_df = ecrabs['Biospecimen'].set_index('Biospecimen_ID')
    aliquot_df = ecrabs['Aliquot'].drop_duplicates(subset='Biospecimen_ID').set_index('Biospecimen_ID')
    vax_df = dfs_clin['Vax']
    kp2_df = vax_df[vax_df['Vaccination_Status'].str.contains('KP.2|JN.1')].set_index('Research_Participant_ID')
    snet_key = seronet_key['Source']
    current_key = snet_key[snet_key['Submission'].str.contains('Intake')]
    key_no_current_month = current_key[~current_key['Submission'].str.contains(collection_month)]
    intake_df = (intake_df.join(aliquot_df, rsuffix='_1')
                          .join(kp2_df, rsuffix='_1', on ='Research_Participant_ID'))
    intake_df = intake_df[intake_df['Sample ID'].isin(samples_in_period)]
    intake_df['Days from KP.2'] = intake_df.apply(lambda row: days_from_vax(row['Biospecimen_Collection_Date_Duration_From_Index'], row['SARS-CoV-2_Vaccination_Date_Duration_From_Index']), axis=1)
    intake_df['In Window?'] = intake_df.apply(lambda row: in_window(key_no_current_month, row['Research_Participant_ID'], row['Days from KP.2']), axis=1)
    #Create Add to Intake Sheet
    intake_df.rename(columns={'Visit_Number':'Visit_ID', 'Aliquot_Volume':'Total Volume/Sample Yield'}, inplace=True)
    intake_df['Sunday_Prior_To_Visit_1'] = intake_df.apply(lambda row: baseline_sunday(row['Visit_ID'], row['Date']), axis=1)
    intake_df['Purpose of Visit'] = 'Post-Vaccine'
    intake_col_order = ['Research_Participant_ID', 'Cohort', 'Visit_ID', 'Biospecimen_ID', 'Biospecimen_Type', 'Initial_Volume_of_Biospecimen',
                        'Total Volume/Sample Yield', 'Collection_Tube_Type', 'Live_Cells_Hemocytometer_Count', 'Total_Cells_Hemocytometer_Count',
                        'Viability_Hemocytometer_Count', 'Live_Cells_Automated_Count', 'Total_Cells_Automated_Count', 'Viability_Automated_Count',
                        'Sunday_Prior_To_Visit_1', 'Purpose of Visit']
    in_window_df = intake_df[intake_df['In Window?'] == 'Yes'].copy()
    add_to_intake = in_window_df.reset_index().loc[:, intake_col_order]
    #Create Add to OOW Sheet
    oow_col_order = ['Biospecimen_ID', 'Sample ID', 'Days from KP.2', 'Note']
    oow_df = intake_df[intake_df['In Window?'] == 'No'].copy()
    oow_df['Note'] = oow_df['Days from KP.2'].astype(str).str.contains('Pre-Vaccine').apply(lambda val: 'Second pre-vax' if val else '')
    add_to_oow = oow_df.reset_index().loc[:, oow_col_order]
    #Create Add to Key Sheet
    source_col_order = ['Participant ID', 'Sample ID', 'Biospecimen_ID', 'Research_Participant_ID', 'Submission']
    intake_df['Submission'] = str(collection_month) + '_Intake'
    add_to_source = intake_df.reset_index().loc[:, source_col_order]
    #Filter ecrabs and dfs_clin based on sids in add to intake
    filter_vals = in_window_df['Sample ID'].astype(str).unique()
    suffix = collection_month
    if len(suffix) > 0:
        suffix = '_' + suffix
    ecrabs.update({'Baseline':dfs_clin['Baseline'], 'FollowUp':dfs_clin['FollowUp']})    
    biospec_sheets = ecrabs
    clinical_sheets = {'COVID':dfs_clin['COVID'], 'Vax':dfs_clin['Vax'], 'Meds':dfs_clin['Meds'], 'Cancer':dfs_clin['Cancer']}
    with pd.ExcelWriter(output_folder + 'monthly_processing_filtered{}.xlsx'.format(suffix)) as writer:
            for sname, sheet in biospec_sheets.items():
                try:
                    sheet = sheet[sheet['Sample ID'].astype(str).isin(filter_vals)]
                    sheet.to_excel(writer, sheet_name=sname, index=False, na_rep='N/A')
                except:
                    print(sname, "not included")
                    continue
    with pd.ExcelWriter(output_folder + 'monthly_clinical_filtered{}.xlsx'.format(suffix)) as writer:
            for sname, sheet in clinical_sheets.items():
                try:
                    sheet = sheet[sheet['Sample ID'].astype(str).isin(filter_vals)]
                    sheet.to_excel(writer, sheet_name=sname, index=False, na_rep='N/A')
                except:
                    print(sname, "not included")
                    continue
    with pd.ExcelWriter(output_folder + 'intake.xlsx') as writer:
            add_to_intake.to_excel(writer, sheet_name = 'Intake', index=False, na_rep='N/A')
            add_to_oow.to_excel(writer, sheet_name = 'SNet Key OOW', index=False, na_rep='N/A')
            add_to_source.to_excel(writer, sheet_name = 'SNet Key Source', index=False, na_rep='N/A')
    print('Written to', output_folder, 'intake.xlsx')
    return add_to_intake

if __name__ == '__main__':
    if len(sys.argv) != 1:
        argParser = argparse.ArgumentParser(description='Make files for monthly data submission.')
        argParser.add_argument('-s', '--report_start', action='store', type=pd.to_datetime)
        argParser.add_argument('-e', '--report_end', action='store', type=pd.to_datetime)
        argParser.add_argument('-d', '--debug', action='store_true')
        args = argParser.parse_args()
    else:
        sg.theme('Dark Purple 6')

        layout = [[sg.Text('Intake')],
                  [sg.Checkbox("Debug", key='debug', default=False)],
                  [sg.Text('Start date'), sg.Input(key='report_start', default_text='1/25/2025'), sg.CalendarButton(button_text="choose date",close_when_date_chosen=True, target="report_start", format='%m/%d/%Y')],
                        [sg.Text('End date'), sg.Input(key='report_end', default_text='2/22/2025'), sg.CalendarButton(button_text="choose date",close_when_date_chosen=True, target="report_end", format='%m/%d/%Y')],
                    [sg.Submit(), sg.Cancel()]]

        window = sg.Window("Intake Generation Script", layout)

        event, values = window.read()
        window.close()

        if event =='Cancel':
            quit()
        else:
            values['report_start'] = pd.to_datetime(values['report_start'])
            values['report_end'] = pd.to_datetime(values['report_end'])
            args = ValuesToClass(values)
    df_accrual, ecrabs, dfs_clin, seronet_key = accrue(args)
    make_intake(df_accrual, ecrabs, dfs_clin, seronet_key)