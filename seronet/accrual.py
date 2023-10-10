import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import argparse
import util
import sys
import PySimpleGUI as sg
from seronet.d4_all_data import pull_from_source
from seronet.ecrabs import make_ecrabs
from seronet.clinical_forms import write_clinical
import os
from helpers import ValuesToClass


def lost_calculate(row):
    if row['Date'] != row['Last_Date']:
        return 'No'
    elif (datetime.date.today() - row['Date']).days < 300:
        return 'No'
    else:
        return 'Unknown'

def unscheduled_calculate(row):
    mit_days = np.array([0, 30, 60, 90, 180, 300, 540, 720])
    pri_months = [0, 6, 12]
    mit_cohorts = ['MARS', 'IRIS', 'TITAN']
    pri_cohorts = ['PRIORITY']
    if row['Primary_Cohort'] in mit_cohorts:
        if row['Days from Index'] <= 0 or np.abs(mit_days - row['Days from Index']).min() <= 14:
            return ['No', 'N/A']
        else:
            return ['Yes', 'Other']
    elif row['Primary_Cohort'] in pri_cohorts:
        pri_dates = [row['Baseline_Date'] + relativedelta(months=val) for val in pri_months]
        if np.abs([(pri_date - row['Date']).days for pri_date in pri_dates]).min() <= 14:
            return ['No', 'N/A']
        else:
            return ['Yes', 'Other']
    else:
        print(row['Cohort'], 'is not handled')
        print('Fatal Error. Exiting...')
        exit(1)

def yes_no(val):
    if val == 'N/A':
        return 'N/A'
    elif val > 0:
        return 'Yes'
    else:
        return 'No'

def accrue(args):
    intermediate = 'monthly_report'
    if not args.use_cache:
        all_data = pull_from_source(args.debug).query('Date <= @args.report_end').copy()
        ecrabs = make_ecrabs(all_data, output_fname=intermediate, debug=args.debug)
        dfs_clin = write_clinical(pd.DataFrame(ecrabs['Biospecimen']), 'monthly_tmp', debug=args.debug)
    else:
        all_data = pd.read_excel(util.unfiltered, keep_default_na=False).query('Date <= @args.report_end').copy()
        dfs_clin = pd.read_excel(util.cross_d4 + 'monthly_tmp.xlsx', sheet_name = None, keep_default_na=False)
    seronet_key = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name=None)
    exclusions = seronet_key['Exclusions']
    exclude_ppl = set(exclusions['Participant ID'].unique())
    exclude_ids = set(exclusions['Research_Participant_ID'].unique())
    ppl_filter = ~all_data['Participant ID'].isin(exclude_ppl)
    id_filter = ~all_data['Seronet ID'].isin(exclude_ids)
    all_data = all_data[ppl_filter & id_filter].copy()
    data_ppl = set(all_data['Seronet ID'].unique())
    ppl_cols = ['Research_Participant_ID', 'Age', 'Race', 'Ethnicity', 'Sex_At_Birth', 'Sunday_Prior_To_Visit_1']
    baseline_dates = all_data.drop_duplicates(subset='Seronet ID').set_index('Seronet ID').loc[:, 'Date']
    baseline_sundays = baseline_dates - np.mod(pd.to_datetime(baseline_dates).dt.weekday + 1, 7) * datetime.timedelta(days=1)
    ppl_data = (dfs_clin['Baseline']
                    .loc[:, ppl_cols[:-1]] # column subset
                    .query('Research_Participant_ID in @data_ppl') # row subset
                    .join(baseline_sundays, on='Research_Participant_ID') # add date info
                    .rename(columns={'Date': ppl_cols[-1]})) # rename date column
    output_outer = util.cross_d4 + 'Accrual/'
    output_inner = output_outer + '{}/'.format(datetime.date.today())

    vax_cols = ['Research_Participant_ID', 'Visit_Number', 'Vaccination_Status', 'SARS-CoV-2_Vaccine_Type', 'SARS-CoV-2_Vaccination_Date_Duration_From_Visit1']
    orig_date = 'SARS-CoV-2_Vaccination_Date_Duration_From_Index'
    vax_data = dfs_clin['Vax'].loc[:, vax_cols[:-1] + [orig_date]].query('Research_Participant_ID in @data_ppl').copy()
    vax_visits = vax_data.drop_duplicates(subset=['Research_Participant_ID', 'Visit_Number']).groupby('Research_Participant_ID').cumcount() + 1
    vax_visits.index = [(row['Research_Participant_ID'], row['Visit_Number']) for _, row in vax_data.drop_duplicates(subset=['Research_Participant_ID', 'Visit_Number']).iterrows()]
    vax_data['Visit_Number'] = vax_data.apply(lambda row: vax_visits[(row['Research_Participant_ID'], row['Visit_Number'])], axis=1)
    index_to_baseline = all_data.drop_duplicates(subset='Seronet ID').set_index('Seronet ID').loc[:, 'Days from Index']
    vax_data[vax_cols[-1]] = vax_data[orig_date].apply(lambda val: 0 if val == 'N/A' else val) - vax_data['Research_Participant_ID'].apply(lambda val: index_to_baseline[val])
    vax_data.loc[(vax_data[orig_date] == 'N/A'), vax_cols[-1]] = 'N/A'

    sample_cols = ['Site_Cohort_Name', 'Primary_Cohort', 'Research_Participant_ID', 'Visit_Number', 'Visit_Date_Duration_From_Visit_1', 'SARS_CoV_2_Infection_Status', 'Serum_Shipped_To_FNL', 'Serum_Volume_For_FNL', 'PBMC_Shipped_To_FNL', 'Num_PBMC_Vials_For_FNL', 'PBMC_Concentration', 'Unscheduled_Visit', 'Unscheduled_Visit_Purpose', 'Lost_To_FollowUp', 'Final_Visit', 'Collected_In_This_Reporting_Period']
    keep_cols = ['Seronet ID', 'Cohort', 'Days from Index', 'Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials', 'Sample ID', 'Date']
    col_map = {
        'Seronet ID': 'Research_Participant_ID',
        'Cohort': 'Site_Cohort_Name',
        'Volume of Serum Collected (mL)': 'Serum_Volume_For_FNL',
        'PBMC concentration per mL (x10^6)': 'PBMC_Concentration',
        '# of PBMC vials': 'Num_PBMC_Vials_For_FNL'
    }
    df_start = all_data.loc[:, keep_cols].rename(columns=col_map).query('Date <= @args.report_end').copy()
    cohort_key = {'M': 'Cancer', 'I': 'IBD', 'T': 'Transplant', 'P': 'Chronic Conditions', 'G': 'Healthy Control'}
    df_start['Primary_Cohort'] = df_start['Research_Participant_ID'].apply(lambda val: cohort_key[val[3:4]])
    df_start['Visit_Date_Duration_From_Visit_1'] = df_start['Days from Index'] - df_start['Research_Participant_ID'].apply(lambda val: index_to_baseline[val])
    df_start['Visit_Number'] = df_start.groupby('Research_Participant_ID').cumcount() + 1
    manifests = seronet_key['Aliquots Shipped'].assign(sample_id=lambda df: df['Sample ID'].astype(str))
    volcols = ['sample_id', 'Volume (mL)']
    pbmcs = manifests[manifests['Sample Type'] == 'PBMC'].loc[:, volcols].groupby('sample_id').sum()
    sera = manifests[manifests['Sample Type'] == 'Serum'].loc[:, volcols].groupby('sample_id').sum()
    df_start['Serum_Shipped_To_FNL'] = df_start['Sample ID'].apply(lambda val: min(sera.loc[val, 'Volume (mL)'], 4.5) if val in sera.index else 0)
    df_start['PBMC_Shipped_To_FNL'] = df_start['Sample ID'].apply(lambda val: pbmcs.loc[val, 'Volume (mL)'] if val in pbmcs.index else 0)
    df_start['Collected_In_This_Reporting_Period'] = (df_start['Date'] >= args.report_start).apply(lambda val: "Yes" if val else "No")
    cov_info = dfs_clin['COVID'].assign(sample_id=lambda df: df['Sample ID'].astype(str).str.strip().str.upper()).drop_duplicates(subset='sample_id', keep='last').set_index('sample_id')
    cov_map = {'Positive, Test Not Specified': 'Has Reported Infection',
               'Negative, Test Not Specified': 'Has Not Reported Infection',
               'Positive by PCR': 'Has Reported Infection',
               'Negative by PCR': 'Has Not Reported Infection',
               'Positive by Rapid Antigen Test': 'Has Reported Infection',
               'Negative by Rapid Antigen Test': 'Has Not Reported Infection',
               'Positive by Antibody Test': 'Has Reported Infection',
               'Negative by Antibody Test': 'Has Not Reported Infection',
               'Likely COVID Positive': 'Has Reported Infection',
               'No COVID event reported': 'Has Not Reported Infection',
               'No COVID data collected': 'Not Reported',
               '': 'Has Not Reported Infection'}
    cov_map = {k.upper(): v for k, v in cov_map.items()}
    df_start['SARS_CoV_2_Infection_Status'] = df_start['Sample ID'].astype(str).apply(lambda val: cov_map[cov_info.loc[val.strip().upper(), 'COVID_Status'].strip().upper()] if val.strip().upper() in cov_info.index else 'Not Reported')
    last_date = df_start.loc[:, ['Research_Participant_ID', 'Date']].groupby('Research_Participant_ID').max()
    baseline_date = df_start.loc[:, ['Research_Participant_ID', 'Date']].groupby('Research_Participant_ID').min()
    df_start['Last_Date'] = df_start['Research_Participant_ID'].apply(lambda val: last_date.loc[val, 'Date'])
    df_start['Baseline_Date'] = df_start['Research_Participant_ID'].apply(lambda val: baseline_date.loc[val, 'Date'])
    df_start['Lost_To_FollowUp'] = df_start.apply(lost_calculate, axis=1)
    df_start['Final_Visit'] = df_start['Lost_To_FollowUp']
    df_start['Unscheduled_Visit'] = 'No'
    df_start['Unscheduled_Visit_Purpose'] = 'N/A'
    blanks_filter = df_start['Serum_Volume_For_FNL'] == ''
    df_start.loc[blanks_filter, 'Serum_Volume_For_FNL'] = 0
    vol_filter = (df_start['Serum_Volume_For_FNL'] != df_start['Serum_Shipped_To_FNL']) & (df_start['Serum_Shipped_To_FNL'] != 0)
    df_start['Serum_Volume_For_FNL'] = df_start['Serum_Volume_For_FNL'] * ~vol_filter + df_start['Serum_Shipped_To_FNL'] * vol_filter
    df_start['Serum_Shipped_To_FNL'] = df_start['Serum_Shipped_To_FNL'].apply(yes_no)
    df_start.loc[df_start['Serum_Volume_For_FNL'] == 0, 'Serum_Shipped_To_FNL'] = 'N/A'
    df_start.loc[df_start['PBMC_Shipped_To_FNL'] > 0, 'Num_PBMC_Vials_For_FNL'] = df_start.loc[df_start['PBMC_Shipped_To_FNL'] > 0, 'PBMC_Shipped_To_FNL']
    df_start['PBMC_Shipped_To_FNL'] = df_start['PBMC_Shipped_To_FNL'].apply(yes_no)
    df_start['Num_PBMC_Vials_For_FNL'] = pd.to_numeric(df_start['Num_PBMC_Vials_For_FNL'], errors='coerce').fillna(0).astype(int)
    df_start.loc[df_start['Num_PBMC_Vials_For_FNL'] > 2, 'Num_PBMC_Vials_For_FNL'] = 2
    df_start.loc[df_start['Num_PBMC_Vials_For_FNL'] == 0, 'PBMC_Shipped_To_FNL'] = 'N/A'

    if not args.debug:
        if not os.path.exists(output_inner):
            os.makedirs(output_inner)
        with pd.ExcelFile(output_outer + 'GAEA_no_biospec.xlsx') as gaea_file:
            gaea_ppl = gaea_file.parse(sheet_name='participant_info')
            gaea_vax = gaea_file.parse(sheet_name='vaccination_status')
            gaea_samples = gaea_file.parse(sheet_name='visit_info')
        pd.concat([ppl_data, gaea_ppl]).to_excel(output_inner + 'Accrual_Participant_Info.xlsx', index=False, na_rep='N/A')
        pd.concat([vax_data.drop(orig_date, axis='columns'), gaea_vax]).to_excel(output_inner + 'Accrual_Vaccination_Status.xlsx', index=False, na_rep='N/A')
        pd.concat([df_start.loc[:, sample_cols], gaea_samples]).to_excel(output_inner + 'Accrual_Visit_Info.xlsx', index=False, na_rep='N/A')
        df_start.loc[:, sample_cols + ['Sample ID']].to_excel(output_outer + 'Latest_Accrual_SIDs.xlsx', index=False, na_rep='N/A')

if __name__ == '__main__':
    if len(sys.argv) != 1:
        
        argParser = argparse.ArgumentParser(description='Make files for monthly data submission.')
        argParser.add_argument('-c', '--use_cache', action='store_true')
        argParser.add_argument('-s', '--report_start', action='store', type=pd.to_datetime)
        argParser.add_argument('-e', '--report_end', action='store', type=pd.to_datetime)
        argParser.add_argument('-d', '--debug', action='store_true')
        args = argParser.parse_args()
    
    else:
        sg.theme('Dark Blue 17')

        layout = [[sg.Text('Accrual')],
                  [sg.Checkbox("Use Cache", key='use_cache', default=False), \
                    sg.Checkbox("debug", key='debug', default=False)],
                  [sg.Text('Start date'), sg.Input(key='report_start', default_text='1/1/2021'), sg.CalendarButton(button_text="choose date",close_when_date_chosen=True, target="report_start", format='%m/%d/%Y')],
                        [sg.Text('End date'), sg.Input(key='report_end', default_text='12/31/2025'), sg.CalendarButton(button_text="choose date",close_when_date_chosen=True, target="report_end", format='%m/%d/%Y')],
                    [sg.Submit(), sg.Cancel()]]

        window = sg.Window("Accrual Generation Script", layout)

        event, values = window.read()
        window.close()

        if event =='Cancel':
            quit()
        else:
            values['report_start'] = pd.to_datetime(values['report_start'])
            values['report_end'] = pd.to_datetime(values['report_start'])
            args = ValuesToClass(values)
        
        accrue(args)
