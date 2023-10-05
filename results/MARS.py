import pandas as pd
import numpy as np
from bisect import bisect_left
import argparse
import util
import sys
import PySimpleGUI as sg
from helpers import query_dscf, query_research, query_intake, clean_sample_id, try_datediff, ValuesToClass

def lookup_query(row, source, latest_date, coi):
    doi = row['Date Collected']
    if doi > latest_date:
        doi = latest_date
    return source.loc[(doi, row['participant_id']), coi]

def fallible(row, date_lkp):
    try:
        output = bisect_left(date_lkp[row['participant_id']], pd.to_datetime(row['Date Collected']).date())
    except:
        print("Failed on {} at {}".format(row['participant_id'], row.name))
        print(date_lkp[row['participant_id']])
        print(row['Date Collected'])
        exit(1)
    return output

def mars_report(args):
    mars_data = pd.read_excel(util.mars_folder + 'MARS for D4 Long.xlsx', sheet_name=None)
    ppl_info = mars_data['Baseline Info']
    exclusions = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name='Exclusions')
    exclude_ppl = set(exclusions['Participant ID'].unique())
    mars_ppl = set(ppl_info['Participant ID'].str.strip().unique()) - exclude_ppl
    intake = query_intake(participants=mars_ppl)
    soi = set(intake.index.to_numpy())
    tcell_data = pd.read_excel(util.proc + 'T Cell Experiments.xlsx', sheet_name=None)
    tcell_samples = tcell_data['Data'].set_index('Name').assign(sample_id=clean_sample_id)
    tcell_samples.loc[tcell_samples['IU/mL'] == '> 10¶', 'IU/mL'] = 20
    # tcell_ids = set(tcell_samples['sample_id'].unique())
    tcell_wide = tcell_samples.pivot_table(values="IU/mL", index='sample_id', columns="Peptide \npreparation")
    for col in tcell_wide.columns:
        tcell_wide.loc[tcell_wide[col] == 20, col] = '>10'
    # tcell_results = tcell_data['IU.mL']
    paired_data = pd.read_excel(util.project_ws + 'Morgan/T Cell Experiments/MARS Heparin Tube Collection.xlsx', sheet_name=None)
    paired_df = paired_data['Pre,Post Biv Boost']
    paired_ppl = set(paired_df['Participant ID'].str.strip().to_numpy())
    research_results = query_research(sid_list=soi)
    proc_cols = ['Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', 'viability', '# of PBMC vials', 'cpt_vol', 'sst_vol', 'proc_comment']
    sample_info = query_dscf(sid_list=soi).loc[:, proc_cols].copy()
    shared_samples = pd.read_excel(util.shared_samples, sheet_name='Released Samples').assign(sample_id=clean_sample_id)
    shared_info = shared_samples.pivot_table(values='# Aliquots', columns='Sample Type', index='sample_id', aggfunc='sum').fillna(0)
    vax_data = mars_data['COVID Vaccinations']
    vax_date_lkp = {}
    vax_dose_lkp = {}
    vax_type_lkp = {}
    for pid in mars_ppl:
        vax_date_lkp[pid] = []
        vax_dose_lkp[pid] = []
        vax_type_lkp[pid] = []
        df = vax_data[vax_data['Participant ID'] == pid].sort_values(by='SARS-CoV-2_Vaccination_Date')
        vax_date_lkp[pid].extend(df['SARS-CoV-2_Vaccination_Date'].dt.date.to_numpy())
        vax_dose_lkp[pid].extend(df['Vaccination_Status'].to_numpy())
        vax_type_lkp[pid].extend(df['SARS-CoV-2_Vaccine_Type'].to_numpy())
    vax_dates = vax_data.set_index(['Participant ID', 'Vaccination_Status']).loc[:, ['SARS-CoV-2_Vaccination_Date']].unstack().droplevel(0, axis='columns').applymap(lambda val: val.date())
    vax_dates['Dose 1'] = pd.to_datetime(vax_dates['Dose 1 of 1'].fillna(' ').astype(str) + vax_dates['Dose 1 of 2'].fillna(' ').astype(str), errors='coerce').dt.date # needs to be a space (' ') because of fillna bug since pandas 1.3
    vax_dates['Final Primary Dose'] = pd.to_datetime(vax_dates['Dose 1 of 1'].fillna(' ').astype(str) + vax_dates['Dose 2 of 2'].fillna(' ').astype(str), errors='coerce').dt.date
    vax_types = vax_data.set_index(['Participant ID', 'Vaccination_Status']).loc[:, ['SARS-CoV-2_Vaccine_Type']].unstack().droplevel(0, axis='columns')
    vax_types['Dose 1'] = vax_types['Dose 1 of 1'].fillna('').astype(str) + vax_types['Dose 1 of 2'].fillna('').astype(str)
    vax_types['Final Primary Dose'] = vax_types['Dose 1 of 1'].fillna('').astype(str) + vax_types['Dose 2 of 2'].fillna('').astype(str)

    cov_data = mars_data['COVID Infections']
    cov_date_lkp = {}
    cov_test_lkp = {}
    for pid in mars_ppl:
        cov_date_lkp[pid] = []
        cov_test_lkp[pid] = []
        df = cov_data[cov_data['Participant ID'] == pid].sort_values(by='Report_Time')
        cov_date_lkp[pid].extend(df['Report_Time'].dt.date.to_numpy())
        cov_test_lkp[pid].extend(df['COVID_Status'].to_numpy())

    cancer_cols = ['Cancer', 'Year_Of_Diagnosis', 'In_Remission', 'Response Status', 'Chemotherapy']
    cancer_spec = mars_data['Cancer-specific'].set_index('Participant ID').loc[:, cancer_cols]

    dem_cols = ['Age', 'Sex_At_Birth', 'Race', 'Ethnicity', 'BMI']
    source_df = (intake.join(research_results)
                       .join(ppl_info.set_index('Participant ID').loc[:, dem_cols], on='participant_id')
                       .join(cancer_spec, on='participant_id')
                       .join(sample_info)
                       .join(shared_info.rename(columns=lambda col: col + ' Shared'))
                       .join(vax_dates.rename(columns=lambda col: col + ' Date'), on='participant_id')
                       .join(vax_types.rename(columns=lambda col: col + ' Type'), on='participant_id')
                       .join(tcell_wide)
                       .sort_values(by=['participant_id', 'Date Collected']))
    source_df['Paired T Cell Data Available'] = source_df['participant_id'].isin(paired_ppl).apply(lambda val: "Yes" if val else "No")
    source_df['Vax Count Pre-Sample'] = source_df.apply(lambda row: fallible(row, vax_date_lkp), axis=1)
    for pid in mars_ppl:
        vax_date_lkp[pid].append(pd.NA)
        vax_dose_lkp[pid].append('Unvaccinated')
        vax_type_lkp[pid].append(pd.NA)
    source_df['Most Recent Vaccine Date'] = source_df.apply(lambda row: vax_date_lkp[row['participant_id']][row['Vax Count Pre-Sample'] - 1], axis=1)
    source_df['Most Recent Vaccine Type'] = source_df.apply(lambda row: vax_type_lkp[row['participant_id']][row['Vax Count Pre-Sample'] - 1], axis=1)
    source_df['Most Recent Vaccine Dose'] = source_df.apply(lambda row: vax_dose_lkp[row['participant_id']][row['Vax Count Pre-Sample'] - 1], axis=1)
    source_df['Days to Most Recent Vaccine'] = source_df.apply(lambda row: try_datediff(row['Most Recent Vaccine Date'], row['Date Collected']), axis=1)

    day_cols = []
    for col in vax_dates.columns:
        day_cols.append('Days to ' + col)
        source_df['Days to ' + col] = source_df.apply(lambda row: try_datediff(row[col + ' Date'], row['Date Collected']), axis=1)

    source_df['COVID Count Pre-Sample'] = source_df.apply(lambda row: fallible(row, cov_date_lkp), axis=1)
    for pid in mars_ppl:
        cov_date_lkp[pid].append(pd.NA)
        cov_test_lkp[pid].append(pd.NA)
    source_df['Most Recent Infection Date'] = source_df.apply(lambda row: cov_date_lkp[row['participant_id']][row['COVID Count Pre-Sample'] - 1], axis=1)
    source_df['Most Recent Infection Test'] = source_df.apply(lambda row: cov_test_lkp[row['participant_id']][row['COVID Count Pre-Sample'] - 1], axis=1)
    source_df['Days to Most Recent Infection'] = source_df.apply(lambda row: try_datediff(row['Most Recent Infection Date'], row['Date Collected']), axis=1)
    tcell_cols = tcell_wide.columns.to_list()
    ab_cols = ['Qualitative', 'Quantitative', 'COV22'] + research_results.columns.to_list()
    output_tcell_wide_cols = ['Date Collected', 'participant_id'] + tcell_cols + ab_cols + dem_cols
    output_tcell_wide = source_df.loc[:, output_tcell_wide_cols].dropna(subset=tcell_cols)
    output_tcell_wide_paired = output_tcell_wide[output_tcell_wide['participant_id'].isin(paired_ppl)].sort_values(by=['participant_id', 'Date Collected'])
    output_tcell_long = tcell_samples.loc[:, ['IU/mL', 'Tr #', 'Peptide \npreparation', 'sample_id']].join(output_tcell_wide.drop(tcell_cols, axis='columns'), on='sample_id', how='inner')
    vax_df_cols = ['Date Collected', 'participant_id'] + ab_cols + ['Days to Most Recent Vaccine', 'Vax Count Pre-Sample', 'COVID Count Pre-Sample', 'Most Recent Vaccine Date', 'Most Recent Vaccine Type', 'Most Recent Vaccine Dose', 'Most Recent Infection Date', 'Most Recent Infection Test', 'Days to Most Recent Infection', 'Days to Dose 1', 'Days to Final Primary Dose', 'Days to Dose 3'] + dem_cols + ['Cancer', 'Chemotherapy']
    vax_dfs = {}
    for dose_type in source_df['Most Recent Vaccine Dose'].unique():
        df = source_df[(source_df['Most Recent Vaccine Dose'] == dose_type)].copy()
        if dose_type == 'Unvaccinated':
            df['Clean'] = df.apply(lambda row: pd.isna(row['Most Recent Infection Date']), axis=1)
        else:
            df['Clean'] = df.apply(lambda row: pd.isna(row['Most Recent Infection Date']) or (row['Most Recent Vaccine Date'] > row['Most Recent Infection Date']), axis=1)
        vax_dfs[dose_type] = df.loc[:, vax_df_cols[:7] + ['Clean'] + vax_df_cols[7:]]
    abs_long = source_df.loc[:, vax_df_cols]
    if not args.debug:
        with pd.ExcelWriter(util.mars_folder + 'MARS Central Info.xlsx') as writer:
            source_df.to_excel(writer, sheet_name='Source')
            abs_long.to_excel(writer, sheet_name='Long')
            output_tcell_wide.to_excel(writer, sheet_name='TCell Wide')
            output_tcell_wide_paired.to_excel(writer, sheet_name='TCell Wide Paired')
            for dose_type in vax_dfs.keys():
                vax_dfs[dose_type].to_excel(writer, sheet_name=dose_type.replace(':', '_'))
            output_tcell_long.to_excel(writer, 'TCell Long')
    return source_df

if __name__ == '__main__':
    
    if len(sys.argv) != 1:
        argparser = argparse.ArgumentParser(description='Generate report for all MARS samples')
        argparser.add_argument('-o', '--output_file', action='store', default='tmp', help="Prefix for the output file (in addition to current date)")
        argparser.add_argument('-d', '--debug', action='store_true', help="Print to the command line but do not write to file")
        args = argparser.parse_args()

    else:
        sg.theme('Dark Blue 17')

        layout = [[sg.Text('MARS Result Generation Script')],
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

    source_df = mars_report(args)
    print("{} samples from {} participants.".format(
        source_df.shape[0],
        source_df['participant_id'].unique().size
    ))
