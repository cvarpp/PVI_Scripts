import pandas as pd
import numpy as np
import argparse
import util
from cam_convert import transform_cam
from helpers import query_intake, query_dscf, clean_sample_id, try_datediff, map_dates

def sufficient(val):
    try:
        return float(val) > 0
    except:
        return False

def seronet_annotate(df):
    df = df.copy()
    baseline_df = df.sort_values(by=['participant_id', 'Date Collected']).drop_duplicates(subset='participant_id').reset_index().set_index('participant_id')
    df['Seronet ID'] = df['participant_id'].apply(lambda val: "14_{}{}".format(baseline_df.loc[val, 'Cohort'][:1], baseline_df.loc[val, 'sample_id']))
    df['Index Date'] = pd.NA
    idx_third = df[(df['Cohort'] == 'TITAN')].index
    df.loc[idx_third, 'Index Date'] = df.loc[idx_third, '3rd Dose Date']
    df['Final Primary Date'] = df.apply(lambda row: row['1st Dose Date'] if str(row['Vaccine'])[:1].upper() == 'J' else row['2nd Dose Date'], axis=1)
    idx_vax = df[df['Cohort'].isin(['MARS', 'IRIS'])].index
    df.loc[idx_vax, 'Index Date'] = df.loc[idx_vax, 'Final Primary Date']
    unvax_index = df[df['Index Date'].isna()].index.intersection(idx_vax.union(idx_third))
    df.loc[unvax_index, 'Cohort'] = 'UNVAX'
    idx_base = df[df['Cohort'].isin(['PRIORITY', 'GAEA', 'UNVAX'])].index
    df['Baseline Date'] = df['participant_id'].apply(lambda val: baseline_df.loc[val, 'Date Collected'])
    df.loc[idx_base, 'Index Date'] = df.loc[idx_base, 'Baseline Date']
    return df

def cohort_data():
    iris_data = pd.read_excel(util.iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID']).assign(Cohort='IRIS', PID=lambda df: df['Participant ID'].str.upper().str.strip()).set_index('PID')
    titan_data = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).rename(columns={'Umbrella Corresponding Participant ID': 'Participant ID'}).dropna(subset=['Participant ID']).assign(Cohort='TITAN', PID=lambda df: df['Participant ID'].str.upper().str.strip()).set_index('PID')
    mars_data = pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID']).assign(Cohort='MARS', PID=lambda df: df['Participant ID'].str.upper().str.strip()).set_index('PID')
    gaea_data = pd.read_excel(util.gaea_folder + 'GAEA Tracker.xlsx', sheet_name='Summary').dropna(subset=['Participant ID']).assign(Cohort='GAEA', PID=lambda df: df['Participant ID'].str.upper().str.strip()).set_index('PID')
    participants = np.concatenate((mars_data.index.to_numpy(),
                                   titan_data.index.to_numpy(),
                                   iris_data.index.to_numpy(),
                                   gaea_data.index.to_numpy()))
    titan_convert = {v: k for k, v in {
        'Vaccine': 'Vaccine Type',
        '1st Dose Date': 'Vaccine #1 Date',
        '2nd Dose Date': 'Vaccine #2 Date',
        '3rd Dose Date': '3rd Dose Vaccine Date',
        '3rd Dose Vaccine': '3rd Dose Vaccine Type',
        'Boost Date': 'First Booster Vaccine Type',
        'Boost Vaccine': 'First Booster Dose Date',
        'Boost 2 Date': 'Second Booster Dose Date',
        'Boost 2 Vaccine': 'Second Booster Vaccine Type',
        'Baseline Date': 'Baseline date'
    }.items()}
    mars_convert = {v: k for k, v in {
        'Vaccine': 'Vaccine Name',
        '1st Dose Date': 'Vaccine #1 Date',
        '2nd Dose Date': 'Vaccine #2 Date',
        '3rd Dose Date': '3rd Vaccine',
        '3rd Dose Vaccine': '3rd Vaccine Type ',
        'Boost Date': '4th vaccine',
        'Boost Vaccine': '4th Vaccine Type',
        'Boost 2 Date': '5th vaccine',
        'Boost 2 Vaccine': '5th Vaccine Type',
        'Baseline Date': 'T1'
    }.items()}
    iris_convert = {v: k for k, v in {
        'Vaccine': 'Which Vaccine?',
        '1st Dose Date': 'First Dose Date',
        '2nd Dose Date': 'Second Dose Date',
        '3rd Dose Date': 'Third Dose Date',
        '3rd Dose Vaccine': 'Third Dose Type',
        'Boost Date': 'Fourth Dose Date',
        'Boost Vaccine': 'Fourth Dose Type',
        'Boost 2 Date': 'Fifth Dose Date',
        'Boost 2 Vaccine': 'Fifth Dose Type',
        'Baseline Date': 'Baseline Date'
    }.items()}
    gaea_convert = {v: k for k, v in {
        'Vaccine': 'Vaccine Type',
        '1st Dose Date': 'Dose #1 Date',
        '2nd Dose Date': 'Dose #2 Date',
        '3rd Dose Date': '3rd Vaccine Date',
        '3rd Dose Vaccine': '3rd Vaccine Type ',
        'Boost Date': '4th Vaccine Date',
        'Boost Vaccine': '4th Vaccine Type',
        'Boost 2 Date': '5th Vaccine Date',
        'Boost 2 Vaccine': '5th Vaccine Type',
        'Baseline Date': 'Baseline Date'
    }.items()}
    source_dfs = {'TITAN': titan_data, 'MARS': mars_data, 'IRIS': iris_data, 'GAEA': gaea_data}
    conversions = {'TITAN': titan_convert, 'MARS': mars_convert, 'IRIS': iris_convert, 'GAEA': gaea_convert}
    source_df = pd.concat([
        titan_data.rename(columns=titan_convert),
        mars_data.rename(columns=mars_convert),
        iris_data.rename(columns=iris_convert),
        gaea_data.rename(columns=gaea_convert)
    ]).loc[:, ['Cohort'] + [v for v in titan_convert.values()]]
    return source_df.pipe(map_dates, [col for col in source_df.columns if 'Date' in col])

def pull_from_source(debug=False):
    seronet_source = pd.ExcelFile(util.seronet_data + 'SERONET Key.xlsx')
    submitted_key = seronet_source.parse(sheet_name='Source').drop_duplicates(subset=['Participant ID']).set_index('Participant ID')
    sample_exclusions = seronet_source.parse(sheet_name='Sample Exclusions')
    exclude_samples = set(sample_exclusions.assign(sample_id=clean_sample_id)['sample_id'].unique())
    exclusions = seronet_source.parse(sheet_name='Exclusions')
    exclude_ppl = set(exclusions['Participant ID'].str.upper().str.strip().unique())

    source_df = cohort_data()
    participant_study = {rname: row['Cohort'] for rname, row in source_df.iterrows()}
    all_samples = query_intake(include_research=True)
    priority_ppl = all_samples[(all_samples['Study'] == 'PRIORITY') | (all_samples['participant_id'].str[:3] == 'MSH')]['participant_id'].unique()
    participant_study.update({pid: 'PRIORITY' for pid in priority_ppl})
    samples = all_samples[all_samples['participant_id'].isin(participant_study.keys()) & ~all_samples['participant_id'].isin(exclude_ppl)]
    samples = samples.loc[~samples.index.isin(exclude_samples), :].copy()
    samples_of_interest = samples.index.to_numpy()

    shared_samples = pd.read_excel(util.shared_samples, sheet_name='Released Samples')
    no_pbmcs = set([str(sid).strip().upper() for sid in shared_samples[shared_samples['Sample Type'] == 'PBMC']['Sample ID'].unique()])
    proc_unfiltered = query_dscf(sid_list=samples_of_interest, no_pbmcs=no_pbmcs)
    cam_df = transform_cam(debug=debug).drop_duplicates(subset='sample_id').set_index('sample_id')
    cam_df['coll_time'] = cam_df['Time Collected'].fillna(cam_df['Time'])
    cam_df['coll_inits'] = cam_df['Phlebotomist'].fillna('CCT (Clinical Care Team)')
    proc_unfiltered['coll_time'] = proc_unfiltered.apply(lambda row: cam_df.loc[row.name, 'coll_time'] if row.name in cam_df.index else row['Time Collected'], axis=1)
    proc_unfiltered['coll_inits'] = proc_unfiltered.apply(lambda row: cam_df.loc[row.name, 'coll_inits'] if row.name in cam_df.index else 'CCT (Clinical Care Team)', axis=1)
    serum_or_cells = proc_unfiltered['Volume of Serum Collected (mL)'].apply(sufficient) | proc_unfiltered['PBMC concentration per mL (x10^6)'].apply(sufficient)
    if not debug:
        proc_unfiltered[~serum_or_cells].to_excel(util.script_output + 'missing_info.xlsx', index=False)
    sample_info = proc_unfiltered[serum_or_cells].copy()
    sample_info.loc[sample_info['Volume of Serum Collected (mL)'] > 4.5, 'Volume of Serum Collected (mL)'] = 4.5

    proc_cols = ['Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials', 'coll_time', 'coll_inits', 'rec_time', 'proc_time', 'serum_freeze_time', 'cell_freeze_time', 'proc_inits', 'viability', 'cpt_vol', 'sst_vol', 'proc_comment']
    key_cols = ['Cohort', 'Vaccine', '1st Dose Date', '2nd Dose Date', '3rd Dose Date', '3rd Dose Vaccine', 'Boost Date', 'Boost Vaccine', 'Boost 2 Date', 'Boost 2 Vaccine', 'Baseline Date']

    cutoff_date = pd.to_datetime('2022-11-20').date()
    samples = samples.join(sample_info.loc[:, proc_cols], how='inner').join(source_df.loc[:, key_cols], on='participant_id')
    samples['Cohort'] = samples['Cohort'].fillna('PRIORITY')
    samples = samples[(samples['Cohort'] != 'IRIS') | (samples['Date Collected'] <= cutoff_date)].pipe(seronet_annotate)
    pid_replace = samples[samples['participant_id'].isin(submitted_key.index)].index
    samples['Seronet ID'] = samples.apply(lambda row: submitted_key.loc[row['participant_id'], 'Research_Participant_ID'] if row['participant_id'] in submitted_key.index else row['Seronet ID'], axis=1)

    columns = ['Cohort', 'Seronet ID', 'Index Date', 'Days from Index', 'Vaccine', '1st Dose Date', 'Days to 1st',
                '2nd Dose Date', 'Days to 2nd', '3rd Dose Date', 'Days to 3rd', 'Boost Vaccine', 'Boost Date', 'Days to Boost',
                'Boost 2 Vaccine', 'Boost 2 Date', 'Days to Boost 2', 'Participant ID', 'Date', 'Post-Baseline', 'Sample ID',
                'Visit Type', 'Qualitative', 'Quantitative', 'Spike endpoint', 'AUC', 'Log2AUC',
                'Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials', 'coll_inits',
                'coll_time', 'rec_time', 'proc_time', 'serum_freeze_time', 'cell_freeze_time', 'proc_inits', 'viability',
                'cpt_vol', 'sst_vol', 'proc_comment']

    days = ['Days from Index', 'Days to 1st', 'Days to 2nd', 'Days to 3rd', 'Days to Boost', 'Days to Boost 2', 'Post-Baseline']
    dates = ['Index Date', '1st Dose Date', '2nd Dose Date', '3rd Dose Date', 'Boost Date', 'Boost 2 Date', 'Baseline Date']
    for day_col, date_col in zip(days, dates):
        samples[day_col] = (samples['Date Collected'] - samples[date_col]).dt.days
    samples['Sample ID']=  samples.index.to_numpy()
    report = samples.rename(
        columns={'participant_id': 'Participant ID', 'Date Collected': 'Date', 'Visit Type / Samples Needed': 'Visit Type'}
    ).loc[:, columns].sort_values(by=['Participant ID', 'Date']).copy()
    if not debug:
        report.to_excel(util.unfiltered, index=False, freeze_panes=(1,3))
        print("Report written to {}".format(util.unfiltered))
    print("{} samples characterized.".format(report.shape[0]))
    return report

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Annotate SERONET samples with processing data')
    argParser.add_argument('-d', '--debug', action='store_true')
    args = argParser.parse_args()
    pull_from_source(args.debug)