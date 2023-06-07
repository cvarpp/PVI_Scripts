import pandas as pd
import numpy as np
import argparse
import util
from helpers import query_dscf, query_research

def sufficient(val):
    try:
        return float(val) > 0
    except:
        return False

def pull_from_source(debug=False):
    iris_data = pd.read_excel(util.iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID'])
    titan_data = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).rename(columns={'Umbrella Corresponding Participant ID': 'Participant ID'}).dropna(subset=['Participant ID'])
    mars_data = pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID'])
    gaea_data = pd.read_excel(util.gaea_folder + 'GAEA Tracker.xlsx', sheet_name='Summary').dropna(subset=['Participant ID'])
    for df in [iris_data, titan_data, mars_data, gaea_data]:
        df['Participant ID'] = df['Participant ID'].apply(lambda val: val.strip().upper())
    participants = np.concatenate((mars_data['Participant ID'].unique(),
                                   titan_data['Participant ID'].unique(),
                                   iris_data['Participant ID'].unique(),
                                   gaea_data['Participant ID'].unique()))
    participant_study = {p: 'MARS' for p in mars_data['Participant ID'].unique()}
    participant_study.update({p: 'TITAN' for p in titan_data['Participant ID'].unique()})
    participant_study.update({p: 'IRIS' for p in iris_data['Participant ID'].unique()})
    participant_study.update({p: 'GAEA' for p in gaea_data['Participant ID'].unique()})
    mars_data.set_index('Participant ID', inplace=True)
    iris_data.set_index('Participant ID', inplace=True)
    titan_data.set_index('Participant ID', inplace=True)
    gaea_data.set_index('Participant ID', inplace=True)
    titan_convert = {
        'Vaccine': 'Vaccine Type',
        '1st Dose Date': 'Vaccine #1 Date',
        '2nd Dose Date': 'Vaccine #2 Date',
        '3rd Dose Date': '3rd Dose Vaccine Date',
        '3rd Dose Vaccine': '3rd Dose Vaccine Type',
        'Boost Date': 'First Booster Vaccine Type',
        'Boost Vaccine': 'First Booster Dose Date',
        'Boost 2 Date': 'Second Booster Vaccine Type',
        'Boost 2 Vaccine': 'Second Booster Dose Date',
        'Baseline Date': 'Baseline date'
    }
    mars_convert = {
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
    }
    iris_convert = {
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
    }
    gaea_convert = {
        'Vaccine': 'Vaccine Type',
        '1st Dose Date': 'Dose #1 Date',
        '2nd Dose Date': 'Dose #2 Date',
        '3rd Dose Date': '3rd Vaccine Date',
        '3rd Dose Vaccine': '3rd Vaccine Type ',
        'Boost Date': '4th Vaccine Date',
        'Boost Vaccine': '4th Vaccine Type',
        'Boost 2 Date': '5th Vaccine Type',
        'Boost 2 Vaccine': '5th Vaccine Type',
        'Baseline Date': 'Baseline Date'
    }
    # mars_data['3rd Dose Date'] = ''
    # mars_data['3rd Dose Vaccine'] = ''
    # gaea_data['3rd Dose Date'] = ''
    # gaea_data['3rd Dose Vaccine'] = ''
    # iris_data['3rd Dose Date'] = ''
    # iris_data['3rd Dose Vaccine'] = ''
    source_dfs = {'TITAN': titan_data, 'MARS': mars_data, 'IRIS': iris_data, 'GAEA': gaea_data}
    conversions = {'TITAN': titan_convert, 'MARS': mars_convert, 'IRIS': iris_convert, 'GAEA': gaea_convert}
    intake_source = pd.ExcelFile(util.intake)
    collection_log = intake_source.parse(sheet_name='Collection Log').assign(sample_id=lambda df: df['Sample ID'].astype(str).str.strip().str.upper()).drop('Sample ID', axis=1).set_index('sample_id')
    samples = intake_source.parse(sheet_name='Sample Intake Log', header=util.header_intake, dtype=str)
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    samplesClean = samples.dropna(subset=['Participant ID'])
    cutoff_date = pd.to_datetime('2022-11-20').date()
    participant_samples = {participant: [] for participant in participants}
    submitted_key = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name='Source').drop_duplicates(subset=['Participant ID']).set_index('Participant ID')
    sample_exclusions = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name='Sample Exclusions')
    exclude_samples = set(sample_exclusions['Sample ID'].astype(str).str.upper().str.strip().unique())
    samples_of_interest = set()
    for _, sample in samplesClean.iterrows():
        sample_id = str(sample['Sample ID']).strip().upper()
        if len(sample_id) != 5 or sample_id in exclude_samples:
            continue
        participant = str(sample['Participant ID']).strip().upper()
        if str(sample['Study']).strip().upper() == 'PRIORITY':
            if participant not in participant_samples.keys():
                participant_samples[participant] = []
                participant_study[participant] = 'PRIORITY'
        if participant in participant_samples.keys():
            if str(sample[newCol]).strip().upper() == "NEGATIVE":
                sample[newCol2] = "Negative"
            if pd.isna(sample[newCol2]):
                result_new = '-'
            elif type(sample[newCol2]) == int:
                result_new = sample[newCol2]
            elif str(sample[newCol2]).strip().upper() == "NEGATIVE":
                result_new = 1.
            else:
                result_new = sample[newCol2]
            try:
                sample_date = pd.to_datetime(str(sample['Date Collected']))
            except:
                print(sample_id, "has invalid date", sample['Date Collected'])
                sample_date = pd.to_datetime('1/1/1900')
            participant_samples[participant].append((sample_date, str(sample_id).strip().upper(), sample[visit_type], sample[newCol], result_new, sample['Blood Collector Initials']))
            samples_of_interest.add(str(sample_id).strip().upper())

    

    research_results = query_research(samples_of_interest)

    shared_samples = pd.read_excel(util.shared_samples, sheet_name='Released Samples')
    no_pbmcs = set([str(sid).strip().upper() for sid in shared_samples[shared_samples['Sample Type'] == 'PBMC']['Sample ID'].unique()])
    all_samples = query_dscf(sid_list=samples_of_interest, no_pbmcs=no_pbmcs)
    all_samples['coll_time'] = all_samples.apply(lambda row: collection_log.loc[row.name, 'Time Collected'] if row.name in collection_log.index else row['Time Collected'], axis=1)
    serum_or_cells = all_samples['Volume of Serum Collected (mL)'].apply(sufficient) | all_samples['PBMC concentration per mL (x10^6)'].apply(sufficient)
    all_samples[~serum_or_cells].to_excel(util.script_output + 'missing_info.xlsx', index=False)
    sample_info = all_samples[serum_or_cells].copy()
    sample_info.loc[sample_info['Volume of Serum Collected (mL)'] > 4.5, 'Volume of Serum Collected (mL)'] = 4.5

    proc_cols = ['Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials', 'coll_time', 'rec_time', 'proc_time', 'serum_freeze_time', 'cell_freeze_time', 'proc_inits', 'viability', 'cpt_vol', 'sst_vol', 'proc_comment']

    columns = ['Cohort', 'Seronet ID', 'Index Date', 'Days from Index', 'Vaccine', '1st Dose Date', 'Days to 1st', '2nd Dose Date', 'Days to 2nd', '3rd Dose Date', 'Days to 3rd', 'Boost Vaccine', 'Boost Date', 'Days to Boost', 'Boost 2 Vaccine', 'Boost 2 Date', 'Days to Boost 2', 'Participant ID', 'Date', 'Post-Baseline', 'Sample ID', 'Visit Type', 'Qualitative', 'Quantitative', 'Spike endpoint', 'AUC', 'Log2AUC', 'Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials', 'coll_inits', 'coll_time', 'rec_time', 'proc_time', 'serum_freeze_time', 'cell_freeze_time', 'proc_inits', 'viability', 'cpt_vol', 'sst_vol', 'proc_comment']

    data = {col: [] for col in columns}
    participant_data = {}
    exclusions = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name='Exclusions')
    exclude_ppl = set(exclusions['Participant ID'].unique())
    
    for participant, samples in participant_samples.items():
        if participant in exclude_ppl:
            continue
        try:
            samples.sort(key=lambda x: x[0])
        except:
            print(participant, "has samples that won't sort (see below)")
            print(samples)
            print()
            continue
        if len(samples) < 1:
            print(participant, "has no samples")
            print()
            continue
        baseline = samples[0][0]
        participant_data[participant] = {}
        cohort = participant_study[participant].strip().upper()
        seronet_id = "14_{}{}".format(cohort[:1], samples[0][1])
        key_cols = ['Vaccine', '1st Dose Date', '2nd Dose Date', '3rd Dose Date', '3rd Dose Vaccine', 'Boost Date', 'Boost Vaccine', 'Boost 2 Date', 'Boost 2 Vaccine', 'Baseline Date']
        if cohort in conversions.keys():
            source_df = source_dfs[cohort]
            converter = conversions[cohort]
            for k in key_cols:
                if 'Date' in k:
                    try:
                        participant_data[participant][k] = pd.to_datetime(source_df.loc[participant, converter[k]])
                    except:
                        participant_data[participant][k] = source_df.loc[participant, converter[k]]
                else:
                    participant_data[participant][k] = source_df.loc[participant, converter[k]]
            if cohort == 'TITAN':
                participant_data[participant]['Index Date'] = participant_data[participant]['3rd Dose Date']
            elif cohort == 'GAEA':
                participant_data[participant]['Index Date'] = participant_data[participant]['Baseline Date']
            elif str(participant_data[participant]['Vaccine'])[:1].upper() == 'J':
                participant_data[participant]['Index Date'] = participant_data[participant]['1st Dose Date']
            else:
                participant_data[participant]['Index Date'] = participant_data[participant]['2nd Dose Date']
            if type(participant_data[participant]['Index Date']) != pd.Timestamp or pd.isna(participant_data[participant]['Index Date']):
                participant_data[participant]['Index Date'] = baseline
        elif participant_study[participant] == 'PRIORITY':
            for k in key_cols:
                participant_data[participant][k] = ''
            participant_data[participant]['Index Date'] = baseline
        else:
            print(participant, "is not in the study! They are in", participant_study[participant])
            exit(1)
        if participant in submitted_key.index:
            seronet_id = submitted_key.loc[participant, 'Research_Participant_ID']

        for date_, sample_id, visit_type, result, result_new, coll_inits in samples:
            sample_id = str(sample_id).strip().upper()
            if cohort == 'IRIS' :
                if date_ > cutoff_date:
                    continue
            if sample_id not in sample_info.index:
                continue
            for col in proc_cols:
                data[col].append(sample_info.loc[sample_id, col])
            data['Cohort'].append(participant_study[participant])
            data['Index Date'].append(participant_data[participant]['Index Date'])
            try:
                data['Days from Index'].append(int((date_ - data['Index Date'][-1]).days))
            except:
                data['Days from Index'].append('')
            data['Vaccine'].append(participant_data[participant]['Vaccine'])
            data['1st Dose Date'].append(participant_data[participant]['1st Dose Date'])
            try:
                data['Days to 1st'].append(int((date_ - data['1st Dose Date'][-1]).days))
            except:
                data['Days to 1st'].append('')
            data['2nd Dose Date'].append(participant_data[participant]['2nd Dose Date'])
            try:
                data['Days to 2nd'].append(int((date_ - data['2nd Dose Date'][-1]).days))
            except:
                data['Days to 2nd'].append('')
            data['3rd Dose Date'].append(participant_data[participant]['3rd Dose Date'])
            try:
                data['Days to 3rd'].append(int((date_ - data['3rd Dose Date'][-1]).days))
            except:
                data['Days to 3rd'].append('')
            data['Boost Date'].append(participant_data[participant]['Boost Date'])
            try:
                data['Days to Boost'].append(int((date_ - data['Boost Date'][-1]).days))
            except:
                data['Days to Boost'].append('')
            data['Boost Vaccine'].append(participant_data[participant]['Boost Vaccine'])
            data['Boost 2 Date'].append(participant_data[participant]['Boost 2 Date'])
            try:
                data['Days to Boost 2'].append(int((date_ - data['Boost 2 Date'][-1]).days))
            except:
                data['Days to Boost 2'].append('')            
            
            data['Boost 2 Vaccine'].append(participant_data[participant]['Boost 2 Vaccine'])
            data['Seronet ID'].append(seronet_id)
            data['Participant ID'].append(participant)
            data['Date'].append(date_)
            data['Post-Baseline'].append((date_ - baseline).days)
            data['Sample ID'].append(sample_id)
            data['Visit Type'].append(visit_type)
            data['Qualitative'].append(result)
            data['Quantitative'].append(result_new)
            if sample_id in research_results.index:
                for col in ['Spike endpoint', 'AUC']:
                    data[col].append(research_results.loc[sample_id, col])
                if type(data['AUC'][-1]) in [int, float]:
                    data['Log2AUC'].append(np.log2(data['AUC'][-1]))
                else:
                    data['Log2AUC'].append('-')
            else:
                for col in ['Spike endpoint', 'AUC']:
                    data[col].append('-')
                data['Log2AUC'].append('-')
            data['coll_inits'].append(coll_inits)
    report = pd.DataFrame(data)
    if not debug:
        report.to_excel(util.unfiltered, index=False)
        print("Report written to {}".format(util.unfiltered))
    print("{} samples characterized.".format(report.shape[0]))
    return report

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Annotate SERONET samples with processing data')
    argParser.add_argument('-d', '--debug', action='store_true')
    args = argParser.parse_args()
    pull_from_source(args.debug)