import pandas as pd
import numpy as np
import util

def clean_auc(df):
    df = df.copy()
    neg_spike = df['Spike endpoint'].astype(str).str.strip().str.upper().str[:1] == "N"
    neg_auc = df['AUC'] == 0.
    df.loc[neg_spike | neg_auc, 'AUC'] = 1.
    return df['AUC']

def clean_sample_id(df):
    return df['Sample ID'].astype(str).str.strip().str.upper()

def clean_research(df):
    return (df.assign(sample_id=clean_sample_id,
                      AUC=clean_auc)
              .dropna(subset=['AUC'])
              .loc[:, ['sample_id', 'Spike endpoint', 'AUC']])

def convert_serum(vol):
    try:
        serum_volume = float(str(vol).strip().strip("mlML ").split()[0])
    except:
        try:
            serum_volume = float(str(vol).strip().strip("ulUL ").split()[0]) / 1000.
        except:
            if type(vol) == str:
                serum_volume = 0
            else:
                serum_volume = vol
    if type(serum_volume) != str and serum_volume > 4.5:
        serum_volume = 4.5
    return serum_volume

def clean_serum(df):
        return df['Total volume of serum (ml)'].apply(convert_serum)

def clean_pbmc(row, no_pbmcs):
    pbmc_conc = row['# cells per aliquot']
    if row['sample_id'] in no_pbmcs:
        return 0.
    elif (type(pbmc_conc) != float or pd.isna(pbmc_conc)) and type(pbmc_conc) != int:
        return 0.
    else:
        return float(pbmc_conc)

def clean_vials(row):
    vial_count = row['# of aliquots frozen']
    if row['pbmc_conc'] == 0.:
        return 0
    elif type(vial_count) in [int, float] and vial_count > 2:
        return 2
    else:
        return vial_count

def clean_cells(df, no_pbmcs):
    df = df.copy()
    df['pbmc_conc'] = df.apply(lambda row: clean_pbmc(row, no_pbmcs), axis=1)
    df['vial_count'] = df.apply(clean_vials, axis=1)
    return df

def sufficient(val):
    try:
        return float(val) > 0
    except:
        return False

def pull_from_source():
    iris_data = pd.read_excel(util.iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID'])
    iris_data['Participant ID'] = iris_data['Participant ID'].apply(lambda val: val.strip().upper())
    titan_data = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_data['Participant ID'] = titan_data['Umbrella Corresponding Participant ID'].apply(lambda val: val.strip().upper())
    mars_data = pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID'])
    mars_data['Participant ID'] = mars_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = [p.strip().upper() for p in np.concatenate((mars_data['Participant ID'].unique(), titan_data['Participant ID'].unique(), iris_data['Participant ID'].unique()))]
    participant_study = {p: 'MARS' for p in mars_data['Participant ID'].unique()}
    participant_study.update({p: 'TITAN' for p in titan_data['Participant ID'].unique()})
    participant_study.update({p: 'IRIS' for p in iris_data['Participant ID'].unique()})
    mars_data.set_index('Participant ID', inplace=True)
    iris_data.set_index('Participant ID', inplace=True)
    titan_data.set_index('Participant ID', inplace=True)
    titan_convert = {
        'Vaccine': 'Vaccine Type',
        '1st Dose Date': 'Vaccine #1 Date',
        '2nd Dose Date': 'Vaccine #2 Date',
        'Boost Date': '3rd Dose Vaccine Date',
        'Boost Vaccine': '3rd Dose Vaccine Type'
    }
    mars_convert = {
        'Vaccine': 'Vaccine Name',
        '1st Dose Date': 'Vaccine #1 Date',
        '2nd Dose Date': 'Vaccine #2 Date',
        'Boost Date': '3rd Vaccine',
        'Boost Vaccine': '3rd Vaccine Type '
    }
    iris_convert = {
        'Vaccine': 'Which Vaccine?',
        '1st Dose Date': 'First Dose Date',
        '2nd Dose Date': 'Second Dose Date',
        'Boost Date': 'Third Dose Date',
        'Boost Vaccine': 'Third Dose Type'
    }
    source_dfs = {'TITAN': titan_data, 'MARS': mars_data, 'IRIS': iris_data}
    conversions = {'TITAN': titan_convert, 'MARS': mars_convert, 'IRIS': iris_convert}
    intake_source = pd.ExcelFile(util.intake)
    collection_log = intake_source.parse(sheet_name='Collection Log').assign(sample_id=lambda df: df['Sample ID'].astype(str).str.strip().str.upper()).drop('Sample ID', axis=1).set_index('sample_id')
    samples = intake_source.parse(sheet_name='Sample Intake Log', header=util.header_intake, dtype=str)
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    samplesClean = samples.dropna(subset=['Participant ID'])
    participant_samples = {participant: [] for participant in participants}
    submitted_key = pd.read_excel(util.clin_ops + 'Cross-Project/Seronet Task D4/Data/SERONET Key.xlsx', sheet_name='Source').drop_duplicates(subset=['Participant ID']).set_index('Participant ID')
    samples_of_interest = set()
    for _, sample in samplesClean.iterrows():
        sample_id = str(sample['Sample ID']).strip().upper()
        if len(sample_id) != 5:
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
            participant_samples[participant].append((sample_date, sample_id, sample[visit_type], sample[newCol], result_new, sample['Blood Collector Initials']))
            samples_of_interest.add(sample_id)

    research_source = pd.ExcelFile(util.research)
    research_samples_1 = research_source.parse(sheet_name='Inputs').pipe(clean_research)
    research_samples_2 = research_source.parse(sheet_name='Archive').pipe(clean_research)
    research_results = pd.concat([research_samples_2, research_samples_1]).drop_duplicates(subset=['sample_id'], keep='last').query('sample_id in @samples_of_interest').set_index('sample_id')

    correct_2p = {'Comments': 'COMMENTS'}
    shared_samples = pd.read_excel(util.shared_samples, sheet_name='Released Samples')
    no_pbmcs = set([str(sid).strip().upper() for sid in shared_samples[shared_samples['Sample Type'] == 'PBMC']['Sample ID'].unique()])
    dscf_info = pd.ExcelFile(util.dscf)
    bsl2p_samples = dscf_info.parse(sheet_name='BSL2+ Samples', header=1).assign(sample_id=clean_sample_id).query('sample_id in @samples_of_interest').rename(columns=correct_2p)
    bsl2_samples = dscf_info.parse(sheet_name='BSL2 Samples')
    all_samples = pd.concat([bsl2p_samples, bsl2_samples]).assign(sample_id=clean_sample_id).query('sample_id in @samples_of_interest').drop_duplicates(subset=['sample_id'], keep='last').assign(serum_vol=clean_serum).pipe(clean_cells, no_pbmcs)
    all_samples['coll_time'] = all_samples.apply(lambda row: collection_log.loc[row['sample_id'], 'Time Collected'] if row['sample_id'] in collection_log.index else row['Time Collected'], axis=1)
    serum_or_cells = all_samples['serum_vol'].apply(sufficient) | all_samples['pbmc_conc'].apply(sufficient)
    all_samples[~serum_or_cells].to_excel(util.script_output + 'missing_info.xlsx', index=False)
    col_mapper = {
        'serum_vol': 'Volume of Serum Collected (mL)',
        'pbmc_conc': 'PBMC concentration per mL (x10^6)',
        'vial_count': '# of PBMC vials',
        'Time Received': 'rec_time',
        'Time Started Processing': 'proc_time',
        'Time put in -80: SERUM': 'serum_freeze_time',
        'Time put in -80: PBMC': 'cell_freeze_time',
        'Processed by (initials)': 'proc_inits',
        '% Viability': 'viability',
        'CPT/EDTA VOL': 'cpt_vol',
        'SST VOL': 'sst_vol',
        'COMMENTS': 'proc_comment'
    }
    sample_info = all_samples[serum_or_cells].rename(columns=col_mapper).set_index('sample_id')

    proc_cols = ['Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials', 'coll_time', 'rec_time', 'proc_time', 'serum_freeze_time', 'cell_freeze_time', 'proc_inits', 'viability', 'cpt_vol', 'sst_vol', 'proc_comment']

    columns = ['Cohort', 'Seronet ID', 'Vaccine', '1st Dose Date', 'Days to 1st', '2nd Dose Date', 'Days to 2nd', 'Boost Vaccine', 'Boost Date', 'Days to Boost', 'Participant ID', 'Date', 'Post-Baseline', 'Sample ID', 'Visit Type', 'Qualitative', 'Quantitative', 'Spike endpoint', 'AUC', 'Log2AUC', 'Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials', 'coll_inits', 'coll_time', 'rec_time', 'proc_time', 'serum_freeze_time', 'cell_freeze_time', 'proc_inits', 'viability', 'cpt_vol', 'sst_vol', 'proc_comment']
    print("Samples not in DSCF:")
    data = {col: [] for col in columns}
    participant_data = {}
    for participant, samples in participant_samples.items():
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
        key_cols =['Vaccine', '1st Dose Date', '2nd Dose Date', 'Boost Date', 'Boost Vaccine']
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
        elif participant_study[participant] == 'PRIORITY':
            for k in key_cols:
                participant_data[participant][k] = ''
        else:
            print(participant, "is not in the study! They are in", participant_study[participant])
            exit(1)
        if participant in submitted_key.index:
            seronet_id = submitted_key.loc[participant, 'Research_Participant_ID']
        for date_, sample_id, visit_type, result, result_new, coll_inits in samples:
            sample_id = str(sample_id).strip().upper()
            if sample_id not in sample_info.index:
                print(sample_id)
                continue
            for col in proc_cols:
                data[col].append(sample_info.loc[sample_id, col])
            data['Cohort'].append(participant_study[participant])
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
            data['Boost Date'].append(participant_data[participant]['Boost Date'])
            try:
                data['Days to Boost'].append(int((date_ - data['Boost Date'][-1]).days))
            except:
                data['Days to Boost'].append('')
            data['Boost Vaccine'].append(participant_data[participant]['Boost Vaccine'])
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
                data['Log2AUC'].append(np.log2(data['AUC'][-1]))
            else:
                for col in ['Spike endpoint', 'AUC']:
                    data[col].append('-')
                data['Log2AUC'].append('-')
            data['coll_inits'].append(coll_inits)
    report = pd.DataFrame(data)
    report.to_excel(util.unfiltered, index=False)
    print("Hopefully not too many sample IDs above!")
    return report

if __name__ == '__main__':
    pull_from_source()