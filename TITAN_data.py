import pandas as pd
import numpy as np
import util
import argparse

def clean_auc(df):
    df = df.copy()
    neg_spike = df['Spike endpoint'].astype(str).str.strip().str.upper().str[:2] == "NE"
    neg_auc = df['AUC'] == 0.
    df.loc[neg_spike | neg_auc, 'AUC'] = 1.
    return df['AUC']

def clean_sample_id(df):
    return df['Sample ID'].astype(str).str.strip().str.upper()

def clean_research(df):
    return (df.assign(sample_id=clean_sample_id,
                      AUC=clean_auc)
              .dropna(subset=['AUC'])
              .query("AUC not in ['-']")
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
    else:
        return vial_count

def clean_cells(df, no_pbmcs):
    df = df.copy()
    df['pbmc_conc'] = df.apply(lambda row: clean_pbmc(row, no_pbmcs), axis=1)
    df['vial_count'] = df.apply(clean_vials, axis=1)
    return df

def make_timepoint(days):
    tps = [30, 90, 180, 300]
    windows = [14, 21, 21, 21]
    try:
        if days <= 0:
            return 'Pre-Dose 3'
        else:
            for tp, window in zip(tps, windows):
                if abs(days - tp) <= window:
                    return 'Day {}'.format(tp)
    except:
        return "None"
    return "None"

def try_convert_date(val):
    try:
        return pd.to_datetime(str(val))
    except:
        return val

def map_dates(df, date_cols):
    df = df.copy()
    for col in date_cols:
        df[col] = df[col].apply(try_convert_date)
    return df

def query_titan():
    titan_data = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_data['Participant ID'] = titan_data['Umbrella Corresponding Participant ID'].apply(lambda val: val.strip())
    titan_data.set_index('Participant ID', inplace=True)
    titan_convert = {
        'Vaccine': 'Vaccine Type',
        '1st Dose Date': 'Vaccine #1 Date',
        '2nd Dose Date': 'Vaccine #2 Date',
        'Boost Date': '3rd Dose Vaccine Date',
        'Boost Vaccine': '3rd Dose Vaccine Type',
        'Boost 2 Date': 'Booster Dose Date',
        'Boost 2 Vaccine': 'Vaccine Type.1'
    }
    cols = [k for k in titan_convert]
    date_cols = [col for col in cols if 'Date' in col]
    reverse_convert = {v: k for k, v in titan_convert.items()}
    participant_data = (titan_data.rename(columns=reverse_convert)
                            .loc[:, cols]
                            .pipe(map_dates, date_cols))
    return participant_data

def query_dscf(samples_of_interest):
    correct_2p = {'Comments': 'COMMENTS'}
    no_pbmcs = set()
    dscf_info = pd.ExcelFile(util.dscf)
    bsl2p_samples = dscf_info.parse(sheet_name='BSL2+ Samples', header=1).assign(sample_id=clean_sample_id).query('sample_id in @samples_of_interest').rename(columns=correct_2p)
    bsl2_samples = dscf_info.parse(sheet_name='BSL2 Samples')
    all_samples = pd.concat([bsl2p_samples, bsl2_samples]).assign(sample_id=clean_sample_id).query('sample_id in @samples_of_interest').drop_duplicates(subset=['sample_id'], keep='last').assign(serum_vol=clean_serum).pipe(clean_cells, no_pbmcs)
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
    keep_cols = ['Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials']
    sample_info = all_samples.rename(columns=col_mapper).set_index('sample_id').loc[:, keep_cols]
    return sample_info.copy()

def clean_quant(row):
    if row['Qualitative'] == "Negative":
        return 1.
    else:
        return

def clean_path(df):
    df = df.copy()
    df['Qualitative'] = df[util.qual].apply(lambda val: "Negative" if str(val).strip().upper()[:2] == "NE" else val)
    df['Quantitative'] = df.apply(lambda row: 1 if row['Qualitative'] == 'Negative' else row[util.quant], axis=1)
    return df

def query_intake(participants):
    drop_cols = ['Order #', 'PVI #', 'Date Processed', 'Tubes Pulled?', 'Process Location', 'Box', 'Specimen Type Sent to Patho', 'Date Given to Patho', 'Delivered by', 'Received by', 'Results received', 'RT-QPCR Result NS (Clinical)', 'Clinical Ab Result Shared?', 'Date Shared', 'Shared By', 'Viability', 'Notes', 'Blood Collector Initials', 'New or Follow-up?', 'Participant ID', 'Sample ID', 'Ab Detection S/P Result (Clinical) (Titer or Neg)', 'Ab Concentration (Units - AU/mL)']
    intake_source = pd.ExcelFile(util.intake)
    samples = intake_source.parse(sheet_name='Sample Intake Log', header=util.header_intake)
    samplesClean = (samples.dropna(subset=['Participant ID'])
                        .assign(participant_id=lambda df: df['Participant ID'].str.strip(),
                                sample_id=clean_sample_id)
                        .query('participant_id in @participants')
                        .query('sample_id.str.len() == 5')
                        .pipe(clean_path)
                        .set_index('sample_id')
                        .drop(drop_cols, axis=1))
    return samplesClean

def date_subtract(sample_date, reference_date):
    try:
        return int((sample_date - reference_date).days)
    except:
        return pd.NA

def pull_data():
    participant_data = query_titan()
    participants = [p for p in participant_data.index]
    participant_samples = query_intake(participants)
    samples_of_interest = participant_samples.index.to_numpy()
    sample_info = query_dscf(samples_of_interest)

    research_source = pd.ExcelFile(util.research)
    research_samples_1 = research_source.parse(sheet_name='Inputs').pipe(clean_research)
    research_samples_2 = research_source.parse(sheet_name='Archive').pipe(clean_research)
    research_results = pd.concat([research_samples_2, research_samples_1]).drop_duplicates(subset=['sample_id'], keep='last').query('sample_id in @samples_of_interest').set_index('sample_id')

    df = (participant_samples
            .join(sample_info)
            .join(research_results)
            .join(participant_data, on='participant_id')
            .rename(columns={'Date Collected': 'Date'}))
    date_cols = ['1st Dose Date', '2nd Dose Date', 'Boost Date', 'Boost 2 Date']
    for date_col in date_cols:
        day_col = 'Days to ' + date_col[:-5]
        df[day_col] = df.apply(lambda row: date_subtract(row['Date'], row[date_col]), axis=1)
    df.to_excel(util.script_folder + 'data/titan_intermediate.xlsx')
    return df

def titanify(df):
    df['AUC'] = df['AUC'].astype(float)
    writer = pd.ExcelWriter(util.script_output + 'new_format/titan_consolidated.xlsx')
    df.to_excel(writer, sheet_name='Long-Form')
    df['Timepoint'] = df['Days to Boost'].apply(make_timepoint)
    tp_order = ['Pre-Dose 3', 'Day 30', 'Day 90', 'Day 180', 'Day 300']
    tp_df = (df[df['Timepoint'] != 'None']
                .drop_duplicates(subset=['participant_id', 'Timepoint'], keep='last')
                .pivot_table(values='AUC', index='participant_id', columns='Timepoint')
                .reindex(tp_order, axis=1))
    tp_df.to_excel(writer, sheet_name='Wide-Form')
    tp_sample_ids = (df[df['Timepoint'] != 'None'].reset_index()
                .drop_duplicates(subset=['participant_id', 'Timepoint'], keep='last')
                .pivot_table(values='sample_id', index='participant_id', columns='Timepoint', aggfunc=np.sum)
                .reindex(tp_order, axis=1))
    tp_sample_ids.to_excel(writer, sheet_name='Wide-Form Sample ID Key')
    writer.save()
    writer.close()

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Annotate and split TITAN samples')
    argParser.add_argument('-u', '--update', action='store_true')
    args = argParser.parse_args()
    if args.update:
        titanify(pull_data())
    else:
        titanify(pd.read_excel(util.script_folder + 'data/titan_intermediate.xlsx', index_col='sample_id'))
