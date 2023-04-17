import pandas as pd
import numpy as np
import util

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
        return df['Total volume of serum (mL)'].apply(convert_serum)

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

def query_dscf(sid_list=None, no_pbmcs=set()):
    correct_2p = {'Comments': 'COMMENTS'}
    dscf_info = pd.ExcelFile(util.dscf)
    bsl2p_samples = dscf_info.parse(sheet_name='BSL2+ Samples', header=1).rename(columns=correct_2p)
    bsl2_samples = dscf_info.parse(sheet_name='BSL2 Samples')
    date_cols = ['Date Processing Started']
    all_samples = (pd.concat([bsl2p_samples, bsl2_samples])
                     .assign(sample_id=clean_sample_id)
                     .drop_duplicates(subset=['sample_id'], keep='last')
                     .assign(serum_vol=clean_serum)
                     .pipe(clean_cells, no_pbmcs)
                     .pipe(map_dates, date_cols))
    if sid_list is not None:
        samples = all_samples.query('sample_id in @sid_list').copy()
    else:
        samples = all_samples
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
    sample_info = samples.rename(columns=col_mapper).set_index('sample_id')
    return sample_info.copy()

def clean_path(df):
    df = df.copy()
    df['Qualitative'] = df[util.qual].apply(lambda val: "Negative" if str(val).strip().upper()[:2] == "NE" else val)
    df['Quantitative'] = df.apply(lambda row: 1 if row['Qualitative'] == 'Negative' else row[util.quant], axis=1)
    return df

def query_intake(participants=None):
    drop_cols = ['Date Processed', 'Tubes Pulled?', 'Process Location', 'Box #', 'Specimen Type Sent to Patho', 'Date Given to Patho', 'Delivered by', 'Received by', 'Results received', 'RT-QPCR Result NS (Clinical)', 'Viability', 'Notes', 'Blood Collector Initials', 'New or Follow-up?', 'Participant ID', 'Sample ID', 'Ab Detection S/P Result (Clinical) (Titer or Neg)', 'Ab Concentration (Units - AU/mL)']
    intake_source = pd.ExcelFile(util.intake)
    intake = intake_source.parse(sheet_name='Sample Intake Log', header=util.header_intake)
    date_cols = ['Date Collected']
    all_samples = (intake.dropna(subset=['Participant ID'])
                        .assign(participant_id=lambda df: df['Participant ID'].str.strip().str.upper(),
                                sample_id=clean_sample_id)
                        .query('sample_id.str.len() == 5')
                        .pipe(clean_path)
                        .pipe(map_dates, date_cols)
                        .set_index('sample_id')
                        .drop(drop_cols, axis=1))
    if participants is not None:
        samples = all_samples.query('participant_id in @participants').copy()
    else:
        samples = all_samples
    return samples

def query_research(sid_list=None):
    research_source = pd.ExcelFile(util.research)
    research_samples_1 = research_source.parse(sheet_name='Inputs').pipe(clean_research)
    research_samples_2 = research_source.parse(sheet_name='Archive').pipe(clean_research)
    all_research_results = pd.concat([research_samples_2, research_samples_1]).drop_duplicates(subset=['sample_id'], keep='last')
    if sid_list is not None:
        research_results = all_research_results.query('sample_id in @sid_list')
    else:
        research_results = all_research_results
    return research_results.set_index('sample_id').copy()


def try_datediff(start_date, end_date):
    try:
        return int((end_date - start_date).days)
    except:
        return np.nan

def coerce_date(val):
    return pd.to_datetime(val, errors='coerce').date()

def permissive_datemax(date_list, comp_date):
    placeholder = pd.to_datetime('1.1.1950').date()
    max_date = placeholder
    for date_ in date_list:
        date_ = coerce_date(date_)
        if not pd.isna(date_) and date_ > max_date and date_ < comp_date:
            max_date = date_
    if max_date > placeholder:
        return max_date

def map_dates(df, date_cols):
    df = df.copy()
    for col in date_cols:
        df[col] = df[col].apply(coerce_date)
    return df
