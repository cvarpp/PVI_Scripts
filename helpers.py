import util
import hashlib as hlib
from codecs import encode
from codecs import decode
import zipfile as zf
import pandas as pd
import numpy as np
import util
import os
class ValuesToClass(object):
    def __init__(self,values):
        for key in values:
            setattr(self, key, values[key])

def try_convert(val):
    '''
    Converts number values to floating point and converts 0s to 1s.
    Prepares data for log-transformation. Returns np.nan when conversion fails.
    '''
    if val == 0.:
        return 1.
    try:
        return float(str(val).strip())
    except:
        return np.nan

def clean_auc(df):
    '''
    Cleans AUC values in a dataframe, converting any negative values to 1.
    '''
    df = df.copy()
    neg_spike = df['Spike endpoint'].astype(str).str.strip().str.upper().str[:2] == "NE"
    neg_auc = df['AUC'] == 0.
    df.loc[neg_spike | neg_auc, 'AUC'] = 1.
    return df['AUC'].apply(try_convert).astype(float)

def clean_sample_id(df):
    '''
    Converts sample IDs to strings, stripping leading and trailing spaces as well as
    converting all alphabetic characters to uppercase.
    '''
    return df['Sample ID'].astype(str).str.strip().str.upper().str.replace('\.0', '', regex=True)

def clean_research(df):
    '''
    Cleans AUC and sample_id columns of common issues and drops rows missing AUC.

    Meant to be passed to the pandas `pipe` method
    '''
    return (df.assign(sample_id=clean_sample_id,
                      AUC=clean_auc)
              .dropna(subset=['AUC'])
              .query("AUC not in ['-']")
              .loc[:, ['sample_id', 'Spike endpoint', 'AUC']]
              .assign(Log2AUC=lambda df: np.log2(df['AUC'])))

def convert_serum(vol):
    '''
    Cleans serum values of known issues (inconsistent units and reporting of those units).

    Returns values in mL and assumes unitless values are in mL.
    '''
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

def fallible(f, default=np.nan):
    '''
    A wrapper around a function which returns a specified default value in case of error.
    
    Common use: `df['Float Version of C1'] = df['C1'].apply(fallible(float))`
    '''
    def tmp(val):
        try:
            return f(val)
        except:
            return default
    return tmp

def query_dscf(sid_list=None, no_pbmcs=set(), use_cache=False, update_cache=False):
    '''
    Retrieve processing data (volumes, cell counts, and times) for all samples.
    Aggregates across 3 separate files and multiple tabs.

    `sid_list` retrieves a specified subset of sample IDs

    `no_pbmcs` will override initially reported PBMC counts to reflect used sample IDs
    '''
    if use_cache and os.path.exists('local_cache/dscf.h5') and '/proc_info' in pd.HDFStore('local_cache/dscf.h5', mode='r').keys():
        all_samples = pd.read_hdf(util.proc + 'dscf.h5', key='proc_info')
    else:
        correct_2p = {'Comments': 'COMMENTS',
                    'Date of specimen processed ': 'Date Processing Started',}
        dscf_info = pd.ExcelFile(util.dscf)
        bsl2p_samples = dscf_info.parse(sheet_name='BSL2+ Samples', header=1).rename(columns=correct_2p)
        bsl2_samples = dscf_info.parse(sheet_name='BSL2 Samples')
        dscf_archive = pd.ExcelFile(util.proc + 'DSCF Archive/Data Sample Collection Form Archive.xlsx')
        bsl2p_archive = dscf_archive.parse(sheet_name='BSL2+ Samples', header=1).rename(columns=correct_2p)
        bsl2_archive = dscf_archive.parse(sheet_name='BSL2 Samples')
        correct_crp = {'CELL COUNTER (Total)': 'Cell Count',
                    '# Aliquots Frozen': '# of aliquots frozen',
                    'Total Volume of Serum after first spin (ml)': 'Total volume of serum (mL)',
                    'ACD VOLUME': 'CPT/EDTA VOL',
                    'Date of specimen processed': 'Date Processing Started'}
        crp_samples = dscf_info.parse(sheet_name='CRP').rename(columns=correct_crp)
        crp_samples['# cells per aliquot'] = crp_samples.apply(fallible(lambda row: row['Cell Count'] / row['# of aliquots frozen']), axis=1)
        crp_samples.loc[crp_samples['Total volume of serum after second spin (ml)'].str.upper().str.strip() != 'X', 'Total volume of serum (mL)'] = crp_samples.loc[crp_samples['Total volume of serum after second spin (ml)'].str.upper().str.strip() != 'X', 'Total volume of serum after second spin (ml)']
        date_cols = ['Date Processing Started']
        correct_new = {'# PBMCs per Aliquot': '# cells per aliquot',
                    '# Aliquots': '# of aliquots frozen',
                    'Total Plasma Vol. (mL)': 'Total volume of plasma (mL)',
                    'Total Serum Vol. (mL)': 'Total volume of serum (mL)',
                    'SST Volume': 'SST VOL',
                    'Cell Tube Volume (mL)': 'CPT/EDTA VOL',
                    'Time in -80C (Serum)': 'Time put in -80: SERUM',
                    'Time in Freezing Device': 'Time put in -80: PBMC',}
        new_samples = pd.read_excel(util.proc + 'Processing Notebook.xlsx', sheet_name='Specimen Dashboard', header=1).rename(columns=correct_new)
        dataframe_list = [bsl2p_archive, bsl2_archive, bsl2p_samples, bsl2_samples, crp_samples, new_samples]
        all_samples = (pd.concat(dataframe_list)
                        .assign(sample_id=clean_sample_id)
                        .drop_duplicates(subset=['sample_id'], keep='last')
                        .assign(serum_vol=clean_serum)
                        .pipe(clean_cells, no_pbmcs)
                        .pipe(map_dates, date_cols))
        if update_cache:
            if not os.path.exists('local_cache/'):
                os.mkdir('local_cache/')
            all_samples.to_hdf('local_cache/dscf.h5', key='proc_info')
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

def query_research(sid_list=None):
    '''
    Accesses research results from Krammer lab ELISA testing.
    Prioritizes the most recent results for any re-testing.

    The optional parameter `sid_list` filters results to a given subset of samples.
    '''
    research_source = pd.ExcelFile(util.research)
    research_samples_1 = research_source.parse(sheet_name='Inputs').pipe(clean_research)
    research_samples_2 = research_source.parse(sheet_name='Archive').pipe(clean_research)
    all_research_results = pd.concat([research_samples_2, research_samples_1]).drop_duplicates(subset=['sample_id'], keep='last')
    if sid_list is not None:
        research_results = all_research_results.query('sample_id in @sid_list')
    else:
        research_results = all_research_results
    return research_results.set_index('sample_id').copy()

def clean_path(df):
    '''
    Cleans the various antibody results returned from the clinical micro lab and adds
    log-transformed values for the continuous values.

    Meant to be passed to the pandas `pipe` method.
    '''
    df = df.copy()
    df['Qualitative'] = df[util.qual].apply(lambda val: "Negative" if str(val).strip().upper()[:2] == "NE" else val)
    df['Quant_str'] = df[util.quant].fillna('').astype(str)
    df['Quantitative'] = df.apply(lambda row: 1 if row['Qualitative'] == 'Negative' else row[util.quant], axis=1).apply(try_convert).astype(float)
    df['COV22_str'] = df['COV22 Results'].fillna('').astype(str)
    df['COV22'] = df['COV22 Results'].apply(lambda val: 1 if str(val).strip().upper()[:2] == 'NE' else val).apply(try_convert).astype(float)
    df['Log2Quant'] = np.log2(df['Quantitative'])
    df['Log2COV22'] = np.log2(df['COV22'])
    return df

def query_intake(participants=None, include_research=False, use_cache=False, update_cache=False, from_date=None):
    '''
    Queries the intake log, optionally limited to a subset of participants.

    `participants` is a list of participant IDs to filter on

    `include_research` will join the resulting dataframe with the results from `query_research`
    for convenience

    'from_date' only includes samples collected on or after a specified date
    '''
    if use_cache and os.path.exists('local_cache/dscf.h5') and '/intake_info' in pd.HDFStore('local_cache/dscf.h5', mode='r').keys():
        all_samples = pd.read_hdf(util.tracking + 'intake.h5', key='intake_info')
    else:
        drop_cols = ['Date Processed', 'Tubes Pulled?', 'Process Location', 'Box #', 'Specimen Type Sent to Patho', 'Date Given to Patho', 'Delivered by', 'Received by', 'Results received', 'RT-QPCR Result NS (Clinical)', 'Viability', 'Notes', 'Blood Collector Initials', 'New or Follow-up?', 'Participant ID', 'Sample ID', 'Ab Detection S/P Result (Clinical) (Titer or Neg)', 'Ab Concentration (Units - AU/mL)']
        intake_source = pd.ExcelFile(util.intake)
        intake = intake_source.parse(sheet_name='Sample Intake Log', header=util.header_intake)
        latest_intake = pd.read_excel(util.tracking + 'Sample Intake Log.xlsx', sheet_name='Sample Intake Log', header=6)
        date_cols = ['Date Collected']
        all_samples = (pd.concat([intake, latest_intake]).dropna(subset=['Participant ID'])
                            .assign(participant_id=lambda df: df['Participant ID'].str.strip().str.upper(),
                                    sample_id=clean_sample_id)
                            .query('sample_id.str.len() == 5')
                            .pipe(clean_path)
                            .pipe(map_dates, date_cols)
                            .drop_duplicates(subset=['sample_id', 'Date Collected'], keep='last')
                            .set_index('sample_id')
                            .drop(drop_cols, axis=1)
                            .sort_values(by=['participant_id', 'Date Collected']))
        if update_cache:
            if not os.path.exists('local_cache/'):
                os.mkdir('local_cache/')
            all_samples.to_hdf(util.tracking + 'intake.h5', key='intake_info')
    if participants is not None:
        samples = all_samples.query('participant_id in @participants').copy()
    if from_date is not None:
        all_samples['Date Collected'] = pd.to_datetime(all_samples['Date Collected'])
        dt = pd.to_datetime(from_date)
        samples = all_samples.query('`Date Collected` >= @dt')
    else:
        samples = all_samples
    if include_research:
        samples = samples.join(query_research())
    return samples

def try_datediff(start_date, end_date):
    try:
        return int((end_date - start_date).days)
    except:
        return np.nan

def coerce_date(val):
    '''
    A shorter wrapper for `apply`ing to create typed date columns from
    potentially messy data.
    '''
    return pd.to_datetime(val, errors='coerce').date()

def permissive_datemax(date_list, comp_date):
    '''
    Returns the last date before a given `comp_date` in `date_list`.

    Ignores errors quite freely, and defaults to 1/1/1950.

    Specifically useful for repeated COVID infections and vaccinations.
    '''
    placeholder = pd.to_datetime('1.1.1950').date()
    max_date = placeholder
    for date_ in date_list:
        date_ = coerce_date(date_)
        if not pd.isna(date_) and date_ > max_date and date_ < comp_date:
            max_date = date_
    if max_date > placeholder:
        return max_date

def map_dates(df, date_cols):
    '''
    Converts a set of date columns to the proper type in one function call.
    '''
    df = df.copy()
    for col in date_cols:
        df[col] = df[col].apply(coerce_date)
    return df


def corned_beef(userkey):
    try:
        locked = zf.ZipFile(util.script_folder + 'data/Corned_Beef_reference.zip', 'r')
        idset = locked.read("Corned_Beef_reference.txt")
    except:
        return None
    hash_machine = hlib.sha256()
    hash_machine.update(encode(userkey))
    hash = hash_machine.hexdigest()
    if hash in decode(idset):
        return("Validated")
    else:
        return None

def immune_history(vaccine_dates, vaccine_types, infection_dates, visit_date):
   '''
   Returns a string listing all of a participant's previous immune events on a given 'visit_date'.
   
   'vaccine_dates', 'vaccine types', and 'infection dates'  are all lists.

   Example Application: sample_info['Immune History'] = sample_info.apply(lambda row: immune_history(row[vaccine_date_cols], row[vaccine_type_cols], row[infection_date_cols], row['Date Collected']), axis=1)
   '''
   events = {'Event Date':[], 'Event Type':[]}
   for date, type in zip(vaccine_dates, vaccine_types):
      if pd.to_datetime(date) < pd.to_datetime(visit_date):
         events['Event Date'].append(date)
      if pd.to_datetime(date) < pd.to_datetime(visit_date) and pd.to_datetime(date) >= pd.to_datetime('2024-08-29'):
         if 'novavax' in str(type).lower():
            events['Event Type'].append('JN1')
         else:
            events['Event Type'].append('KP2')
      elif pd.to_datetime(date) < pd.to_datetime(visit_date) and pd.to_datetime(date) >= pd.to_datetime('2023-08-29'):
         events['Event Type'].append('XBB')
      elif pd.to_datetime(date) < pd.to_datetime(visit_date) and pd.to_datetime(date) >= pd.to_datetime('2022-08-29'):
         events['Event Type'].append('BvB')
      elif pd.to_datetime(date) < pd.to_datetime(visit_date):
         events['Event Type'].append('V') 
   for infection in infection_dates:
      if pd.to_datetime(infection) < pd.to_datetime(visit_date):
         events['Event Date'].append(infection)
         events['Event Type'].append('I') 
   events_frame = pd.DataFrame.from_dict(events).set_index('Event Date')
   events_sorted = events_frame.sort_index()
   history = "-".join(events_sorted['Event Type'])
   return history
