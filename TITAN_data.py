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

def transform_id(df):
    df = df.copy()
    transplant_types = ["Kidney", "Liver", "Other", "Multi"]
    transplant_str = "KLOM"
    df['HIV'] = df['Bloodborne Pathogen'].apply(lambda val: "H" if val == "HIV" else "_")
    df['COVID'] = df['COVID Pre-Enrollment'].apply(lambda val: "C+" if val == "Yes" else "__")
    df['Transplant'] = df['Transplant Group'].apply(lambda val: transplant_str[transplant_types.index(val)])
    df['Full ID'] = df['TITAN ID'] +'_' + df['Transplant'] + df['COVID'] + df['HIV']
    df['Full Transplant'] = (df['Transplant Group'] + ':' + df['Transplant Other'].fillna("")).str.strip(': ')
    return df

def query_titan():
    titan_data = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_data['Participant ID'] = titan_data['Umbrella Corresponding Participant ID'].apply(lambda val: val.strip())
    titan_data.set_index('Participant ID', inplace=True)
    titan_convert = {
        'Vaccine': 'Vaccine Type',
        '1st Dose Date': 'Vaccine #1 Date',
        '2nd Dose Date': 'Vaccine #2 Date',
        '3rd Dose Date': '3rd Dose Vaccine Date',
        '3rd Dose Vaccine': '3rd Dose Vaccine Type',
        'Boost Date': 'First Booster Dose Date',
        'Boost Vaccine': 'First Booster Vaccine Type',
        'Boost 2 Date': 'Second Booster Dose Date',
        'Boost 2 Vaccine': 'Second Booster Vaccine Type',
        'COVID Pre-Enrollment': 'Had Prior COVID (qualiying dose)?',
        'COVID Pre-Enrollment Date': 'Date of PCR positive',
        'Transplant Group': 'Transplant Group',
        'Bloodborne Pathogen': 'Blood Borne Path',
        'Age at Enrollment': 'Age at Enrollment',
        'Gender': 'Gender',
        'Study Participation Status': 'Study Participation Status',
        'TITAN ID': 'TITAN ID',
        'Transplant Other': 'Multi/ Other'
    }
    titan_data = titan_data[titan_data['Study Participation Status'].apply(lambda s: "Withdrawn" not in s)]
    cols = [k for k in titan_convert]
    date_cols = [col for col in cols if 'Date' in col]
    reverse_convert = {v: k for k, v in titan_convert.items()}
    participant_data = (titan_data.rename(columns=reverse_convert)
                            .loc[:, cols]
                            .pipe(map_dates, date_cols)
                            .pipe(transform_id))
    post3_covid_renamer = {
        'COVID Post-Third': 'COVID 19 since 3rd dose?',
        'COVID Post-Third Date': 'Date of + PCR',
        'Monoclonal Post-Third': 'moAb?',
        'Monoclonal Post-Third Date': 'moAb Date',
    }
    post4_covid_renamer = {
        'COVID Post-Boost': 'COVID 19 since first booster dose?',
        'COVID Post-Boost Date': 'Date of + PCR',
        'Monoclonal Post-Boost': 'moAb?',
        'Monoclonal Post-Boost Date': 'moAb Date',
    }
    post5_covid_renamer = {
        'COVID Post-Boost 2': 'COVID 19 since second booster dose?',
        'COVID Post-Boost 2 Date': 'Date of + PCR',
        'Monoclonal Post-Boost 2': 'moAb?',
        'Monoclonal Post-Boost 2 Date': 'moAb Date',
    }
    reverse_post3 = {v: k for k, v in post3_covid_renamer.items()}
    reverse_post4 = {v: k for k, v in post4_covid_renamer.items()}
    reverse_post5 = {v: k for k, v in post5_covid_renamer.items()}
    titan_third = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Third Dose', header=1).dropna(subset=['Umbrella Participant ID']).set_index('TITAN ID').rename(columns=reverse_post3)
    titan_third['Full Meds'] = titan_third.loc[:, 'Maintenance immunosuppresion at time of third dose '].fillna("") + ":" + titan_third.loc[:, 'Other, Specify'].fillna("")
    titan_fourth = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Booster Dose #1', header=1).dropna(subset=['Umbrella Participant ID']).set_index('TITAN ID').rename(columns=reverse_post4)
    titan_fourth['Full Meds'] = titan_fourth.loc[:, 'Maintenance immunosuppresion at time of first booster dose '].fillna("") + ":" + titan_fourth.loc[:, 'Other, Specify'].fillna("")
    titan_fifth = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Booster Dose #2', header=1).dropna(subset=['Umbrella Participant ID']).set_index('TITAN ID').rename(columns=reverse_post5)
    titan_fifth['Full Meds'] = titan_fifth.loc[:, 'Maintenance immunosuppresion at time of second booster dose '].fillna("") + ":" + titan_fifth.loc[:, 'Other, Specify'].fillna("")
    participant_data['Meds at third dose'] = participant_data['TITAN ID'].apply(lambda val: titan_third.loc[val, 'Full Meds'])
    participant_data['AM'] = participant_data['Meds at third dose'].apply(lambda val: "Yes" if "AM" in str(val) else "No")
    participant_data['Prednisone'] = participant_data['Meds at third dose'].apply(lambda val: "Yes" if "Pred" in str(val) else "No")
    participant_data['CNI'] = participant_data['Meds at third dose'].apply(lambda val: "Yes" if "CNI" in str(val) else "No")
    participant_data['Meds at first booster dose'] = participant_data['TITAN ID'].apply(lambda val: titan_fourth.loc[val, 'Full Meds'])
    participant_data['Meds at second booster dose'] = participant_data['TITAN ID'].apply(lambda val: titan_fifth.loc[val, 'Full Meds'])
    for k in post3_covid_renamer:
        participant_data[k] = participant_data['TITAN ID'].apply(lambda val: titan_third.loc[val, k])
    for k in post4_covid_renamer:
        participant_data[k] = participant_data['TITAN ID'].apply(lambda val: titan_fourth.loc[val, k])
    for k in post5_covid_renamer:
        participant_data[k] = participant_data['TITAN ID'].apply(lambda val: titan_fifth.loc[val, k])
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
    date_cols = ['1st Dose Date', '2nd Dose Date', '3rd Dose Date', 'Boost Date', 'Boost 2 Date', 'COVID Pre-Enrollment Date', 'COVID Post-Third Date', 'Monoclonal Post-Third Date', 'COVID Post-Boost Date', 'Monoclonal Post-Boost Date', 'COVID Post-Boost 2 Date', 'Monoclonal Post-Boost 2 Date']
    for date_col in date_cols:
        day_col = 'Days from ' + date_col[:-5]
        df[day_col] = df.apply(lambda row: date_subtract(row['Date'], row[date_col]), axis=1)
    df.dropna(subset=['AUC'], inplace=True)
    df.to_excel(util.script_folder + 'data/titan_intermediate.xlsx')
    return df

def titanify(df):
    df['AUC'] = df['AUC'].astype(float)
    df['Log2AUC'] = np.log2(df['AUC'])
    output_fname = util.script_output + 'new_format/titan_consolidated.xlsx'
    writer = pd.ExcelWriter(output_fname)
    df.to_excel(writer, sheet_name='Source')
    df['Timepoint'] = df['Days from 3rd Dose'].apply(make_timepoint)
    df['Boost Timepoint'] = df['Days from Boost'].apply(make_timepoint).apply(lambda val: "Pre-Dose 4" if val == "Pre-Dose 3" else val)
    longform_columns = ['Full ID', 'Transplant', 'COVID', 'HIV', 'Days from 3rd Dose', 'AUC', 'Spike endpoint', 'Log2AUC', 'Full Transplant', 'participant_id', 'AM', 'Prednisone', 'CNI', 'Gender', 'Age at Enrollment', 'Vaccine', 'Boost Vaccine', 'Timepoint', 'Boost Timepoint', 'Days from Boost', 'Days from Boost 2', 'Days from 1st Dose', 'Days from 2nd Dose', 'Days from COVID Pre-Enrollment', 'Days from COVID Post-Third', 'Days from COVID Post-Boost', 'Days from COVID Post-Boost 2', 'Days from Monoclonal Post-Third', 'Days from Monoclonal Post-Boost', 'Days from Monoclonal Post-Boost 2']
    df_long = df.loc[:, longform_columns].sort_values(by=['Full ID', 'Days from 3rd Dose'])
    df_long.to_excel(writer, sheet_name='Long-Form')
    tp_order = ['Pre-Dose 3', 'Day 30', 'Day 90', 'Day 180', 'Day 300']
    tp_order_boost = ['Pre-Dose 4', 'Day 30', 'Day 90', 'Day 180', 'Day 300']
    df_wide = (df[df['Timepoint'] != 'None']
                .drop_duplicates(subset=['Full ID', 'Timepoint'], keep='last')
                .pivot_table(values='AUC', index='Full ID', columns='Timepoint')
                .reindex(tp_order, axis=1))
    df_wide.to_excel(writer, sheet_name='Wide-Form')
    ppl_key = df_long.drop_duplicates(subset=['Full ID']).set_index('Full ID')
    df_wide_annot = df_wide.copy()
    for col in ['Transplant', 'COVID', 'HIV', 'AM', 'Prednisone']:
        df_wide_annot[col] = df_wide_annot.apply(lambda row: ppl_key.loc[row.name, col], axis=1)
    df_wide_annot.to_excel(writer, sheet_name='Wide-Form Annotated')
    tp_sample_ids = (df[df['Timepoint'] != 'None'].reset_index()
                .drop_duplicates(subset=['Full ID', 'Timepoint'], keep='last')
                .pivot_table(values='sample_id', index='Full ID', columns='Timepoint', aggfunc=np.sum)
                .reindex(tp_order, axis=1))
    tp_sample_ids.to_excel(writer, sheet_name='Wide-Form Sample ID Key')
    df_pre = df[df['Boost Timepoint'] == "Pre-Dose 4"]
    df_post = df[df['Boost Timepoint'] != "Pre-Dose 4"]

    df_wide[df_wide.apply(lambda row: ppl_key.loc[row.name, 'Transplant'] == 'K', axis=1)].to_excel(writer, sheet_name='Wide Kidney')
    df_wide[df_wide.apply(lambda row: ppl_key.loc[row.name, 'Transplant'] == 'L', axis=1)].to_excel(writer, sheet_name='Wide Liver')
    df_wide[df_wide.apply(lambda row: ppl_key.loc[row.name, 'Transplant'] in "OM", axis=1)].to_excel(writer, sheet_name='Wide Other+Multi')
    df_wide[df_wide.apply(lambda row: ppl_key.loc[row.name, 'HIV'] == 'H', axis=1)].to_excel(writer, sheet_name='Wide HIV pos')
    df_wide[df_wide.apply(lambda row: ppl_key.loc[row.name, 'HIV'] != 'H', axis=1)].to_excel(writer, sheet_name='Wide HIV neg')
    (df_pre[(df_pre['Timepoint'] != 'None')]
        .drop_duplicates(subset=['Full ID', 'Timepoint'], keep='last')
        .pivot_table(values='AUC', index='Full ID', columns='Timepoint')
        .reindex(tp_order, axis=1).to_excel(writer, sheet_name='Wide 3rd Dose'))
    (df[(df['Boost Timepoint'] != 'None')]
        .drop_duplicates(subset=['Full ID', 'Boost Timepoint'], keep='last')
        .pivot_table(values='AUC', index='Full ID', columns='Boost Timepoint')
        .reindex(tp_order_boost, axis=1).to_excel(writer, sheet_name='Wide 4th Dose'))
    df_wide[df_wide.apply(lambda row: ppl_key.loc[row.name, 'AM'] == 'Yes', axis=1)].to_excel(writer, sheet_name='Wide AM')
    df_wide[df_wide.apply(lambda row: ppl_key.loc[row.name, 'AM'] == 'No', axis=1)].to_excel(writer, sheet_name='Wide No AM')
    df_wide[df_wide.apply(lambda row: ppl_key.loc[row.name, 'Prednisone'] == 'Yes', axis=1)].to_excel(writer, sheet_name='Wide Prednisone')
    df_wide[df_wide.apply(lambda row: ppl_key.loc[row.name, 'Prednisone'] == 'No', axis=1)].to_excel(writer, sheet_name='Wide No Prednisone')

    df_long[df_long['Transplant'] == 'K'].to_excel(writer, sheet_name='Kidney')
    df_long[df_long['Transplant'] == 'L'].to_excel(writer, sheet_name='Liver')
    df_long[df_long['Transplant'].apply(lambda val: val in "OM")].to_excel(writer, sheet_name='Other+Multi')
    df_long[df_long['HIV'] == 'H'].to_excel(writer, sheet_name='HIV pos')
    df_long[df_long['HIV'] != 'H'].to_excel(writer, sheet_name='HIV neg')
    df_pre.to_excel(writer, sheet_name='3rd Dose')
    df_post.to_excel(writer, sheet_name='4th Dose')
    df_long[df_long['AM'] == 'Yes'].to_excel(writer, sheet_name='AM')
    df_long[df_long['AM'] == 'No'].to_excel(writer, sheet_name='No AM')
    df_long[df_long['Prednisone'] == 'Yes'].to_excel(writer, sheet_name='Prednisone')
    df_long[df_long['Prednisone'] == 'No'].to_excel(writer, sheet_name='No Prednisone')
    writer.save()
    writer.close()
    print("Data written to {}".format(output_fname))

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Annotate and split TITAN samples')
    argParser.add_argument('-u', '--update', action='store_true')
    args = argParser.parse_args()
    if args.update:
        titanify(pull_data())
    else:
        titanify(pd.read_excel(util.script_folder + 'data/titan_intermediate.xlsx', index_col='sample_id'))
