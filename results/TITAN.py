import pandas as pd
import numpy as np
import util
import argparse
import sys
import PySimpleGUI as sg
from helpers import try_datediff, query_dscf, query_intake, query_research, map_dates, ValuesToClass

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
    reverse_convert = {v: k for k, v in titan_convert.items()}
    participant_data = (titan_data.rename(columns=reverse_convert)
                            .loc[:, titan_convert.keys()]
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
    all_converters = titan_convert | post3_covid_renamer | post4_covid_renamer | post5_covid_renamer
    date_cols = [col for col in all_converters.keys() if 'Date' in col]
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
    return participant_data.pipe(map_dates, date_cols)

def pull_data():
    participant_data = query_titan()
    participants = [p for p in participant_data.index]
    participant_samples = query_intake(participants).drop(['Order #', 'PVI #'], axis='columns')
    samples_of_interest = participant_samples.index.to_numpy()
    keep_cols = ['Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials']
    sample_info = query_dscf(sid_list=samples_of_interest).loc[:, keep_cols]
    research_results = query_research(sid_list=samples_of_interest)

    df = (participant_samples
            .join(sample_info)
            .join(research_results)
            .join(participant_data, on='participant_id')
            .rename(columns={'Date Collected': 'Date'}))
    date_cols = ['1st Dose Date', '2nd Dose Date', '3rd Dose Date', 'Boost Date', 'Boost 2 Date', 'COVID Pre-Enrollment Date', 'COVID Post-Third Date', 'Monoclonal Post-Third Date', 'COVID Post-Boost Date', 'Monoclonal Post-Boost Date', 'COVID Post-Boost 2 Date', 'Monoclonal Post-Boost 2 Date']
    for date_col in date_cols:
        day_col = 'Days from ' + date_col[:-5]
        df[day_col] = df.apply(lambda row: try_datediff(row[date_col], row['Date']), axis=1)
    df.dropna(subset=['AUC'], inplace=True)
    df.to_excel(util.script_folder + 'data/titan_intermediate.xlsx')
    return df

def titanify(df):
    df['AUC'] = df['AUC'].astype(float)
    df['Log2AUC'] = np.log2(df['AUC'])
    output_fname = util.script_output + 'new_format/titan_consolidated.xlsx'
    df['Timepoint'] = df['Days from 3rd Dose'].apply(make_timepoint)
    df['Boost Timepoint'] = df['Days from Boost'].apply(make_timepoint).apply(lambda val: "Pre-Dose 4" if val == "Pre-Dose 3" else val)
    longform_columns = ['Full ID', 'Transplant', 'COVID', 'HIV', 'Days from 3rd Dose', 'AUC', 'Spike endpoint', 'Log2AUC', 'Full Transplant', 'participant_id', 'AM', 'Prednisone', 'CNI', 'Gender', 'Age at Enrollment', 'Vaccine', 'Boost Vaccine', 'Timepoint', 'Boost Timepoint', 'Days from Boost', 'Days from Boost 2', 'Days from 1st Dose', 'Days from 2nd Dose', 'Days from COVID Pre-Enrollment', 'Days from COVID Post-Third', 'Days from COVID Post-Boost', 'Days from COVID Post-Boost 2', 'Days from Monoclonal Post-Third', 'Days from Monoclonal Post-Boost', 'Days from Monoclonal Post-Boost 2']
    df_long = df.loc[:, longform_columns].sort_values(by=['Full ID', 'Days from 3rd Dose'])
    tp_order = ['Pre-Dose 3', 'Day 30', 'Day 90', 'Day 180', 'Day 300']
    tp_order_boost = ['Pre-Dose 4', 'Day 30', 'Day 90', 'Day 180', 'Day 300']
    df_wide = (df[df['Timepoint'] != 'None']
                .drop_duplicates(subset=['Full ID', 'Timepoint'], keep='last')
                .pivot_table(values='AUC', index='Full ID', columns='Timepoint')
                .reindex(tp_order, axis=1))
    # print(df_long.loc[:, ['Days from ']])
    ppl_key = df_long.drop_duplicates(subset=['Full ID']).set_index('Full ID')
    df_wide_annot = df_wide.copy()
    for col in ['Transplant', 'COVID', 'HIV', 'AM', 'Prednisone']:
        df_wide_annot[col] = df_wide_annot.apply(lambda row: ppl_key.loc[row.name, col], axis=1)
    tp_sample_ids = (df[df['Timepoint'] != 'None'].reset_index()
                .drop_duplicates(subset=['Full ID', 'Timepoint'], keep='last')
                .pivot_table(values='sample_id', index='Full ID', columns='Timepoint', aggfunc=np.sum)
                .reindex(tp_order, axis=1))
    df_pre = df[df['Boost Timepoint'] == "Pre-Dose 4"]
    df_post = df[df['Boost Timepoint'] != "Pre-Dose 4"]
    if not args.debug:
        with pd.ExcelWriter(output_fname) as writer:
            df.to_excel(writer, sheet_name='Source')
            df_long.to_excel(writer, sheet_name='Long-Form')
            df_wide.to_excel(writer, sheet_name='Wide-Form')
            df_wide_annot.to_excel(writer, sheet_name='Wide-Form Annotated')
            tp_sample_ids.to_excel(writer, sheet_name='Wide-Form Sample ID Key')
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
        print("Data written to {}".format(output_fname))
    print("{} samples from {} participants.".format(
        df_long.shape[0],
        df_long['participant_id'].unique().size
    ))

if __name__ == '__main__':
    if len(sys.argv) !=1:

        argParser = argparse.ArgumentParser(description='Generate report for TITAN samples')
        argParser.add_argument('-c', '--use_cache', action='store_true', help='Use cached results instead of pulling from source sheets')
        argParser.add_argument('-d', '--debug', action='store_true', help="Print to the command line but do not write to file")
        args = argParser.parse_args()
    
    else:
        sg.theme('Dark Blue 17')

        layout = [[sg.Text('TITAN Result Generation Script')],
                  [sg.Checkbox('Use Cache?', key='use_cache', default=False)],
                  [sg.Checkbox('Debug?', key='debug', default=False)],
                  [sg.Submit(), sg.Cancel()]]
        
        window = sg.Window('Titan Results Script', layout)

        event,  values = window.read()
        window.close()

        if event=='Cancel':
            quit()
        else:
            args = ValuesToClass(values)

    if not args.use_cache:
        titanify(pull_data())
    else:
        titanify(pd.read_excel(util.script_folder + 'data/titan_intermediate.xlsx', index_col='sample_id'))
