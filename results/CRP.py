import pandas as pd
import numpy as np
from datetime import date
import util
import argparse
import sys
import PySimpleGUI as sg
from helpers import try_datediff, permissive_datemax, query_intake, query_research, map_dates, query_dscf, ValuesToClass
from bisect import bisect_left

def auc_logger(val):
    if val == '-' or str(val).strip() == '' or len(str(val)) > 15:
        return '-'
    elif float(val) == 0.:
        return 0.
    else:
        return np.log2(float(val))

def find_last_event(row):
    if pd.isna(row['Days to Last Vax']) and pd.isna(row['Days to Last Infection']):
        return "No Vaccine or Infection"
    elif pd.isna(row['Days to Last Infection']) or row['Days to Last Vax'] < row['Days to Last Infection']:
        return "Vax"
    else:
        return "Infection"

def crp_results():
    crp_data = (pd.read_excel(util.crp_folder + 'CRP Patient Tracker.xlsx', sheet_name='Tracker', header=4)
                  .dropna(subset=['Participant ID'])
                  .assign(participant_id = lambda df: df['Participant ID'].str.strip())
                  .set_index('participant_id')
                  .drop('Participant ID', axis='columns'))
    participants = crp_data.index.to_numpy()

    sample_info = query_intake(participants=participants)
    sample_info = sample_info[sample_info['Study'] == 'CRP'].copy()
    samples = sample_info.index.to_numpy()
    research_results = query_research(sid_list=samples)
    proc_cols = ['Volume of Serum Collected (mL)', 'Total volume of plasma (mL)', 'PBMC concentration per mL (x10^6)', 'viability', '# of PBMC vials', 'cpt_vol', 'sst_vol', 'proc_comment']
    proc_data = query_dscf(sid_list=samples).loc[:, proc_cols]

    first_cols = ['Participant ID', 'Date', 'Sample ID', 'last_event', 'Post-Vax', 'Prior Boosts', 'Prior COVID Infections', 'Days to Vaccine #1',
                'Days to 3rd Dose Vaccine', 'Days to Last Infection', 'Days to Last Vax',
                'Qualitative', 'Quantitative', 'Spike endpoint', 'AUC', 'Log2AUC', 'Has Path?', 'Has AUC?', 'COVID-19 Vaccine Type',
                '3rd Dose Vaccine Type', 'Days to Infection 1', 'Days to Infection 2', 'Days to Infection 3']
    dem_cols = ['Sex', 'Gender', 'Age at Visit', 'Race', 'Ethnicity']
    vax_cols = ['COVID-19 Vaccine Type', 'Vaccine #1 Date', 'Vaccine #2 Date', '3rd Dose? or Booster?',
                '3rd Dose Vaccine Type', '3rd Dose Vaccine Date', '4th Dose Vaccine Type',
                '4th Dose Date', '5th Dose Vaccine Type', '5th Dose Date']
    inf_cols = ['Positive Test COVID-19?', 'How Many?', 'Infection 1 Date', 'Infection 1 Test Type',
                'Infection 2 Date', 'Infection 2 Test Type', 'Infection 3 Date', 'Infection 3 Test Type']
    dose_dates = [col for col in vax_cols if 'Date' in col]
    inf_dates = [col for col in inf_cols if 'Date' in col]
    sample_info['Date'] = sample_info['Date Collected']
    intake_drops = ['Date Collected', 'Date Shared', 'Clinical Ab Result Shared?', 'Shared By']
    sample_info = (sample_info.join(research_results)
                              .join(crp_data.loc[:, inf_cols + vax_cols + dem_cols], on='participant_id')
                              .join(proc_data)
                              .drop(intake_drops, axis='columns')
                              .sort_values(by=['participant_id', 'Date'])
                              .reset_index().copy())
    baseline_data = sample_info.drop_duplicates(subset='participant_id').set_index('participant_id')
    sample_info['Days Post-Baseline'] = (sample_info['Date'] - sample_info['participant_id'].apply(lambda val: baseline_data.loc[val, 'Date'])).dt.days
    date_cols = [col for col in sample_info.columns if 'Date' in col]
    sample_info = sample_info.pipe(map_dates, date_cols)
    for date_col in date_cols:
        if date_col not in ['Date']:
            day_col = 'Days to ' + date_col[:-5]
            sample_info[day_col] = sample_info.apply(lambda row: try_datediff(row[date_col], row['Date']), axis=1)
    sample_info['Most Recent Infection'] = sample_info.apply(lambda row: permissive_datemax([row[inf_date] for inf_date in inf_dates], row['Date']), axis=1)
    sample_info['Most Recent Vax'] = sample_info.apply(lambda row: permissive_datemax([row[dose_date] for dose_date in dose_dates], row['Date']), axis=1)
    sample_info['Days to Last Infection'] = sample_info.apply(lambda row: try_datediff(row['Most Recent Infection'], row['Date']), axis=1)
    sample_info['Days to Last Vax'] = sample_info.apply(lambda row: try_datediff(row['Most Recent Vax'], row['Date']), axis=1)
    sample_info['Participant ID'] = sample_info['participant_id']
    sample_info['Sample ID'] = sample_info['sample_id']
    sample_info['Visit Type'] = sample_info[util.visit_type]
    sample_info['Log2AUC'] = sample_info['AUC'].apply(auc_logger)
    sample_info['Has Path?'] = sample_info['Quantitative'].apply(lambda val: "Yes" if not pd.isna(val) else "No")
    sample_info['Has AUC?'] = sample_info['AUC'].apply(lambda val: "Yes" if not pd.isna(val) else "No")


    pre_infs = []
    for inf_dt in ['Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date']:
        pre_col = 'Pre ' + inf_dt[:-5]
        sample_info[pre_col] = sample_info['Date'] > sample_info[inf_dt]
        pre_infs.append(pre_col)
    sample_info['Prior COVID Infections'] = sample_info.loc[:, pre_infs].sum(axis='columns')

    pre_boosts = []
    for boost_dt in ['3rd Dose Vaccine Date', '4th Dose Date', '5th Dose Date']:
        pre_col = 'Pre ' + boost_dt[:8]
        sample_info[pre_col] = sample_info['Date'] > sample_info[boost_dt]
        pre_boosts.append(pre_col)
    sample_info['Prior Boosts'] = sample_info.loc[:, pre_boosts].sum(axis='columns')

    assert not any(row['Days to Last Vax'] > 0 and row['Days to Vaccine #2'] < 0 for _, row in sample_info.iterrows())
    sample_info['Post-Vax'] = sample_info['Days to Last Vax'] > 0

    col_order = (first_cols +
                 [col for col in sample_info.columns if 'Days to' in col and col not in first_cols] +
                 [col for col in vax_cols if col not in first_cols] +
                 [col for col in inf_cols if col not in first_cols] +
                 dem_cols +
                 ['Visit Type', 'Most Recent Infection', 'Most Recent Vax'] +
                 proc_cols)
    if sample_info['Participant ID'].unique().size < participants.size:
        print("Participants with no samples in report:")
        print([p for p in participants if p not in sample_info['Participant ID'].to_numpy()])
    return sample_info.assign(last_event=lambda df: df.apply(find_last_event, axis=1)).loc[:, col_order].assign(total_pbmcs=lambda df: df['PBMC concentration per mL (x10^6)'] * df['# of PBMC vials'])

if __name__ == '__main__':
    if len(sys.argv) != 1:
        argparser = argparse.ArgumentParser(description='Generate report for all CRP samples')
        argparser.add_argument('-o', '--output_file', action='store', default='tmp', help="Prefix for the output file (in addition to current date)")
        argparser.add_argument('-d', '--debug', action='store_true', help="Print to the command line but do not write to file")
        args = argparser.parse_args()
    else:
        sg.theme('Dark Blue 17')

        layout = [[sg.Text('CRP Result Generation Script')],
                  [sg.Checkbox('Debug?', key='debug', default=False)],
                  [sg.Text('Output File Name:'),sg.Input(key='output_file')],
                  [sg.Submit(), sg.Cancel()]]
        
        window = sg.Window('CRP Results Script', layout)

        event,  values = window.read()
        window.close()

        if event=='Cancel':
            quit()
        else:
            args = ValuesToClass(values)

    report = crp_results()
    converter = pd.read_excel(util.tracking + 'AUC Converter.xlsx', sheet_name='Key').set_index('Spike Endpoint')
    report['False Path Value'] = report.apply(lambda row: row['AUC'] / converter.loc[row['Spike endpoint'], 'Corrective Factor'] if row['Spike endpoint'] in converter.index else np.nan, axis=1)
    report['Path Plus'] = report.apply(lambda row: row['Quantitative'] if not pd.isna(row['Quantitative']) else row['False Path Value'], axis=1)
    output_filename = util.crp_folder + 'dataset_{}_{}.xlsx'.format(args.output_file, date.today().strftime("%m.%d.%y"))
    oneline_cols = ['Post-Vax', 'Prior Boosts', 'Prior COVID Infections']
    just_proc = ['Participant ID', 'Sample ID', 'Date',
                 'Volume of Serum Collected (mL)', 'Total volume of plasma (mL)', 'total_pbmcs',
                 'PBMC concentration per mL (x10^6)', 'viability', '# of PBMC vials',
                 'cpt_vol', 'sst_vol', 'proc_comment']
    if not args.debug:
        with pd.ExcelWriter(output_filename) as writer:
            report.to_excel(writer, sheet_name='Source', index=False)
            report.loc[:, just_proc].to_excel(writer, sheet_name='Processing Data', index=False)
            report[~report['Post-Vax']].to_excel(writer, sheet_name='Unvaccinated', index=False)
            df_all = report.drop_duplicates(subset='Participant ID', keep='first')
            last_vax_valencies = []
            for _, row in df_all.iterrows():
                valency = ['Monovalent']
                for val in [row['3rd Dose Vaccine Type'], row['4th Dose Vaccine Type'], row['5th Dose Vaccine Type']]:
                    if 'bivalent' in str(val).lower():
                        valency.append('Bivalent')
                    else:
                        valency.append('Monovalent')
                last_vax_valencies.append(valency[row['Prior Boosts']])
            df_all['Last Vax Valency'] = last_vax_valencies
            df_all['Prior Boosts + Valency'] = df_all['Prior Boosts'].astype(str) + ': ' + df_all['Last Vax Valency']
            pd.crosstab([df_all['Post-Vax'], df_all['Prior Boosts']], df_all['Prior COVID Infections']).to_excel(writer, sheet_name='Vax vs Inf Baseline')
            pd.crosstab([df_all['Post-Vax'], df_all['Prior Boosts + Valency']], df_all['Prior COVID Infections']).to_excel(writer, sheet_name='Vax vs Inf Baseline Bivalent')
            pd.crosstab([df_all['Post-Vax'], df_all['Prior Boosts + Valency']], [df_all['last_event'], df_all['Prior COVID Infections']]).to_excel(writer, sheet_name='Vax vs Inf Base Biv Col')
            for ycol in ['AUC', 'Quantitative', 'Path Plus']:
                df_local = df_all[~df_all[ycol].apply(pd.isna)]
                pd.crosstab([df_local['Post-Vax'], df_local['Prior Boosts + Valency']], [df_local['last_event'], df_local['Prior COVID Infections']]).to_excel(writer, sheet_name='Vax vs Inf Base Biv Col {}'.format(ycol[:5]))
            pd.crosstab([df_all['last_event'], df_all['Post-Vax'], df_all['Prior Boosts + Valency']], df_all['Prior COVID Infections']).to_excel(writer, sheet_name='Vax vs Inf Base Biv Row')
            for last_event in ['Vax', 'Infection']:
                df_local = df_all[df_all['last_event'] == last_event]
                pd.crosstab([df_local['Post-Vax'], df_local['Prior Boosts + Valency']], df_local['Prior COVID Infections']).to_excel(writer, sheet_name='Vax vs Inf Base Biv {}'.format(last_event))
            for last_event, last_ev_col in zip(['Vax', 'Infection'], ['Days to Last Vax', 'Days to Last Infection']):
                df_local = df_all[(df_all['last_event'] == last_event) & ~df_all['AUC'].apply(pd.isna)].drop_duplicates(subset='Participant ID')
                for col in ['Prior Boosts', 'Prior COVID Infections']:
                    dfs = {}
                    for num in [0, 1, 2, 3]:
                        dfs[num] = df_local[df_local[col] == num].loc[:, ['Participant ID', last_ev_col, 'AUC']].rename(columns={'AUC': '{} {}'.format(num, col[6:])}).set_index(['Participant ID', last_ev_col])
                    df_output = dfs[0].join(dfs[1], how='outer').join(dfs[2], how='outer').join(dfs[3], how='outer').reset_index()
                    df_output.to_excel(writer, sheet_name="AUC {} by {}".format(last_event, col[:12].strip()), index=False)
            for last_event, last_ev_col in zip(['Vax', 'Infection'], ['Days to Last Vax', 'Days to Last Infection']):
                df_local = df_all[(df_all['last_event'] == last_event) & ~df_all['Quantitative'].apply(pd.isna)].drop_duplicates(subset='Participant ID')
                for col in ['Prior Boosts', 'Prior COVID Infections']:
                    dfs = {}
                    for num in [0, 1, 2, 3]:
                        dfs[num] = df_local[df_local[col] == num].loc[:, ['Participant ID', last_ev_col, 'Quantitative']].rename(columns={'Quantitative': '{} {}'.format(num, col[6:])}).set_index(['Participant ID', last_ev_col])
                    df_output = dfs[0].join(dfs[1], how='outer').join(dfs[2], how='outer').join(dfs[3], how='outer').reset_index()
                    df_output.to_excel(writer, sheet_name="Path {} by {}".format(last_event, col[:12].strip()), index=False)
            for last_event, last_ev_col in zip(['Vax', 'Infection'], ['Days to Last Vax', 'Days to Last Infection']):
                df_local = df_all[(df_all['last_event'] == last_event) & ~df_all['Path Plus'].apply(pd.isna)].drop_duplicates(subset='Participant ID')
                for col in ['Prior Boosts', 'Prior COVID Infections']:
                    dfs = {}
                    for num in [0, 1, 2, 3]:
                        dfs[num] = df_local[df_local[col] == num].loc[:, ['Participant ID', last_ev_col, 'Path Plus']].rename(columns={'Path Plus': '{} {}'.format(num, col[6:])}).set_index(['Participant ID', last_ev_col])
                    df_output = dfs[0].join(dfs[1], how='outer').join(dfs[2], how='outer').join(dfs[3], how='outer').reset_index()
                    df_output.to_excel(writer, sheet_name="Path+ {} by {}".format(last_event, col[:12].strip()), index=False)

        print("CRP report written to {}".format(output_filename))
    print("{} samples from {} participants.\n{} with research results, {} with path results".format(
        report.shape[0],
        report['Participant ID'].unique().size,
        report[report['Has AUC?'] == 'Yes'].shape[0],
        report[report['Has Path?'] == 'Yes'].shape[0]
    ))