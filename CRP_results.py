import pandas as pd
import numpy as np
from datetime import date
import util
import argparse
from helpers import try_datediff, permissive_datemax, query_intake, query_research, map_dates
from bisect import bisect_left

def auc_logger(val):
    if val == '-' or str(val).strip() == '' or len(str(val)) > 15:
        return '-'
    elif float(val) == 0.:
        return 0.
    else:
        return np.log2(float(val))


def fallible(row, date_lkp):
    try:
        output = bisect_left(date_lkp[row['participant_id']], pd.to_datetime(row['Date Collected']).date())
    except:
        print("Failed on {} at {}".format(row['participant_id'], row.name))
        print(date_lkp[row['participant_id']])
        print(row['Date Collected'])
        exit(1)
    return output

def crp_results():
    crp_data = (pd.read_excel(util.crp_folder + 'CRP Patient Tracker.xlsx', sheet_name='Tracker', header=4)
                  .dropna(subset=['Umbrella Corresponding Participant ID'])
                  .assign(participant_id = lambda df: df['Umbrella Corresponding Participant ID'].str.strip())
                  .set_index('participant_id')
                  .drop('Umbrella Corresponding Participant ID', axis='columns'))
    participants = crp_data.index.to_numpy()

    sample_info = query_intake(participants=participants)
    sample_info = sample_info[sample_info['Study'] == 'CRP'].copy()
    samples = sample_info.index.to_numpy()
    research_results = query_research(sid_list=samples)

    first_cols = ['Participant ID', 'Date', 'Sample ID', 'Post-Vax', 'Prior Boosts', 'Prior COVID Infections', 'Days to Vaccine #1',
                'Days to 3rd Dose Vaccine', 'Days to Last Infection', 'Days to Last Vax',
                'Qualitative', 'Quantitative', 'Spike endpoint', 'AUC', 'Log2AUC', 'COVID-19 Vaccine Type',
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
    sample_info = (sample_info[~sample_info['Qualitative'].astype(str).str.strip().isin(['-', '']) &
                              ~sample_info['Qualitative'].apply(pd.isna)]
                              .join(research_results)
                              .join(crp_data.loc[:, inf_cols + vax_cols + dem_cols], on='participant_id')
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
                 ['Visit Type', 'Most Recent Infection', 'Most Recent Vax'])
    return sample_info.loc[:, col_order]

if __name__ == '__main__':
    argparser = argparse.ArgumentParser(description='CRP Report')
    argparser.add_argument('-o', '--output_file', action='store', default='tmp', help="Prefix for the output file (current date appended")
    argparser.add_argument('-d', '--debug', action='store_true', help="Print to the command line but do not write to file")
    args = argparser.parse_args()

    report = crp_results()
    output_filename = util.crp_folder + 'dataset_{}_{}.xlsx'.format(args.output_file, date.today().strftime("%m.%d.%y"))
    oneline_cols = ['Post-Vax', 'Prior Boosts', 'Prior COVID Infections']
    if not args.debug:
        with pd.ExcelWriter(output_filename) as writer:
            report.to_excel(writer, sheet_name='Source', index=False)
            report[~report['Post-Vax']].to_excel(writer, sheet_name='Unvaccinated')
            df_first = report[report['Post-Vax']].drop_duplicates(subset='Participant ID', keep='first')
            pd.crosstab(df_first['Prior Boosts'], df_first['Prior COVID Infections']).to_excel(writer, sheet_name='Boosts vs Infs Baseline')
            df_last = report[report['Post-Vax']].drop_duplicates(subset='Participant ID', keep='last')
            pd.crosstab(df_last['Prior Boosts'], df_last['Prior COVID Infections']).to_excel(writer, sheet_name='Boosts vs Infs Latest')
            df_all = report.drop_duplicates(subset='Participant ID', keep='first')
            pd.crosstab([df_all['Post-Vax'], df_all['Prior Boosts']], df_all['Prior COVID Infections']).to_excel(writer, sheet_name='Vax vs Inf Baseline')
        print("CRP report written to {}".format(output_filename))
    print("{} samples from {} participants".format(
        report.shape[0],
        report['Participant ID'].unique().size
    ))