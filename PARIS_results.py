import pandas as pd
import numpy as np
from datetime import date
from dateutil import parser
import util
import argparse

def try_date(potential_date):
    try:
        return potential_date.date()
    except:
        return potential_date

def try_datediff(start_date, end_date):
    try:
        return int((end_date - start_date).days)
    except:
        return ''

def paris_results(output):
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    paris_data = pd.read_excel(util.paris + 'Patient Tracking - PARIS.xlsx', sheet_name='Subgroups', header=4)
    paris_data['Participant ID'] = paris_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = paris_data['Participant ID'].unique()
    paris_data.set_index('Participant ID', inplace=True)
    dems = pd.read_excel(util.projects + 'PARIS/Demographics.xlsx').set_index('Subject ID')
    samples = pd.read_excel(util.tracking + 'Sample Intake Log.xlsx', sheet_name='Sample Intake Log', header=6, dtype=str)
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    samplesClean = samples.dropna(subset=['Participant ID'])
    participant_samples = {participant: [] for participant in participants}
    for _, sample in samplesClean.iterrows():
        if len(str(sample['Sample ID'])) != 5:
            continue
        participant_id = sample['Participant ID'].strip().upper()
        if participant_id in participant_samples.keys():
            if sample['Sample ID'] in [x[1] for x in participant_samples[participant_id]]:
                continue
            try:
                participant_samples[participant_id].append([parser.parse(str(sample['Date Collected'])).date(), sample['Sample ID']])
            except:
                print(sample['Date Collected'])
                participant_samples[participant_id].append([parser.parse('1/1/1900').date(), sample['Sample ID']])
    num_samples = max((len(samples) for samples in participant_samples.values()))
    rows = {'Participant ID': []}
    for i in range(num_samples):
        rows['Date_{}'.format(i)] = []
        rows['SampleID_{}'.format(i)] = []
    for participant, samples in participant_samples.items():
        try:
            samples.sort(key=lambda val: val[0])
        except:
            print(participant)
            print(samples)
            continue
        rows['Participant ID'].append(participant)
        for i in range(num_samples):
            if i < len(samples):
                rows['Date_{}'.format(i)].append(samples[i][0])
                rows['SampleID_{}'.format(i)].append(str(samples[i][1]).strip())
            else:
                rows['Date_{}'.format(i)].append('')
                rows['SampleID_{}'.format(i)].append('')
    pd.DataFrame(rows).to_excel(util.paris + 'DataDump.xlsx', index=False)
    last_sample = {'Participant ID': [], 'Date': [], 'Sample ID': []}
    for participant, samples in participant_samples.items():
        try:
            samples.sort(key=lambda val: val[0])
        except:
            print(participant)
            continue
        last_sample['Participant ID'].append(participant)
        if len(samples) > 0:
            last_sample['Date'].append(samples[-1][0])
            last_sample['Sample ID'].append(samples[-1][1])
        else:
            last_sample['Date'].append('N/A')
            last_sample['Sample ID'].append('No baseline')
    pd.DataFrame(last_sample).to_excel(util.paris + 'LastSeen.xlsx', index=False)

    participant_samples = {participant: [] for participant in participants}
    for _, sample in samplesClean.iterrows():
        if len(str(sample['Sample ID'])) != 5:
            continue
        participant = str(sample['Participant ID']).strip().upper()
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
                participant_samples[participant].append((parser.parse(str(sample['Date Collected'])).date(), str(sample['Sample ID']).strip().upper(), sample[visit_type], sample[newCol], result_new))
            except:
                print(sample['Date Collected'])
                participant_samples[participant].append((parser.parse('1/1/1900').date(), str(sample['Sample ID']).strip().upper(), sample[visit_type], sample[newCol], result_new))
    research_samples_1 = pd.read_excel(util.research , sheet_name='Inputs')
    research_samples_2 = pd.read_excel(util.research , sheet_name='Archive')
    research_results = {}
    for _, sample in research_samples_1.iterrows():
        sample_id = str(sample['Sample ID']).strip().upper()
        spike = sample['Spike endpoint']
        if str(spike).strip().upper() == "NEG":
            auc = 1.
        else:
            auc = sample['AUC']
        if not pd.isna(auc):
            research_results[sample_id] = [spike, auc]
    for _, sample in research_samples_2.iterrows():
        sample_id = str(sample['Sample ID']).strip().upper()
        spike = sample['Spike endpoint']
        if str(spike).strip().upper() == "NEG":
            auc = 1.
        else:
            auc = sample['AUC']
        if not pd.isna(auc) and sample_id not in research_results.keys():
            research_results[sample_id] = [spike, auc]
    col_order = ['Participant ID', 'Date', 'Sample ID', 'Days to 1st Vaccine Dose', 'Days to Boost', 'Infection Timing', 'Qualitative', 'Quantitative', 'Spike endpoint', 'AUC', 'Log2AUC', 'Vaccine Type', 'Boost Type', 'Days to Infection 1', 'Days to Infection 2', 'Infection Pre-Vaccine?', 'Number of SARS-CoV-2 Infections', 'Infection on Study', 'First Dose Date', 'Second Dose Date', 'Days to 2nd', 'Boost Date', 'Boost 2 Date', 'Days to Boost 2', 'Boost 2 Type', 'Infection 1 Date', 'Infection 2 Date', 'Post-Baseline', 'Visit Type', 'Gender', 'Age', 'Race', 'Ethnicity: Hispanic or Latino']
    data = {col: [] for col in col_order}
    dem_cols = ['Gender', 'Age', 'Race', 'Ethnicity: Hispanic or Latino']
    shared_cols = ['Infection Pre-Vaccine?', 'Vaccine Type', 'Number of SARS-CoV-2 Infections', 'Infection Timing', 'Boost Type', 'Boost 2 Type']
    date_cols = ['First Dose Date', 'Second Dose Date', 'Boost Date', 'Boost 2 Date', 'Infection 1 Date', 'Infection 2 Date']
    day_cols = ['Days to 1st Vaccine Dose', 'Days to 2nd', 'Days to Boost', 'Days to Boost 2', 'Days to Infection 1', 'Days to Infection 2']
    for participant, samples in participant_samples.items():
        try:
            samples.sort(key=lambda x: x[0])
        except:
            print(participant)
            print(samples)
            continue
        if len(samples) < 1:
            continue
        baseline = samples[0][0]
        for date_, sample_id, visit_type, result, result_new in samples:
            if result == '-' or str(result).strip() == '' or (type(result) == float and pd.isna(result)):
                continue
            sample_id = str(sample_id).strip()
            for dem_col in dem_cols:
                data[dem_col].append(dems.loc[participant, dem_col])
            for shared_col in shared_cols:
                data[shared_col].append(paris_data.loc[participant, shared_col])
            data['Infection on Study'].append(str(paris_data.loc[participant, 'Infection 1 On Study?']).strip().lower() == 'yes' or str(paris_data.loc[participant, 'Infection 2 On Study?']).strip().lower() == 'yes')
            for date_col, day_col in zip(date_cols, day_cols):
                data[date_col].append(try_date(paris_data.loc[participant, date_col]))
                data[day_col].append(try_datediff(data[date_col][-1], date_))
            data['Participant ID'].append(participant)
            data['Date'].append(date_)
            data['Post-Baseline'].append((date_ - baseline).days)
            data['Sample ID'].append(sample_id)
            data['Visit Type'].append(visit_type)
            data['Qualitative'].append(result)
            data['Quantitative'].append(result_new)
            if sample_id in research_results.keys():
                res = research_results[sample_id]
            else:
                res = ['-', '-']
            data['Spike endpoint'].append(res[0])
            data['AUC'].append(res[1])
            try:
                if res[1] == '-' or str(res[1]).strip() == '' or len(str(res[1])) > 15:
                    data['Log2AUC'].append('-')
                elif float(res[1]) == 0.:
                    data['Log2AUC'].append(0.)
                else:
                    data['Log2AUC'].append(np.log2(float(res[1])))
            except:
                print("Log transformation on {} for {} failed. Fatal error, exiting".format(res[1], sample_id))
                exit(1)
    report = pd.DataFrame(data)
    output_filename = util.paris + 'datasets/{}_{}.xlsx'.format(output, date.today().strftime("%m.%d.%y"))
    report.to_excel(output_filename, index=False)
    print("PARIS report written to {}".format(output_filename))
    print("{} samples from {} participants".format(
        report.shape[0],
        report['Participant ID'].unique().size
    ))

if __name__ == '__main__':
    argparser = argparse.ArgumentParser(description='Paris reporting generation')
    argparser.add_argument('-o', '--output_file', action='store', default='tmp', help="What would you like the file to be called")
    arg = argparser.parse_args()

    paris_results(arg.output_file)