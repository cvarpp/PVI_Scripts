import pandas as pd
import numpy as np
from datetime import date
from dateutil import parser
import util
import argparse
from helpers import clean_research, try_datediff, permissive_datemax

def paris_results():
    paris_data = pd.read_excel(util.paris + 'Patient Tracking - PARIS.xlsx', sheet_name='Subgroups', header=4).dropna(subset=['Participant ID'])
    paris_data['Participant ID'] = paris_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = paris_data['Participant ID'].unique()
    paris_data.set_index('Participant ID', inplace=True)
    dems = pd.read_excel(util.projects + 'PARIS/Demographics.xlsx').set_index('Subject ID')
    samples = pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=util.header_intake, dtype=str)
    samplesClean = samples.dropna(subset=['Participant ID'])
    participant_samples = {participant: [] for participant in participants}
    for _, sample in samplesClean.iterrows():
        if len(str(sample['Sample ID'])) != 5:
            continue
        participant = str(sample['Participant ID']).strip().upper()
        if participant in participant_samples.keys():
            if str(sample[util.qual]).strip().upper() == "NEGATIVE":
                sample[util.quant] = "Negative"
            if pd.isna(sample[util.quant]):
                result_new = '-'
            elif type(sample[util.quant]) == int:
                result_new = sample[util.quant]
            elif str(sample[util.quant]).strip().upper() == "NEGATIVE":
                result_new = 1.
            else:
                result_new = sample[util.quant]
            try:
                participant_samples[participant].append((parser.parse(str(sample['Date Collected'])).date(), str(sample['Sample ID']).strip().upper(), sample[util.visit_type], sample[util.qual], result_new))
            except:
                print(sample['Sample ID'], "has invalid date", sample['Date Collected'])
                participant_samples[participant].append((parser.parse('1/1/1900').date(), str(sample['Sample ID']).strip().upper(), sample[util.visit_type], sample[util.qual], result_new))

    research_source = pd.ExcelFile(util.research)
    research_samples_1 = research_source.parse(sheet_name='Inputs').pipe(clean_research)
    research_samples_2 = research_source.parse(sheet_name='Archive').pipe(clean_research)
    research_results = pd.concat([research_samples_2, research_samples_1]).drop_duplicates(subset=['sample_id'], keep='last').set_index('sample_id')

    col_order = ['Participant ID', 'Date', 'Sample ID', 'Days to 1st Vaccine Dose', 'Days to Boost', 'Days to Last Infection', 'Days to Last Vax', 'Infection Timing', 'Qualitative', 'Quantitative', 'Spike endpoint', 'AUC', 'Log2AUC', 'Vaccine Type', 'Boost Type', 'Days to Infection 1', 'Days to Infection 2', 'Days to Infection 3', 'Infection Pre-Vaccine?', 'Number of SARS-CoV-2 Infections', 'Infection on Study', 'First Dose Date', 'Second Dose Date', 'Days to 2nd', 'Boost Date', 'Boost 2 Date', 'Days to Boost 2', 'Boost 2 Type', 'Boost 3 Date', 'Days to Boost 3', 'Boost 3 Type', 'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Most Recent Infection', 'Most Recent Vax', 'Post-Baseline', 'Visit Type', 'Gender', 'Age', 'Race', 'Ethnicity: Hispanic or Latino']
    data = {col: [] for col in col_order}
    dem_cols = ['Gender', 'Age', 'Race', 'Ethnicity: Hispanic or Latino']
    shared_cols = ['Infection Pre-Vaccine?', 'Vaccine Type', 'Number of SARS-CoV-2 Infections', 'Infection Timing', 'Boost Type', 'Boost 2 Type', 'Boost 3 Type']
    date_cols = ['First Dose Date', 'Second Dose Date', 'Boost Date', 'Boost 2 Date', 'Boost 3 Date', 'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date']
    day_cols = ['Days to 1st Vaccine Dose', 'Days to 2nd', 'Days to Boost', 'Days to Boost 2', 'Days to Boost 3', 'Days to Infection 1', 'Days to Infection 2', 'Days to Infection 3']
    dose_dates = ['First Dose Date', 'Second Dose Date', 'Boost Date', 'Boost 2 Date', 'Boost 3 Date']
    inf_dates = ['Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date']
    for participant, samples in participant_samples.items():
        try:
            samples.sort(key=lambda x: x[0])
        except:
            print(participant, "has", samples, "samples that do not sort. Skipping for report...")
            continue
        if len(samples) < 1:
            print(participant, "has no samples. Skipping for report...")
            continue
        baseline = samples[0][0]
        for date_, sample_id, util.visit_type, result, result_new in samples:
            if result == '-' or str(result).strip() == '' or (type(result) == float and pd.isna(result)):
                continue
            sample_id = str(sample_id).strip()
            for dem_col in dem_cols:
                data[dem_col].append(dems.loc[participant, dem_col])
            for shared_col in shared_cols:
                data[shared_col].append(paris_data.loc[participant, shared_col])
            data['Infection on Study'].append(str(paris_data.loc[participant, 'Infection 1 On Study?']).strip().lower() == 'yes' or str(paris_data.loc[participant, 'Infection 2 On Study?']).strip().lower() == 'yes' or str(paris_data.loc[participant, 'Infection 3 On Study?']).strip().lower() == 'yes')
            for date_col, day_col in zip(date_cols, day_cols):
                data[date_col].append(pd.to_datetime(paris_data.loc[participant, date_col], errors='coerce').date())
                data[day_col].append(try_datediff(data[date_col][-1], date_))
            data['Most Recent Infection'].append(permissive_datemax([data[inf_date][-1] for inf_date in inf_dates], date_))
            data['Most Recent Vax'].append(permissive_datemax([data[dose_date][-1] for dose_date in dose_dates], date_))
            data['Days to Last Infection'].append(try_datediff(data['Most Recent Infection'][-1], date_))
            data['Days to Last Vax'].append(try_datediff(data['Most Recent Vax'][-1], date_))
            data['Participant ID'].append(participant)
            data['Date'].append(date_)
            data['Post-Baseline'].append((date_ - baseline).days)
            data['Sample ID'].append(sample_id)
            data['Visit Type'].append(util.visit_type)
            data['Qualitative'].append(result)
            data['Quantitative'].append(result_new)
            if sample_id in research_results.index:
                res = [research_results.loc[sample_id, 'Spike endpoint'], research_results.loc[sample_id, 'AUC']]
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
    (report.loc[:, ['Participant ID', 'Date', 'Sample ID']]
           .drop_duplicates(subset=['Participant ID'], keep='last')
           .to_excel(util.paris + 'LastSeen.xlsx', index=False))
    return report

if __name__ == '__main__':
    argparser = argparse.ArgumentParser(description='Paris reporting generation')
    argparser.add_argument('-o', '--output_file', action='store', default='tmp', help="Prefix for the output file (current date appended")
    argparser.add_argument('-d', '--debug', action='store_true', help="Print to the command line but do not write to file")
    args = argparser.parse_args()

    report = paris_results()
    if not args.debug:
        output_filename = util.paris + 'datasets/{}_{}.xlsx'.format(args.output_file, date.today().strftime("%m.%d.%y"))
        report.to_excel(output_filename, index=False)
        print("PARIS report written to {}".format(output_filename))
    print("{} samples from {} participants".format(
        report.shape[0],
        report['Participant ID'].unique().size
    ))