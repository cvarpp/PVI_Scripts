import pandas as pd
import numpy as np
import pickle
import csv
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import date
import datetime
from dateutil import parser
from openpyxl import load_workbook
import util as ut

if __name__ == '__main__':
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    paris_data = pd.read_excel(ut.paris + 'Patient Tracking - PARIS.xlsx', sheet_name='Subgroups', header=4)
    paris_data['Participant ID'] = paris_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = paris_data['Participant ID'].unique()
    paris_data.set_index('Participant ID', inplace=True)
    dems = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Reports & Data/Projects/PARIS/Demographics.xlsx').set_index('Subject ID')
    samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Sample Tracking/Sample Intake Log.xlsx', sheet_name='Sample Intake Log', header=6, dtype=str)
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
    pd.DataFrame(rows).to_excel(ut.paris + 'DataDump.xlsx', index=False)
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
    pd.DataFrame(last_sample).to_excel(ut.paris + 'LastSeen.xlsx', index=False)

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
    research_samples_1 = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Reports & Data/From Krammer Lab/Master Sheet.xlsx', sheet_name='Inputs')
    research_samples_2 = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Reports & Data/From Krammer Lab/Master Sheet.xlsx', sheet_name='Archive')
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
    data = {'Participant ID': [], 'Date': [], 'Sample ID': [], 'Days to 1st Vaccine Dose': [], 'Days to Boost': [], 'Infection Timing': [], 'Qualitative': [], 'Quantitative': [], 'Spike endpoint': [], 'AUC': [], 'Log2AUC': [], 'Vaccine': [], 'Boost Vaccine': [], 'Days to Infection 1': [], 'Days to Infection 2': [], 'Infection Pre-Vaccine?': [], 'Number of SARS-CoV-2 Infections': [], 'Infection on Study': [], '1st Dose Date': [], '2nd Dose Date': [], 'Days to 2nd': [], 'Boost Date': [], 'Boost 2 Date': [], 'Days to Boost 2': [], 'Infection Date 1': [], 'Infection Date 2': [], 'Post-Baseline': [], 'Visit Type': [], 'Sex': [], 'Age': [], 'Race': [], 'Ethnicity: Hispanic or Latino': []}
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
            data['Infection Pre-Vaccine?'].append(paris_data.loc[participant, 'Infection Pre-Vaccine?'])
            data['Sex'].append(dems.loc[participant, 'Gender'])
            data['Age'].append(dems.loc[participant, 'Age'])
            data['Race'].append(dems.loc[participant, 'Race'])
            data['Ethnicity: Hispanic or Latino'].append(dems.loc[participant, 'Ethnicity: Hispanic or Latino'])
            data['Vaccine'].append(paris_data.loc[participant, 'Which Vaccine?'])
            data['Number of SARS-CoV-2 Infections'].append(paris_data.loc[participant, 'Number of SARS-CoV-2 Infections'])
            data['Infection Timing'].append(paris_data.loc[participant, 'Timing'])
            data['Infection on Study'].append(str(paris_data.loc[participant, 'Infection 1 On Study?']).lower() == 'yes' or str(paris_data.loc[participant, 'Infection 2 On Study?']).lower() == 'yes')
            try:
                data['1st Dose Date'].append(paris_data.loc[participant, 'First Dose Date'].date())
            except:
                data['1st Dose Date'].append(paris_data.loc[participant, 'First Dose Date'])
            try:
                data['Days to 1st Vaccine Dose'].append(int((date_ - data['1st Dose Date'][-1]).days))
            except:
                data['Days to 1st Vaccine Dose'].append('')
            try:
                data['2nd Dose Date'].append(paris_data.loc[participant, 'Second Dose Date'].date())
            except:
                data['2nd Dose Date'].append(paris_data.loc[participant, 'Second Dose Date'])
            try:
                data['Days to 2nd'].append(int((date_ - data['2nd Dose Date'][-1]).days))
            except:
                data['Days to 2nd'].append('')
            try:
                if str(paris_data.loc[participant, 'Boosted?']).lower() == 'yes':
                    data['Boost Date'].append(paris_data.loc[participant, 'Booster Date?'].date())
                else:
                    data['Boost Date'].append('')
            except:
                data['Boost Date'].append(paris_data.loc[participant, 'Booster Date?'])
            try:
                data['Days to Boost'].append(int((date_ - data['Boost Date'][-1]).days))
            except:
                data['Days to Boost'].append('')
            try:
                if str(paris_data.loc[participant, 'Yes?']).strip().lower() == 'yes':
                    data['Boost 2 Date'].append(paris_data.loc[participant, '4th Dose Date'].date())
                else:
                    data['Boost 2 Date'].append('')
            except:
                data['Boost 2 Date'].append(paris_data.loc[participant, '4th Dose Date'])
            try:
                data['Days to Boost 2'].append(int((date_ - data['Boost 2 Date'][-1]).days))
            except:
                data['Days to Boost 2'].append('')
            try:
                data['Infection Date 1'].append(paris_data.loc[participant, 'Infection Date 1'].date())
            except:
                data['Infection Date 1'].append(paris_data.loc[participant, 'Infection Date 1'])
            try:
                data['Days to Infection 1'].append(int((date_ - data['Infection Date 1'][-1]).days))
            except:
                data['Days to Infection 1'].append('')
            try:
                data['Infection Date 2'].append(paris_data.loc[participant, 'Infection Date 2'].date())
            except:
                data['Infection Date 2'].append(paris_data.loc[participant, 'Infection Date 2'])
            try:
                data['Days to Infection 2'].append(int((date_ - data['Infection Date 2'][-1]).days))
            except:
                data['Days to Infection 2'].append('')
            data['Boost Vaccine'].append(paris_data.loc[participant, 'Which Vaccine for Boost?'])
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
                print(res[1])
                print(np.log2(res[1]))
                exit(1)
    report = pd.DataFrame(data)
    report.to_excel(ut.paris + 'datasets/all_results_{}.xlsx'.format(date.today().strftime("%m.%d.%y")), index=False)

