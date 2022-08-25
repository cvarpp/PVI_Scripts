import pandas as pd
import numpy as np
import pickle
import csv
from datetime import date
import datetime
from dateutil import parser
import util

if __name__ == '__main__':
    participants = [str(x).strip().upper() for x in pd.read_excel(util.script_input + 'results_of_interest.xlsx')['Participant ID']]
    samples = pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=6, dtype=str)
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    samplesClean = samples.dropna(subset=['Participant ID'])
    research_samples_1 = pd.read_excel(util.research, sheet_name='Inputs')
    research_samples_2 = pd.read_excel(util.research, sheet_name='Archive')
    research_results = {}
    for _, sample in research_samples_1.iterrows():
        sample_id = str(sample['Sample ID']).strip()
        spike = sample['Spike endpoint']
        if str(spike).strip().upper() == "NEG":
            auc = 1.
        else:
            auc = sample['AUC']
        if not pd.isna(auc):
            research_results[sample_id] = [spike, auc]
    for _, sample in research_samples_2.iterrows():
        sample_id = str(sample['Sample ID']).strip()
        spike = sample['Spike endpoint']
        if str(spike).strip().upper() == "NEG":
            auc = 1.
        else:
            auc = sample['AUC']
        if not pd.isna(auc) and sample_id not in research_results.keys():
            research_results[sample_id] = [spike, auc]
    participant_samples = {participant: [] for participant in participants}
    for _, sample in samplesClean.iterrows():
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
                participant_samples[participant].append((parser.parse(str(sample['Date Collected'])).date(), sample['Sample ID'], sample[visit_type], sample[newCol], result_new))
            except:
                print(sample['Date Collected'])
                participant_samples[participant].append((parser.parse('1/1/1900').date(), sample['Sample ID'], sample[visit_type], sample[newCol], result_new))
    data = {'Participant ID': [], 'Date': [], 'Post-Baseline': [], 'Sample ID': [], 'Visit Type': [], 'Qualitative': [], 'Quantitative': [], 'Spike endpoint': [], 'AUC': []}
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
            sample_id = str(sample_id).strip()
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
    report = pd.DataFrame(data)
    report.to_excel(util.script_output + 'results_{}.xlsx'.format(date.today().strftime("%m.%d.%y")), index=False)
