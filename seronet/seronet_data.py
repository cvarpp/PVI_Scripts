import pandas as pd
import numpy as np
import pickle
import csv
from datetime import date
import datetime
from dateutil import parser
from dateutil.relativedelta import relativedelta
from bisect import bisect_left
import pickle
import sys

if __name__ == '__main__':
    script_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Scripts/'
    script_output = script_folder + 'output/'
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    iris_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/IRIS/'
    iris_data = pd.read_excel(iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID'])
    iris_data['Participant ID'] = iris_data['Participant ID'].apply(lambda val: val.strip().upper())
    titan_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/TITAN/'
    titan_data = pd.read_excel(titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_data['Participant ID'] = titan_data['Umbrella Corresponding Participant ID'].apply(lambda val: val.strip().upper())
    visit_type = "Visit Type / Samples Needed"
    mars_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/MARS/'
    mars_data = pd.read_excel(mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID'])
    mars_data['Participant ID'] = mars_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = [p for p in mars_data['Participant ID'].unique()] + [p for p in titan_data['Participant ID'].unique()] + [p for p in iris_data['Participant ID'].unique()]
    participant_study = {p: 'MARS' for p in mars_data['Participant ID'].unique()}
    participant_study.update({p: 'TITAN' for p in titan_data['Participant ID'].unique()})
    participant_study.update({p: 'IRIS' for p in iris_data['Participant ID'].unique()})
    mars_data.set_index('Participant ID', inplace=True)
    iris_data.set_index('Participant ID', inplace=True)
    titan_data.set_index('Participant ID', inplace=True)
    samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Sample Tracking/Sample Intake Log.xlsx', sheet_name='Sample Intake Log', header=6, dtype=str)
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    samplesClean = samples.dropna(subset=['Participant ID'])
    participant_samples = {participant: [] for participant in participants}# if str(participant).upper()[:4] != 'CITI'}
    for _, sample in samplesClean.iterrows():
        if len(str(sample['Sample ID'])) != 5:
            continue
        participant = str(sample['Participant ID']).strip().upper()
        if str(sample['Study']).strip().upper() == 'PRIORITY':
            if participant not in participant_samples.keys():
                participant_samples[participant] = []
                participant_study[participant] = 'PRIORITY'
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
                sample_date = parser.parse(str(sample['Date Collected'])).date()
            except:
                print(sample['Date Collected'])
                sample_date = parser.parse('1/1/1900').date()
                continue # because we're in TIMP, we don't want dateless samples
            if participant_study[participant] == "MARS":
                if type(mars_data.loc[participant, 'Vaccine #1 Date']) != datetime.datetime or (
                    type(mars_data.loc[participant, 'Vaccine #2 Date']) != datetime.datetime
                ):
                    print("Try to find vaccine dates for ", participant)
                elif mars_data.loc[participant, 'Vaccine #1 Date'].date() < sample_date and (
                     mars_data.loc[participant, 'Vaccine #2 Date'].date() >= sample_date):
                    continue
            if participant_study[participant] == "IRIS":
                if type(iris_data.loc[participant, 'First Dose Date']) != datetime.datetime or (
                    type(iris_data.loc[participant, 'Second Dose Date']) != datetime.datetime
                ):
                    print("Try to find vaccine dates for ", participant)
                elif iris_data.loc[participant, 'First Dose Date'].date() < sample_date and (
                    iris_data.loc[participant, 'Second Dose Date'].date() >= sample_date):
                    continue
            participant_samples[participant].append((sample_date, sample['Sample ID'], sample[visit_type], sample[newCol], result_new))
            
    research_samples_1 = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Reports & Data/From Krammer Lab/Master Sheet.xlsx', sheet_name='Inputs')
    research_samples_2 = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Reports & Data/From Krammer Lab/Master Sheet.xlsx', sheet_name='Archive')
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
    bsl2p_samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Processing Team/Data Sample Collection Form.xlsx', sheet_name='BSL2+ Samples', header=1)
    bsl2_samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Processing Team/Data Sample Collection Form.xlsx', sheet_name='BSL2 Samples')
    shared_samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Sample Tracking/Released Samples/Collaborator Samples Tracker.xlsx', sheet_name='Released Samples')
    no_pbmcs = set([str(sid).strip() for sid in shared_samples[shared_samples['Sample Type'] == 'PBMC']['Sample ID'].unique()])
    missing_info = {'Sample ID': [], 'Serum Volume': [], 'PBMC Conc': [], 'PBMC Vial Count': []}
    if len(sys.argv) == 1 or sys.argv[1] != 'maintain_volumes':
        sample_volumes = {}
        for _, sample in bsl2p_samples.iterrows():
            sample_id = str(sample['Sample ID']).strip()
            try:
                serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("mlML ").split()[0])
            except:
                try:
                    serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("ulUL ").split()[0]) / 1000.
                except:
                    print(sample_id, serum_volume, type(serum_volume))
                    if type(sample['Total volume of serum (ml)']) == str:
                        serum_volume = 0
                    else:
                        serum_volume = sample['Total volume of serum (ml)']
            if type(serum_volume) != str and serum_volume > 4.5:
                serum_volume = 4.5
            if sample_id in no_pbmcs:
                pbmc_conc = 0
                vial_count = 0
            else:
                pbmc_conc = sample['# cells per aliquot']
                vial_count = sample['# of aliquots frozen']
                if (type(pbmc_conc) != float or pd.isna(pbmc_conc)) and type(pbmc_conc) != int:
                    pbmc_conc = 0
                    vial_count = 0
            if type(vial_count) in [int, float] and vial_count > 2:
                vial_count = 2
            if not ((pd.isna(serum_volume) or serum_volume == 0) and (type(pbmc_conc) == str or pd.isna(pbmc_conc) or pbmc_conc == 0)):
                sample_volumes[sample_id] = [serum_volume, pbmc_conc, vial_count]
            else:
                missing_info['Sample ID'].append(sample_id)
                missing_info['Serum Volume'].append(serum_volume)
                missing_info['PBMC Conc'].append(pbmc_conc)
                missing_info['PBMC Vial Count'].append(vial_count)
        for _, sample in bsl2_samples.iterrows():
            sample_id = str(sample['Sample ID']).strip()
            try:
                serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("mlML ").split()[0])
            except:
                try:
                    serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("ulUL ").split()[0]) / 1000.
                except:
                    print(sample_id, serum_volume, type(serum_volume))
                    if type(sample['Total volume of serum (ml)']) == str:
                        serum_volume = 0
                    else:
                        serum_volume = sample['Total volume of serum (ml)']
            if type(serum_volume) != str and serum_volume > 4.5:
                serum_volume = 4.5
            if sample_id in no_pbmcs:
                pbmc_conc = 0
                vial_count = 0
            else:
                pbmc_conc = sample['# cell per aliquot']
                vial_count = sample['# of aliquots frozen']
                if (type(pbmc_conc) != float or pd.isna(pbmc_conc)) and type(pbmc_conc) != int:
                    pbmc_conc = 0
                    vial_count = 0
            if type(vial_count) in [int, float] and vial_count > 2:
                vial_count = 2
            if not pd.isna(serum_volume):
                sample_volumes[sample_id] = [serum_volume, pbmc_conc, vial_count]
            else:
                missing_info['Sample ID'].append(sample_id)
                missing_info['Serum Volume'].append(serum_volume)
                missing_info['PBMC Conc'].append(pbmc_conc)
                missing_info['PBMC Vial Count'].append(vial_count)
        with open("data/mit_volumes.pkl", "wb+") as f:
            pickle.dump(sample_volumes, f)
    else:
        with open('data/mit_volumes.pkl', 'rb') as f:
            sample_volumes = pickle.load(f)
    mit_days = [0, 30, 60, 90, 180, 300, 540, 720]
    pri_months = [0, 6, 12]
    header_1 = ['', '']
    for i in range(1,9):
        header_1.extend(['Visit {}'.format(i), '', ''])
    header_2 = ['Cohort', 'Seronet Participant ID'] + ['Volume of  Serum Collected (mL)', 'PBMCs concentration per ml (x10^6)', '# of 1 vials'] * 8
    data = {'Cohort': [], 'SERONET ID': [], 'Days from Index': [], 'Vaccine': [], '1st Dose Date': [], 'Days to 1st': [], '2nd Dose Date': [], 'Days to 2nd': [], 'Boost Vaccine': [], 'Boost Date': [], 'Days to Boost': [], 'Participant ID': [], 'Date': [], 'Post-Baseline': [], 'Sample ID': [], 'Visit Type': [], 'Qualitative': [], 'Quantitative': [], 'Spike endpoint': [], 'AUC': [], 'Log2AUC': [], 'Volume of Serum Collected (mL)': [], 'PBMC concentration per mL (x10^6)': [], '# of PBMC vials': []}
    short_window = 14
    long_window = 21
    participant_data = {}
    with open(script_output + 'SERONET.csv', 'w+', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header_1)
        writer.writerow(header_2)
        for participant, samples in participant_samples.items():
            try:
                samples.sort(key=lambda x: x[0])
            except:
                print(participant)
                print(samples)
                print()
                continue
            if len(samples) < 1:
                print(participant)
                print()
                continue
            baseline = samples[0][0]
            participant_data[participant] = {}
            times = [d for d in mit_days]
            if participant_study[participant] == 'TITAN':
                seronet_id = "14_T{}".format(samples[0][1])
                index_date = titan_data.loc[participant, '3rd Dose Vaccine Date'].date()
                participant_data[participant]['Vaccine'] = titan_data.loc[participant, 'Vaccine Type']
                try:
                    participant_data[participant]['1st Dose Date'] = titan_data.loc[participant, 'Vaccine #1 Date'].date()
                except:
                    participant_data[participant]['1st Dose Date'] = titan_data.loc[participant, 'Vaccine #1 Date']
                try:
                    participant_data[participant]['2nd Dose Date'] = titan_data.loc[participant, 'Vaccine #2 Date'].date()
                except:
                    participant_data[participant]['2nd Dose Date'] = titan_data.loc[participant, 'Vaccine #2 Date']
                try:
                    participant_data[participant]['Boost Date'] = titan_data.loc[participant, '3rd Dose Vaccine Date'].date()
                except:
                    participant_data[participant]['Boost Date'] = titan_data.loc[participant, '3rd Dose Vaccine Date']
                participant_data[participant]['Boost Vaccine'] = titan_data.loc[participant, '3rd Dose Vaccine Type']
            elif participant_study[participant] == 'IRIS':
                seronet_id = "14_I{}".format(samples[0][1])
                try:
                    index_date = iris_data.loc[participant, 'Second Dose Date'].date()
                except:
                    print("Participant {} is missing vaccine dose 2 date (dose 1 on {})".format(participant, iris_data.loc[participant, 'First Dose Date']))
                    index_date = 'N/A'
                participant_data[participant]['Vaccine'] = iris_data.loc[participant, 'Which Vaccine?']
                try:
                    participant_data[participant]['1st Dose Date'] = iris_data.loc[participant, 'First Dose Date'].date()
                except:
                    participant_data[participant]['1st Dose Date'] = iris_data.loc[participant, 'First Dose Date']
                try:
                    participant_data[participant]['2nd Dose Date'] = iris_data.loc[participant, 'Second Dose Date'].date()
                except:
                    participant_data[participant]['2nd Dose Date'] = iris_data.loc[participant, 'Second Dose Date']
                try:
                    participant_data[participant]['Boost Date'] = iris_data.loc[participant, 'Third Dose Date'].date()
                except:
                    participant_data[participant]['Boost Date'] = iris_data.loc[participant, 'Third Dose Date']
                participant_data[participant]['Boost Vaccine'] = iris_data.loc[participant, 'Third Dose Type']
            elif participant_study[participant] == 'MARS':
                seronet_id = "14_M{}".format(samples[0][1])
                try:
                    index_date = mars_data.loc[participant, 'Vaccine #2 Date'].date()
                except:
                    print("Participant {} is missing vaccine dose 2 date (dose 1 on {})".format(participant, mars_data.loc[participant, 'Vaccine #1 Date']))
                    index_date = 'N/A'
                participant_data[participant]['Vaccine'] = mars_data.loc[participant, 'Vaccine Name']
                try:
                    participant_data[participant]['1st Dose Date'] = mars_data.loc[participant, 'Vaccine #1 Date'].date()
                except:
                    participant_data[participant]['1st Dose Date'] = mars_data.loc[participant, 'Vaccine #1 Date']
                try:
                    participant_data[participant]['2nd Dose Date'] = mars_data.loc[participant, 'Vaccine #2 Date'].date()
                except:
                    participant_data[participant]['2nd Dose Date'] = mars_data.loc[participant, 'Vaccine #2 Date']
                try:
                    participant_data[participant]['Boost Date'] = mars_data.loc[participant, 'Additional Vaccinations'].date()
                except:
                    participant_data[participant]['Boost Date'] = mars_data.loc[participant, 'Additional Vaccinations']
                participant_data[participant]['Boost Vaccine'] = mars_data.loc[participant, 'Additional Vaccine Type ']
            elif participant_study[participant] == 'PRIORITY':
                seronet_id = "14_P{}".format(samples[0][1])
                index_date = baseline
                times = [m for m in pri_months]
                participant_data[participant]['Vaccine'] = ''
                participant_data[participant]['1st Dose Date'] = ''
                participant_data[participant]['2nd Dose Date'] = ''
                participant_data[participant]['Boost Date'] = ''
                participant_data[participant]['Boost Vaccine'] = ''
            dts = [s[0] for s in samples]
            if type(index_date) == datetime.date:
                start_idx = bisect_left(dts, index_date)
            else:
                start_idx = 0
            if start_idx > 0:
                start_idx -= 1
            samples = samples[start_idx:] # This is our messy "remove all but one pre-index sample"
            samples_to_write = [participant_study[participant], seronet_id]
            one_in_window = False
            for date_, sample_id, visit_type, result, result_new in samples:
                include = False
                while True: # Very slapdash
                    if len(times) == 0:
                        break
                    if participant_study[participant] == 'PRIORITY':
                        timediff = relativedelta(months=times[0])
                        window = long_window
                    else:
                        timediff = relativedelta(days=times[0])
                        if times[0] < 100:
                            window = short_window
                        else:
                            window = long_window
                    if type(index_date) == datetime.date and times[0] == 0: # We may later have to add TIM baselines here, try to distinguish
                        if date_ <= index_date:
                            include = True
                            one_in_window = True
                            times.pop(0)
                            break
                        else:
                            samples_to_write.extend(['No Visit', 'No Visit', 'No Visit'])
                            times.pop(0) # Hopefully we're TIM and we missed the baseline
                    elif type(index_date) == datetime.date:
                        if abs(int((index_date + timediff - date_).days)) <= window:
                            one_in_window = True
                            include = True
                            times.pop(0)
                            break
                        elif date_ >= index_date + timediff + relativedelta(days=window):
                            samples_to_write.extend(['No Visit', 'No Visit', 'No Visit'])
                            times.pop(0)
                        else:
                            break
                    else:
                        print(index_date)
                        print()
                        for time in times:
                            samples_to_write.extend(['-', '-', '-'])
                        times = []
                        break
                if not include:
                    continue
                if type(index_date) != datetime.date:
                    days_from_ind = '??? Double check!'
                else:
                    days_from_ind = int((date_ - index_date).days)
                sample_id = str(sample_id).strip()
                data['Cohort'].append(participant_study[participant])
                data['SERONET ID'].append(seronet_id)
                data['Days from Index'].append(days_from_ind)
                data['Vaccine'].append(participant_data[participant]['Vaccine'])
                data['1st Dose Date'].append(participant_data[participant]['1st Dose Date'])
                try:
                    data['Days to 1st'].append(int((date_ - data['1st Dose Date'][-1]).days))
                except:
                    data['Days to 1st'].append('')
                data['2nd Dose Date'].append(participant_data[participant]['2nd Dose Date'])
                try:
                    data['Days to 2nd'].append(int((date_ - data['2nd Dose Date'][-1]).days))
                except:
                    data['Days to 2nd'].append('')
                data['Boost Date'].append(participant_data[participant]['Boost Date'])
                try:
                    data['Days to Boost'].append(int((date_ - data['Boost Date'][-1]).days))
                except:
                    data['Days to Boost'].append('')
                data['Boost Vaccine'].append(participant_data[participant]['Boost Vaccine'])
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
                if res[1] == '-':
                    data['Log2AUC'].append('-')
                elif float(res[1]) == 0.:
                    data['Log2AUC'].append(0.)
                else:
                    data['Log2AUC'].append(np.log2(res[1]))
                data['Spike endpoint'].append(res[0])
                data['AUC'].append(res[1])
                if sample_id in sample_volumes.keys():
                    vols = sample_volumes[sample_id]
                else:
                    vols = ['?', '?', '?']
                if participant_study[participant] == 'PRIORITY':
                    vols[1] = "N/A"
                    vols[2] = "N/A"
                data['Volume of Serum Collected (mL)'].append(vols[0])
                data['PBMC concentration per mL (x10^6)'].append(vols[1])
                data['# of PBMC vials'].append(vols[2])
                samples_to_write.extend(vols)
            while len(times) > 0:
                if participant_study[participant] == 'PRIORITY':
                    timediff = relativedelta(months=times[0])
                else:
                    timediff = relativedelta(days=times[0])
                if datetime.date.today() <= (index_date + timediff) + relativedelta(days=window):
                    samples_to_write.extend(['-', '-', '-'])
                else:
                    samples_to_write.extend(['No Visit', 'No Visit', 'No Visit'])
                times.pop(0)
            if one_in_window:
                writer.writerow(samples_to_write)
    report = pd.DataFrame(data)
    report.to_excel(script_output + 'SERONET_In_Window_Data.xlsx'.format(date.today().strftime("%m.%d.%y")), index=False)

