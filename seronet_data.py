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
import util
from util import script_folder, script_output
import argparse
from d4_all_data import pull_from_source

def filter_windows(unfiltered):
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
                samples = unfiltered[unfiltered['Participant ID'] == participant].sort(by='Date')
            except:
                print(participant, "doesn't sort")
                print()
                continue
            if samples.shape[0] < 1:
                print(participant, "has only one sample")
                print()
                continue
            baseline = samples['Date'].to_numpy()[0]
            baseline_id = samples['Sample ID'].to_numpy()[0]
            participant_data[participant] = {}
            times = [d for d in mit_days]
            cohort = samples['Cohort'].to_numpy()[0]
            d1 = samples['1st Dose Date'].to_numpy()[0]
            d2 = samples['2nd Dose Date'].to_numpy()[0]
            d3 = samples['Boost Date'].to_numpy()[0]
            vax_type = samples['Vaccine'].to_numpy()[0]
            seronet_id = "14_{}{}".format(cohort[0], baseline_id)
            samples = samples[samples['Date'].apply(lambda val: val <= d1 or pd.isna(d2) or val > d2)]
            if cohort == 'TITAN':
                index_date = d3
            elif cohort in ['MARS', 'IRIS']:
                if vax_type[:1] == 'J':
                    index_date = d1
                else:
                    index_date = d2
            elif cohort == 'PRIORITY':
                index_date = baseline
                times = [m for m in pri_months]
            else:
                print(participant, "has invalid cohort", cohort)
                continue
            samples['Clipped Post-Index'] = samples['Date'].apply(lambda val: int((val - index_date).days))
            samples = samples.drop_duplicates(subset=['Clipped Post-Index'])
            samples_to_write = [cohort, seronet_id]
            one_in_window = False
            for date_, sample_id, visit_type, result, result_new in samples:
                include = False
                while True: # Very slapdash
                    if len(times) == 0:
                        break
                    if cohort == 'PRIORITY':
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



if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Make Seronet monthly sample report.')
    parser.add_argument('-u', '--update', action='store_true')
    args = parser.parse_args()
    if args.update:
        unfiltered = pull_from_source()
    else:
        unfiltered = pd.read_excel(util.unfiltered)
    filter_windows(unfiltered)