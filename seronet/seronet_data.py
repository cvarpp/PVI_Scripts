import pandas as pd
import csv
import datetime
from dateutil.relativedelta import relativedelta
import util
import argparse
from seronet.d4_all_data import pull_from_source

def filter_windows(unfiltered):
    unfiltered['Date'] = pd.to_datetime(unfiltered['Date'])
    no_index_date = []
    mit_days = [0, 30, 60, 90, 180, 300, 540, 720]
    pri_months = [0, 6, 12]
    header_1 = ['', '']
    for i in range(1,9):
        header_1.extend(['Visit {}'.format(i), '', ''])
    header_2 = ['Cohort', 'Seronet Participant ID'] + ['Volume of  Serum Collected (mL)', 'PBMCs concentration per ml (x10^6)', '# of PBMC vials'] * 8
    
    columns = [col for col in unfiltered.columns]
    columns.insert(2, 'Days from Index')
    data = {col: [] for col in columns}
    short_window = 14
    long_window = 21
    with open(util.report_form, 'w+', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header_1)
        writer.writerow(header_2)
        for participant in unfiltered['Participant ID'].unique():
            try:
                samples = unfiltered[unfiltered['Participant ID'] == participant].sort_values(by='Date')
            except:
                print(participant, "has samples that don't sort (see below)")
                print(samples)
                print()
                continue
            if samples.shape[0] < 1:
                print(participant, "has no sample")
                print()
                continue
            baseline = samples['Date'].to_numpy()[0]
            # baseline_id = samples['Sample ID'].to_numpy()[0]
            times = [d for d in mit_days]
            cohort = samples['Cohort'].to_numpy()[0]
            d1 = samples['1st Dose Date'].to_numpy()[0]
            d2 = samples['2nd Dose Date'].to_numpy()[0]
            vax_type = samples['Vaccine'].to_numpy()[0]
            seronet_id = samples['Seronet ID'].to_numpy()[0]
            try:
                samples = samples[samples['Date'].apply(lambda val: pd.isna(d1) or val <= d1 or pd.isna(d2) or val > d2)]
            except:
                if cohort != 'PRIORITY':
                    print(participant, d1, d2, "failed filter")
                    continue
            if cohort == 'TITAN':
                d3 = samples['3rd Dose Date'].to_numpy()[0]
                index_date = d3
            elif cohort in ['MARS', 'IRIS']:
                if str(vax_type)[:1] == 'J':
                    index_date = d1
                else:
                    index_date = d2
            elif cohort == 'PRIORITY':
                index_date = baseline
                times = [m for m in pri_months]
            else:
                print(participant, "has invalid cohort", cohort)
                continue
            if pd.isna(index_date):
                no_index_date.append(participant)
                continue
            index_date = pd.to_datetime(index_date).date()
            samples['Clipped Post-Index'] = samples['Date'].apply(lambda val: int((val.date() - index_date).days))
            samples = samples.sort_values(by='Date').drop_duplicates(subset=['Clipped Post-Index'], keep='last').drop('Clipped Post-Index', axis=1)
            samples_to_write = [cohort, seronet_id]
            one_in_window = False
            for _, row in samples.iterrows():
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
                        if row['Date'] <= index_date:
                            include = True
                            one_in_window = True
                            times.pop(0)
                            break
                        else:
                            samples_to_write.extend(['No Visit', 'No Visit', 'No Visit'])
                            times.pop(0) # Hopefully we're TIM and we missed the baseline
                    elif type(index_date) == datetime.date:
                        if abs(int((index_date + timediff - row['Date'].date()).days)) <= window:
                            one_in_window = True
                            include = True
                            times.pop(0)
                            break
                        elif row['Date'] >= index_date + timediff + relativedelta(days=window):
                            samples_to_write.extend(['No Visit', 'No Visit', 'No Visit'])
                            times.pop(0)
                        else:
                            break
                    else:
                        print(index_date, "for participant", participant, "is invalid")
                        print()
                        for _ in times:
                            samples_to_write.extend(['-', '-', '-'])
                        times = []
                        break
                if not include:
                    continue
                data['Cohort'].append(cohort)
                data['Seronet ID'].append(seronet_id)
                days_from_ind = int((row['Date'].date() - index_date).days)
                data['Days from Index'].append(days_from_ind)
                for col in row.index.to_numpy()[2:]:
                    data[col].append(row[col])
                samples_to_write.extend([data['Volume of Serum Collected (mL)'], data['PBMC concentration per mL (x10^6)'], data['# of PBMC vials']])
            while len(times) > 0:
                if cohort == 'PRIORITY':
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
    report.to_excel(util.filtered, index=False)
    print("Participants with no index date:")
    print(no_index_date)
    print()
    return report



if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Make Seronet monthly sample report.')
    argParser.add_argument('-u', '--update', action='store_true')
    args = argParser.parse_args()
    if args.update:
        unfiltered = pull_from_source()
    else:
        unfiltered = pd.read_excel(util.unfiltered)
    filter_windows(unfiltered)