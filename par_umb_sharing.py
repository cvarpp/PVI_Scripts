from email.policy import default
import numpy as np
import pandas as pd
import util
from datetime import date
import datetime
from dateutil import parser
from helpers import query_intake


if __name__ == '__main__':
    paris_info = pd.read_excel(util.paris_tracker, sheet_name='Main', header=8)
    umbrella_info = pd.read_excel(util.umbrella_tracker, sheet_name='Summary', header=0)
    was_shared_col = 'Clinical Ab Result Shared?'
    def not_shared(val):
        return pd.isna(val) or val == 'No' or val == ''

    today = date.today()
    samples = query_intake()
    share_filter = samples[was_shared_col].apply(not_shared)
    result_filter = (samples['Qualitative'] != '') & ~samples['Qualitative'].apply(pd.isna)
    samplesClean = samples[share_filter & result_filter].copy()
    samplesCleanAll = samplesClean.copy()
    #samplesClean = samplesClean[samplesClean['Date Collected'] >= (today - datetime.timedelta(days=60))]
    
    # older_results = samplesClean[samplesClean['Date Collected'] < (today - datetime.timedelta(days=60))]
    older_results = []
    #df for results older than 60 days
    pemail = paris_info[['Subject ID', 'E-mail']].rename(columns={'E-mail': 'Email'})
    uemail = umbrella_info[['Subject ID', 'Email']]
    # emails = pd.concat([pemail, uemail])
    emails = (pd.concat([pemail, uemail])
                .assign(pid=lambda df: df['Subject ID'].str.strip())
                .drop_duplicates(subset='pid').set_index('pid'))

    participants = samplesClean['participant_id'].unique()
    participant_results = {participant: [] for participant in participants}
    for sid, sample in samplesClean.iterrows():
        participant = sample['participant_id']
        if pd.isna(sample[was_shared_col]) or sample[was_shared_col] == "No":
            if str(sample['Qualitative']).strip().upper()[:2] == "NE":
                sample['Qualitative'] = "Negative"
                sample['Quantitative'] = "Negative"
            try:
                participant_results[participant].append((pd.to_datetime(sample['Date Collected']).date(), sid, sample[util.visit_type], sample['Qualitative'], sample['Quantitative']))
            except:
                print("Sample", sid, "improperly recorded as collected on", sample['Date Collected'])
                print("Not included in result sharing")
    data_filtered = {'Participant ID': [], 'Sample ID': [], 'Date': [], 'Email': [], 'Visit Type': [], 'Qualitative': [], 'Quantitative': []}
    last_60_days = today - datetime.timedelta(days=60)
    for participant, results in participant_results.items():
        for result in results:
            date_collected, sample_id, util.visit_type, result_stat, result_value = result
            if date_collected < last_60_days:
                older_results.append({'Participant ID': participant, 'Sample ID': sample_id, 'Date': date_collected, 'Visit Type': util.visit_type,'Qualitative': result_stat,'Quantitative': result_value})
                #add older results to 'older_results' df
                continue
            if date_collected >= last_60_days:
                data_filtered['Participant ID'].append(participant)
                data_filtered['Sample ID'].append(sample_id)
                data_filtered['Date'].append(date_collected)
                data_filtered['Visit Type'].append(util.visit_type)
                data_filtered['Qualitative'].append(result_stat)
                data_filtered['Quantitative'].append(result_value)
                if participant in emails.index:
                    email = emails.loc[participant, 'Email']
                else:
                    email = ''
                data_filtered['Email'].append(email)
if __name__ == '__main__':
    report = pd.DataFrame(data_filtered)
    old_report = pd.DataFrame(older_results)
    report.columns = ['Participant ID', 'Sample ID', 'Date', 'Email', 'Visit Type / Sample ID', 'Qualitative', 'Quantitative']
    output_filename = util.sharing + 'result_reporting_test_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    with pd.ExcelWriter(output_filename) as writer:
        report.to_excel(writer, sheet_name='Needs Results - Recent', index=False)
        old_report.to_excel(writer, sheet_name='Needs Results - Old(60d+)', index=False)
        #make new sheet
    print("Report written to {}".format(output_filename)) 
