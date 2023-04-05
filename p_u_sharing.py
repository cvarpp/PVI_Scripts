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
    results_recieved = 'Results received' # never used
    was_shared_col = 'Clinical Ab Result Shared?'
    def not_shared(val):
        return pd.isna(val) or val == 'No' or val == ''

    samples = query_intake()
    share_filter = samples[was_shared_col].apply(not_shared)
    result_filter = (samples['Qualitative'] != '') & ~samples['Qualitative'].apply(pd.isna)
    samplesClean = samples[share_filter & result_filter].copy()

    pemail = paris_info[['Subject ID', 'E-mail']].rename(columns={'E-mail': 'Email'})
    uemail = umbrella_info[['Subject ID', 'Email']]
    emails = pd.concat([pemail, uemail])

    participants = samplesClean['participant_id'].unique()
    participant_results = {participant: [] for participant in participants}
    for sid, sample in samplesClean.iterrows():
        participant = sample['participant_id']
        if participant in participant_results.keys():
            try:
                participant_results[participant].append((parser.parse(str(sample['Date Collected'])).date(), sid, sample[util.visit_type], sample['Qualitative'], sample['Quantitative']))
            except:
                print(sample['Date Collected'])
                participant_results[participant].append((parser.parse('1/1/1900').date(), sid, sample[util.visit_type], sample['Qualitative'], sample['Quantitative']))
        if pd.isna(sample[was_shared_col]) or sample[was_shared_col] == "No":
            if participant in participant_results.keys():
                    if str(sample['Qualitative']).strip().upper() == "NEGATIVE":
                        sample['Quantitative'] = "Negative"
                    try:
                        participant_results[participant].append((parser.parse(str(sample['Date Collected'])).date(), sample['Sample ID'], sample[util.visit_type], sample['Qualitative'], sample['Quantitative']))
                    except:
                        print("Sample", sid, "improperly recorded as collected on", sample['Date Collected'])
                        print("Not included in result sharing")
    data_filtered = {'Participant ID': [], 'Sample ID': [], 'Date': [], 'Email': [], 'Visit Type': [], 'Qualitative': [], 'Quantitative': []}
    today = date.today()
    last_60_days = today - datetime.timedelta(days=60)
    for participant, results in participant_results.items():
        for result in results:
            if result[0] >= last_60_days:
                for date_collected, sample_id, util.visit_type, result_stat, result_value in participant_results[participant]:
                    data_filtered['Participant ID'].append(participant)
                    data_filtered['Sample ID'].append(sample_id)
                    data_filtered['Date'].append(date_collected)
                    data_filtered['Visit Type'].append(util.visit_type)
                    data_filtered['Qualitative'].append(result_stat)
                    data_filtered['Quantitative'].append(result_value)
                    if not emails.loc[emails['Subject ID'] == participant].empty:
                        email = emails.loc[emails['Subject ID'] == participant, 'Email'].iloc[0]
                        data_filtered['Email'].append(email)
                    else:
                        data_filtered['Email'].append('')
if __name__ == '__main__':
    print(data_filtered)
    report = pd.DataFrame(data_filtered)
    report.columns = ['Participant ID', 'Sample ID', 'Date', 'Email', 'Visit Type / Sample ID', 'Qualitative', 'Quantitative']
    output_filename = util.sharing + 'result_reporting_test_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    writer = pd.ExcelWriter(output_filename)
    report.to_excel(writer, index=False)
    writer.save()
    writer.close()
    print("Report written to {}".format(output_filename))
