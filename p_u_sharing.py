from email.policy import default
import numpy as np
import pandas as pd
import util
from datetime import date
import datetime
from dateutil import parser

if __name__ == '__main__':
    paris_info = pd.read_excel(util.paris_tracker, sheet_name='Main', header=8)
    umbrella_info = pd.read_excel(util.umbrella_tracker, sheet_name='Summary', header=0)
    samples = pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=6, keep_default_na=False)
    visit_type = "Visit Type / Samples Needed"
    results_recieved = 'Results'
    result_stat_col = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    result_value_col = 'Ab Concentration (Units - AU/mL)'
    was_shared_col = 'Clinical Ab Result Shared?'
    result_type = "Visit Type / Samples Needed"
    pemail = paris_info[['Subject ID', 'E-mail']].rename(columns={'E-mail': 'Email'})
    uemail = umbrella_info[['Subject ID', 'Email']]
    samplesClean = samples.dropna(subset=['Participant ID'])
    participants = [str(x).strip().upper() for x in pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=6)['Participant ID']]
    def not_shared(val):
        return pd.isna(val) or val == 'No' or val == ''

    share_filter = samples[was_shared_col].apply(not_shared)
    ppl_filter = samples['Participant ID'] != ''
    result_filter = (samples[result_value_col] != '') & (samples[result_value_col] != 'N/A')
    emails = pd.concat([pemail, uemail])

    samplesClean = samples[share_filter & ppl_filter & result_filter].dropna(subset=['Participant ID', result_value_col])
    participants = [str(x).strip().upper() for x in samplesClean['Participant ID'].unique()]
    participant_results = {participant: [] for participant in participants}
    for _, sample in samplesClean.iterrows():
        participant = str(sample['Participant ID']).strip().upper()
        if participant in participant_results.keys():
            if str(sample[result_stat_col]).strip().upper() == "NEGATIVE":
                sample[result_value_col] = "Negative"
            if pd.isna(sample[result_value_col]):
                result_new = '-'
            elif type(sample[result_value_col]) == int:
                result_new = sample[result_value_col]
            elif str(sample[result_value_col]).strip().upper() == "NEGATIVE":
                result_new = 1.
            else:
                result_new = sample[result_value_col]
            try:
                participant_results[participant].append((parser.parse(str(sample['Date Collected'])).date(), sample['Sample ID'], sample[visit_type], sample[result_stat_col], result_new))
            except:
                print(sample['Date Collected'])
                participant_results[participant].append((parser.parse('1/1/1900').date(), sample['Sample ID'], sample[visit_type], sample[result_stat_col], result_new))
        if pd.isna(sample[was_shared_col]) or sample[was_shared_col] == "No":
            if participant in participant_results.keys():
                    if str(sample[result_stat_col]).strip().upper() == "NEGATIVE":
                        sample[result_value_col] = "Negative"
                    try:
                        participant_results[participant].append((parser.parse(str(sample['Date Collected'])).date(), sample['Sample ID'], sample[visit_type], sample[result_stat_col], sample[result_value_col]))
                    except:
                        print("Sample", sample['Sample_ID'], "improperly recorded as collected on", sample['Date Collected'])
                        print("Not included in result sharing")
    data_filtered = {'Participant ID': [], 'Sample ID': [], 'Date': [], 'Email': [], 'Visit Type': [], 'Qualitative': [], 'Quantitative': []}
    today = date.today()
    last_60_days = today - datetime.timedelta(days=60)
    for participant, results in participant_results.items():
        for result in results:
            if result[0] >= last_60_days:
                for date_collected, sample_id, visit_type, result_stat, result_value in participant_results[participant]:
                    data_filtered['Participant ID'].append(participant)
                    data_filtered['Sample ID'].append(sample_id)
                    data_filtered['Date'].append(date_collected)
                    data_filtered['Visit Type'].append(visit_type)
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
