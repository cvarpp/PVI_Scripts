from email.policy import default
import numpy as np
import pandas as pd
import util
from datetime import date
import datetime
from dateutil import parser

if __name__ == '__main__':
    paris_info = pd.read_excel(util.paris_tracker , sheet_name='Main', header=8)
    umbrella_info = pd.read_excel(util.umbrella_tracker, sheet_name='Summary', header=0)
    samples = pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=6, keep_default_na=False)
    visit_type = "Visit Type / Samples Needed"
    results_recieved = 'Results'
    result_stat_col = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    result_value_col = 'Ab Concentration (Units - AU/mL)'
    was_shared_col = 'Clinical Ab Result Shared?'
    result_type = "Visit Type / Samples Needed"
    samplesClean = samples.dropna(subset=['Participant ID'])
    participants = [str(x).strip().upper() for x in pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=6)['Participant ID']]
    def not_shared(val):
        return pd.isna(val) or val == 'No' or val == ''    
    
    share_filter = samples[was_shared_col].apply(not_shared)
    ppl_filter = samples['Participant ID'] != ''
    result_filter = (samples[result_value_col] != '') & (samples[result_value_col] != 'N/A')
    #merged_emails = pd.merge(paris_info, umbrella_info, on='Participant ID', how='inner')
    #matching_email = merged_emails.loc[merged_emails['Participant ID'] == 'Participant ID']

    drop_cols = ['Participant ID', result_value_col]    

    samplesClean = samples[share_filter & ppl_filter & result_filter].dropna(subset=drop_cols)
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
    data = {'Participant ID': [], 'Date': [], 'Sample ID': [], 'Visit Type': [], 'Qualitative': [], 'Quantitative': []}
    for participant, results in participant_results.items():
       # try:
        #    samples.sort(key=lambda x: x[0])
         #   results.sort(key=lambda x: x[0])
        #except:
         #   print(participant)
          #  print(samples)
           # print(results)
            #continue
        if len(samples) < 1:
            if len(results) < 1:
                continue
        #baseline = samples[0][0]
        for date_, sample_id, visit_type, result, result_new in results:
            sample_id = str(sample_id).strip()
            data['Participant ID'].append(participant)
            data['Date'].append(date_)
            data['Qualitative'].append(result)
            data['Quantitative'].append(result_new)
if __name__ == '__main__':
    report = pd.DataFrame(data)
    output_filename = util.script_output + 'result_reporting_test_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    report.to_excel(output_filename, index=False)
    cutoff_date = datetime.date.today() - datetime.timedelta(days=60)
    old_results = report[report['Date'] <= cutoff_date]
    new_results = report[report['Date'] > cutoff_date]
    assert old_results.shape[0] + new_results.shape[0] == report.shape[0]
    output_filename = util.script_output + 'result_reporting_test_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    writer = pd.ExcelWriter(output_filename)
    new_results.to_excel(writer, sheet_name='To Share')
    old_results.to_excel(writer, sheet_name='Misunderstood')
    writer.save()
    writer.close()
    print("Report written to {}".format(output_filename))