import pandas as pd
import datetime
import util

if __name__ == '__main__':
    titan_data = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Demographics & First Two Doses', header=3).dropna(subset=['Umbrella Participant ID'])
    titan_data['Participant ID'] = titan_data['Umbrella Participant ID'].apply(lambda val: val.strip().upper())
    participants = titan_data['Participant ID'].unique()
    titan_data.set_index('Participant ID', inplace=True)
    titan_clin = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Third Dose', header=1).dropna(subset=['Umbrella Participant ID'])
    titan_clin['Participant ID'] = titan_clin['Umbrella Participant ID'].apply(lambda val: val.strip().upper())
    titan_clin.set_index('Participant ID', inplace=True)
    med_cols = ['Participant ID', 'MRN', 'Health_Condition_Or_Disease', 'Treatment', 'Dosage', 'Dosage_Units', 'Dosage_Regimen', 'Start_Date', 'Stop_Date', 'Report_Time', 'Comments']
    med_dftobe = {col: [] for col in med_cols}
    induct_col = "Anti thymocyte Globulin (Thymoglobulin/ATG) or Basiliximab (Simulect)"
    valid_inducts = ["Thymoglobulin/ATG", "Simulect"]
    for participant, row in titan_data.iterrows():
        transplant_date = row['Date of transplant (only enter this) ']
        induction = row[induct_col]
        if induction in valid_inducts:
            start = transplant_date
            end = transplant_date + datetime.timedelta(days=4)
            med_dftobe['Participant ID'].append(participant)
            med_dftobe['MRN'].append(row['MRN '])
            med_dftobe['Health_Condition_Or_Disease'].append('Transplant (induction immunosuppression)')
            med_dftobe['Treatment'].append(induction)
            med_dftobe['Dosage'].append('Unknown')
            med_dftobe['Dosage_Units'].append('Unknown')
            med_dftobe['Dosage_Regimen'].append('Unknown')
            med_dftobe['Start_Date'].append(start)
            med_dftobe['Stop_Date'].append(end)
            med_dftobe['Report_Time'].append(start)
            med_dftobe['Comments'].append('')
    for participant, row in titan_clin.iterrows():
        third_meds = row['Maintenance immunosuppresion at time of third dose ']
        if type(third_meds) != str:
            continue
        for med in third_meds.split('+'):
            med_dftobe['Participant ID'].append(participant)
            med_dftobe['MRN'].append(row['MRN '])
            med_dftobe['Health_Condition_Or_Disease'].append('Transplant (maintenance immunosuppression)')
            med_dftobe['Treatment'].append(med.strip())
            med_dftobe['Dosage'].append('Unknown')
            med_dftobe['Dosage_Units'].append('Unknown')
            med_dftobe['Dosage_Regimen'].append('Unknown')
            med_dftobe['Start_Date'].append('Not Reported')
            med_dftobe['Stop_Date'].append('Ongoing')
            try:
                med_dftobe['Report_Time'].append(row['Date of 3rd Dose '].date() - datetime.timedelta(days=14))
            except Exception as _e:
                print(participant, "has invalid 3rd dose", row['Date of 3rd Dose '], type(row['Date of 3rd Dose ']))
                med_dftobe['Report_Time'].append(pd.to_datetime('1/1/2021').date())
            med_dftobe['Comments'].append('')
    med_df = pd.DataFrame(med_dftobe)
    med_df.to_excel(util.titan_folder + 'titan_meds.xlsx', index=False)