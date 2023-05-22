import pandas as pd
import datetime
import util

def titan_reader(sname, header=1):
    return (pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name=sname, header=header)
              .dropna(subset=['Umbrella Participant ID'])
              .assign(participant_id=lambda df: df['Umbrella Participant ID'].str.strip())
              .set_index('participant_id'))

if __name__ == '__main__':
    titan_doses = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_doses['Participant ID'] = titan_doses['Umbrella Corresponding Participant ID'].apply(lambda val: val.strip().upper())
    titan_doses.set_index('Participant ID', inplace=True)
    titan_data = titan_reader('Demographics & First Two Doses', header=3)
    participants = titan_data.index.to_numpy()
    titan_clin = titan_reader('Third Dose')
    titan_fourth = titan_reader('Booster Dose #1')
    titan_fifth = titan_reader('Booster Dose #2')
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
    for participant, row in titan_fourth.iterrows():
        fourth_meds = row['Maintenance immunosuppresion at time of first booster dose ']
        if type(fourth_meds) != str:
            continue
        for med in fourth_meds.split('+'):
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
                med_dftobe['Report_Time'].append(titan_doses.loc[participant, 'First Booster Dose Date'].date() - datetime.timedelta(days=14))
            except Exception as _e:
                print(participant, "has invalid first booster", titan_doses.loc[participant, 'First Booster Dose Date'], type(titan_doses.loc[participant, 'First Booster Dose Date']))
                med_dftobe['Report_Time'].append(pd.to_datetime('1/1/2021').date())
            med_dftobe['Comments'].append('')
    for participant, row in titan_fifth.iterrows():
        fifth_meds = row['Maintenance immunosuppresion at time of second booster dose ']
        if type(fifth_meds) != str:
            continue
        for med in fifth_meds.split('+'):
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
                med_dftobe['Report_Time'].append(titan_doses.loc[participant, 'Second Booster Dose Date'].date() - datetime.timedelta(days=14))
            except Exception as _e:
                print(participant, "has invalid second booster", titan_doses.loc[participant, 'Second Booster Dose Date'], type(titan_doses.loc[participant, 'Second Booster Dose Date']))
                med_dftobe['Report_Time'].append(pd.to_datetime('1/1/2021').date())
            med_dftobe['Comments'].append('')
    med_df = pd.DataFrame(med_dftobe)
    med_df.to_excel(util.titan_folder + 'titan_meds.xlsx', index=False)