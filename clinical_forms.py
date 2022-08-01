import pandas as pd
import datetime
from dateutil import parser


# Priority does not use CRF or CPT

def specimenize(row):
    if row['Serum']:
        if row['PBMC']:
            return "Serum|PBMC|Plasma"
        else:
            return "Serum"
    elif row['PBMC']:
        return "PBMC|Plasma"
    else:
        print(row['Sample ID'], 'has neither serum nor plasma and probably should be dropped!')
        return "OOPS"

if __name__ == '__main__':
    base_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Date_Duration_From_Index', 'Lost_to_Follow_Up', 'Final_Visit', 'Age', 'Sex_At_Birth', 'Race', 'Ethnicity', 'Height', 'Weight', 'BMI', 'Location', 'Biospecimens_Collected', 'Diabetes', 'Diabetes_Description_Or_ICD10_codes', 'Hypertension', 'Hypertension_Description_Or_ICD10_codes', 'Obesity', 'Obesity_Description_Or_ICD10_codes', 'Cardiovascular_Disease', 'Cardiovascular_Disease_Description_Or_ICD10_codes', 'Chronic_Lung_Disease', 'Chronic_Lung_Disease_Description_Or_ICD10_codes', 'Chronic_Kidney_Disease', 'Chronic_Kidney_Disease_Description_Or_ICD10_codes', 'Chronic_Liver_Disease', 'Chronic_Liver_Disease_Description_Or_ICD10_codes', 'Acute_Liver_Disease', 'Acute_Liver_Disease_Description_Or_ICD10_codes', 'Immunosuppressive_Condition', 'Immunosuppressive_Condition_Description_Or_ICD10_codes', 'Autoimmune_Disorder', 'Autoimmune_Disorder_Description_Or_ICD10_codes', 'Chronic_Neurological_Condition', 'Chronic_Neurological_Condition_Description_Or_ICD10_codes', 'Chronic_Oxygen_Requirement', 'Chronic_Oxygen_Requirement_Description_Or_ICD10_codes', 'Inflammatory_Disease', 'Inflammatory_Disease_Description_Or_ICD10_codes', 'Viral_Infection', 'Viral_Infection_ICD10_codes_Or_Agents', 'Bacterial_Infection', 'Bacterial_Infection_ICD10_codes_Or_Agents', 'Cancer', 'Cancer_Description_Or_ICD10_codes', 'Substance_Abuse_Disorder', 'Substance_Abuse_Disorder_Description_Or_ICD10_codes', 'Organ_Transplant_Recipient', 'Organ_Transplant_Description_Or_ICD10_codes', 'Other_Health_Condition_Description_Or_ICD10_codes', 'ECOG_Status', 'Smoking_Or_Vaping_Status', 'Alcohol_Use', 'Drug_Type', 'Drug_Use', 'Vaccination_Record', 'Comments']
    follow_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Visit_Date_Duration_From_Index', 'Lost_to_Follow_Up', 'Final_Visit', 'Baseline_Visit', 'Number_of_Missed_Scheduled_Visits', 'Unscheduled_Visit', 'Biospecimens_Collected', 'Diabetes', 'Diabetes_Description_Or_ICD10_codes', 'Hypertension', 'Hypertension_Description_Or_ICD10_codes', 'Obesity', 'Obesity_Description_Or_ICD10_codes', 'Cardiovascular_Disease', 'Cardiovascular_Disease_Description_Or_ICD10_codes', 'Chronic_Lung_Disease', 'Chronic_Lung_Disease_Description_Or_ICD10_codes', 'Chronic_Kidney_Disease', 'Chronic_Kidney_Disease_Description_Or_ICD10_codes', 'Chronic_Liver_Disease', 'Chronic_Liver_Disease_Description_Or_ICD10_codes', 'Acute_Liver_Disease', 'Acute_Liver_Disease_Description_Or_ICD10_codes', 'Immunosuppressive_Condition', 'Immunosuppressive_Condition_Description_Or_ICD10_codes', 'Autoimmune_Disorder', 'Autoimmune_Disorder_Description_Or_ICD10_codes', 'Chronic_Neurological_Condition', 'Chronic_Neurological_Condition_Description_Or_ICD10_codes', 'Chronic_Oxygen_Requirement', 'Chronic_Oxygen_Requirement_Description_Or_ICD10_codes', 'Inflammatory_Disease', 'Inflammatory_Disease_Description_Or_ICD10_codes', 'Viral_Infection', 'Viral_Infection_ICD10_codes_Or_Agents', 'Bacterial_Infection', 'Bacterial_Infection_ICD10_codes_Or_Agents', 'Cancer', 'Cancer_Description_Or_ICD10_codes', 'Substance_Abuse_Disorder', 'Substance_Abuse_Disorder_Description_Or_ICD10_codes', 'Organ_Transplant_Recipient', 'Organ_Transplant_Description_Or_ICD10_codes', 'Other_Health_Condition_Description_Or_ICD10_codes', 'ECOG_Status', 'Smoking_Or_Vaping_Status', 'Alcohol_Use', 'Drug_Type', 'Drug_Use', 'Vaccination_Record', 'Comments']
    covid_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'COVID_Status', 'Breakthrough_COVID', 'SARS-CoV-2_Variant', 'PCR_Test_Date_Duration_From_Index', 'Rapid_Antigen_Test_Date_Duration_From_Index', 'Antibody_Test_Date_Duration_From_Index', 'Symptomatic_COVID', 'Recovered_From_COVID', 'Duration_of_Disease', 'Recovery_Date_Duration_From_Index', 'Disease_Severity', 'Level_Of_Care', 'Symptoms', 'Other_Symptoms', 'COVID_complications', 'Long_COVID_symptoms', 'Other_Long_COVID_symptoms', 'COVID_Therapy', 'Comments']
    vax_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Vaccination_Status', 'SARS-CoV-2_Vaccine_Type', 'SARS-CoV-2_Vaccination_Date_Duration_From_Index', 'SARS-CoV-2_Vaccination_Side_Effects', 'Other_SARS-CoV-2_Vaccination_Side_Effects', 'Comments']
    meds_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Health_Condition_Or_Disease', 'Treatment', 'Dosage', 'Dosage_Units', 'Dosage_Regimen', 'Start_Date_Duration_From_Index', 'Stop_Date_Duration_From_Index', 'Update', 'Comments']
    auto_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Autoimmune_Condition', 'Autoimmune_Condition_ICD10_code', 'Year_Of_Diagnosis_Duration_to_Index', 'Antibody_Name', 'Antibody_Present', 'Update', 'Comments']
    trans_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Organ Transplant', 'Organ_Transplant_Other', 'Number_of_Hematopoietic_Cell_Transplants', 'Number_Of_Solid_Organ_Transplants', 'Date_of_Latest_Hematopoietic_Cell_Transplant_Duration_From_Index', 'Date_of_Latest_Solid_Organ_Transplant_Duration_From_Index', 'Update', 'Comments']
    cancer_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Cancer', 'ICD_10_Code', 'Year_Of_Diagnosis_Duration_From_Index', 'Cured', 'In_Remission', 'In_Unspecified_Therapy', 'Chemotherapy', 'Radiation Therapy', 'Surgery', 'Update', 'Comments']
    iris_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/IRIS/'
    titan_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/TITAN/'
    mars_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/MARS/'
    priority_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/PRIORITY/'

    # include_file = iris_folder + 'IRIS for D4 Long.xlsx'
    # include_file = mars_folder + 'MARS for D4 Long.xlsx'
    include_file = priority_folder + 'PRIORITY for D4 Long.xlsx'
    include_sheet = 'Biospecimen Ref' # People to include
    include_ppl = pd.read_excel(include_file, sheet_name=include_sheet)['Participant ID'].unique()
    participant_study = {}
    for participant in include_ppl:
        participant_study[participant] = 'PRIORITY' # once again, hacky
    first_date = parser.parse('1/1/2021').date()
    last_date = parser.parse('12/31/2021').date()

    # long_form = iris_folder + 'IRIS for D4 Long.xlsx'
    # long_form = mars_folder + 'MARS for D4 Long.xlsx'
    long_form = priority_folder + 'PRIORITY for D4 Long.xlsx'
    all_samples = pd.read_excel(long_form, sheet_name='Biospecimen Ref') # streamline this
    specimen_ids = all_samples['Biospecimen_ID'].unique() # samples to include 
    one_per = all_samples.drop_duplicates('Visit ID')
    one_per['Serum'] = one_per['Visit ID'].apply(lambda val: '{}_10{}'.format(val[:9], val[-1:]) in specimen_ids)
    one_per['PBMC'] = one_per['Visit ID'].apply(lambda val: '{}_20{}'.format(val[:9], val[-1:]) in specimen_ids)
    one_per['Biospecimens_Collected'] = one_per.apply(specimenize, axis=1)
    one_per.set_index('Visit ID', inplace=True)

    future_output = {}
    future_output['Baseline'] = {col: [] for col in base_cols}
    future_output['FollowUp'] = {col: [] for col in follow_cols}
    future_output['COVID'] = {col: [] for col in covid_cols}
    future_output['Vax'] = {col: [] for col in vax_cols}
    future_output['Meds'] = {col: [] for col in meds_cols}
    future_output['Auto'] = {col: [] for col in auto_cols}
    future_output['Transplant'] = {col: [] for col in trans_cols}
    # future_output['Cancer'] = {col: [] for col in cancer_cols}

    current_input = {}
    current_input['Baseline'] = pd.read_excel(long_form, sheet_name='Baseline Info', keep_default_na=False).set_index('Research_Participant_ID')
    current_input['COVID'] = pd.read_excel(long_form, sheet_name='COVID Infections', keep_default_na=False)
    current_input['Vax'] = pd.read_excel(long_form, sheet_name='COVID Vaccinations', keep_default_na=False)
    current_input['Meds'] = pd.read_excel(long_form, sheet_name='Medications', keep_default_na=False)
    # Refine this later for the multi-output
    # current_input['Transplant'] = pd.read_excel(long_form, sheet_name='Transplant-Specific', keep_default_na=False).set_index('Research_Participant_ID')
    # current_input['Cancer'] = pd.read_excel(long_form, sheet_name='Cancer-specific', keep_default_na=False).set_index('Research_Participant_ID')
    # print(current_input['Cancer'].head())
    current_input['Meds']['Reported'] = 'No'
    current_input['Vax']['Reported'] = 'No'
    current_input['COVID']['Reported'] = 'No'

    for visit_id, row_outer in one_per.sort_values('Visit Date').iterrows():
        visit = row_outer['Visit_Number']
        visit_date = row_outer['Visit Date'].date()
        participant = row_outer['Participant ID']
        study = participant_study[participant]
        seronet_id = row_outer['Research_Participant_ID']
        sample_id = row_outer['Sample ID']
        index_date = row_outer['Index Date'].date()
        days_from_index = int((visit_date - index_date).days)
        specimens = row_outer['Biospecimens_Collected']
        '''
        Baseline or Follow-Up
        '''
    # base_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Date_Duration_From_Index', 'Lost_to_Follow_Up', 'Final_Visit', 'Age', 'Sex_At_Birth', 'Race', 'Ethnicity', 'Height', 'Weight', 'BMI', 'Location', 'Biospecimens_Collected', 'Diabetes', 'Diabetes_Description_Or_ICD10_codes', 'Hypertension', 'Hypertension_Description_Or_ICD10_codes', 'Obesity', 'Obesity_Description_Or_ICD10_codes', 'Cardiovascular_Disease', 'Cardiovascular_Disease_Description_Or_ICD10_codes', 'Chronic_Lung_Disease', 'Chronic_Lung_Disease_Description_Or_ICD10_codes', 'Chronic_Kidney_Disease', 'Chronic_Kidney_Disease_Description_Or_ICD10_codes', 'Chronic_Liver_Disease', 'Chronic_Liver_Disease_Description_Or_ICD10_codes', 'Acute_Liver_Disease', 'Acute_Liver_Disease_Description_Or_ICD10_codes', 'Immunosuppressive_Condition', 'Immunosuppressive_Condition_Description_Or_ICD10_codes', 'Autoimmune_Disorder', 'Autoimmune_Disorder_Description_Or_ICD10_codes', 'Chronic_Neurological_Condition', 'Chronic_Neurological_Condition_Description_Or_ICD10_codes', 'Chronic_Oxygen_Requirement', 'Chronic_Oxygen_Requirement_Description_Or_ICD10_codes', 'Inflammatory_Disease', 'Inflammatory_Disease_Description_Or_ICD10_codes', 'Viral_Infection', 'Viral_Infection_ICD10_codes_Or_Agents', 'Bacterial_Infection', 'Bacterial_Infection_ICD10_codes_Or_Agents', 'Cancer', 'Cancer_Description_Or_ICD10_codes', 'Substance_Abuse_Disorder', 'Substance_Abuse_Disorder_Description_Or_ICD10_codes', 'Organ_Transplant_Recipient', 'Organ_Transplant_Description_Or_ICD10_codes', 'Other_Health_Condition_Description_Or_ICD10_codes', 'ECOG_Status', 'Smoking_Or_Vaping_Status', 'Alcohol_Use', 'Drug_Type', 'Drug_Use', 'Vaccination_Record', 'Comments']
    # follow_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Visit_Date_Duration_From_Index', 'Lost_to_Follow_Up', 'Final_Visit', 'Baseline_Visit', 'Number_of_Missed_Scheduled_Visits', 'Unscheduled_Visit', 'Biospecimens_Collected', 'Diabetes', 'Diabetes_Description_Or_ICD10_codes', 'Hypertension', 'Hypertension_Description_Or_ICD10_codes', 'Obesity', 'Obesity_Description_Or_ICD10_codes', 'Cardiovascular_Disease', 'Cardiovascular_Disease_Description_Or_ICD10_codes', 'Chronic_Lung_Disease', 'Chronic_Lung_Disease_Description_Or_ICD10_codes', 'Chronic_Kidney_Disease', 'Chronic_Kidney_Disease_Description_Or_ICD10_codes', 'Chronic_Liver_Disease', 'Chronic_Liver_Disease_Description_Or_ICD10_codes', 'Acute_Liver_Disease', 'Acute_Liver_Disease_Description_Or_ICD10_codes', 'Immunosuppressive_Condition', 'Immunosuppressive_Condition_Description_Or_ICD10_codes', 'Autoimmune_Disorder', 'Autoimmune_Disorder_Description_Or_ICD10_codes', 'Chronic_Neurological_Condition', 'Chronic_Neurological_Condition_Description_Or_ICD10_codes', 'Chronic_Oxygen_Requirement', 'Chronic_Oxygen_Requirement_Description_Or_ICD10_codes', 'Inflammatory_Disease', 'Inflammatory_Disease_Description_Or_ICD10_codes', 'Viral_Infection', 'Viral_Infection_ICD10_codes_Or_Agents', 'Bacterial_Infection', 'Bacterial_Infection_ICD10_codes_Or_Agents', 'Cancer', 'Cancer_Description_Or_ICD10_codes', 'Substance_Abuse_Disorder', 'Substance_Abuse_Disorder_Description_Or_ICD10_codes', 'Organ_Transplant_Recipient', 'Organ_Transplant_Description_Or_ICD10_codes', 'Other_Health_Condition_Description_Or_ICD10_codes', 'ECOG_Status', 'Smoking_Or_Vaping_Status', 'Alcohol_Use', 'Drug_Type', 'Drug_Use', 'Vaccination_Record', 'Comments']
        source_df = current_input['Baseline']
        if visit == 'Baseline(1)':
            sec = 'Baseline'
            add_to = future_output[sec]
            for col in ['Age', 'Sex_At_Birth', 'Race', 'Ethnicity', 'Height', 'Weight', 'BMI', 'Location']:
                if source_df.loc[seronet_id, col] == 'UNK':
                    add_to[col].append('')
                else:
                    add_to[col].append(source_df.loc[seronet_id, col])
        else:
            sec = 'FollowUp'
            add_to = future_output[sec]
            add_to['Visit_Number'].append(visit)
            add_to['Baseline_Visit'].append('No')
            add_to['Number_of_Missed_Scheduled_Visits'].append('N/A')
            add_to['Unscheduled_Visit'].append('No')
        add_to['Research_Participant_ID'].append(seronet_id)
        add_to['Cohort'].append(study)
        add_to['Visit_Date_Duration_From_Index'].append(days_from_index)
        add_to['Lost_to_Follow_Up'].append('No')
        add_to['Final_Visit'].append('No')
        add_to['Biospecimens_Collected'].append(specimens)
        for col in follow_cols[10:-2]: # hard coded
            if visit == 'Baseline(1)':
                add_to[col].append(source_df.loc[seronet_id, col])
            else:
                if "ICD10" in col:
                    add_to[col].append(source_df.loc[seronet_id, col])
                elif str(source_df.loc[seronet_id, col]).strip().upper() == 'YES':
                    add_to[col].append('Condition Status Unknown')
                else:
                    add_to[col].append('Not Reported')
        # This is the same so far, but my intuition is it may someday have to be different
        add_to['Vaccination_Record'].append('N/A')
        add_to['Comments'].append(source_df.loc[seronet_id, 'Comments'])
        '''
        COVID
        '''
    # covid_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'COVID_Status', 'Breakthrough_COVID', 'SARS-CoV-2_Variant', 'PCR_Test_Date_Duration_From_Index', 'Rapid_Antigen_Test_Date_Duration_From_Index', 'Antibody_Test_Date_Duration_From_Index', 'Symptomatic_COVID', 'Recovered_From_COVID', 'Duration_of_Disease', 'Recovery_Date_Duration_From_Index', 'Disease_Severity', 'Level_Of_Care', 'Symptoms', 'Other_Symptoms', 'COVID_complications', 'Long_COVID_symptoms', 'Other_Long_COVID_symptoms', 'COVID_Therapy', 'Comments']
        source_df = current_input['COVID']
        sec = 'COVID'
        add_to = future_output[sec]
        add_to['Research_Participant_ID'].append(seronet_id)
        add_to['Cohort'].append(study)
        add_to['Visit_Number'].append(visit)
        if participant not in source_df['Participant ID'].unique():
            add_to['COVID_Status'].append('No COVID Event Reported')
            for col in covid_cols[4:-1]:
                add_to[col].append('N/A')
            add_to['Comments'].append('')
        else:
            for idx, row in source_df[source_df['Participant ID'] == participant].sort_values('Report_Time').iterrows():
                if row['Reported'] == 'No' and row['Report_Time'] < visit_date:
                    source_df.loc[idx, 'Reported'] = 'Yes'
                    for col in covid_cols[3:-1]:
                        if 'Date' in col:
                            cropped_col = col[:-20] # drop "Duration_From_Index"
                            try:
                                add_to[col].append(int((row[cropped_col].date() - index_date).days))
                            except:
                                add_to[col].append(row[cropped_col])
                        else:
                            add_to[col].append(row[col])
                    add_to['Comments'].append('')
        while len(add_to['Visit_Number']) < len(add_to['COVID_Status']):
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(study)
            add_to['Visit_Number'].append(visit)
        if len(add_to['Visit_Number']) > len(add_to['COVID_Status']):
            add_to['COVID_Status'].append('No COVID Event Reported')
            for col in covid_cols[4:-1]:
                add_to[col].append('N/A')
            add_to['Comments'].append('')
        '''
        Vaccines
        '''
    # vax_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Vaccination_Status', 'SARS-CoV-2_Vaccine_Type', 'SARS-CoV-2_Vaccination_Date_Duration_From_Index', 'SARS-CoV-2_Vaccination_Side_Effects', 'Other_SARS-CoV-2_Vaccination_Side_Effects', 'Comments']
        source_df = current_input['Vax']
        sec = 'Vax'
        add_to = future_output[sec]
        add_to['Research_Participant_ID'].append(seronet_id)
        add_to['Cohort'].append(study)
        add_to['Visit_Number'].append(visit)
        if participant not in source_df['Participant ID'].unique():
            add_to['Vaccination_Status'].append('No vaccination event reported')
            for col in vax_cols[4:-1]:
                add_to[col].append('N/A')
            add_to['Comments'].append('')
        else:
            for idx, row in source_df[source_df['Participant ID'] == participant].iterrows():
                if row['Reported'] == 'No' and row['SARS-CoV-2_Vaccination_Date'].date() < visit_date:
                    source_df.loc[idx, 'Reported'] = 'Yes'
                    for col in vax_cols[3:-1]:
                        if 'Date' in col:
                            cropped_col = col[:-20] # drop "_From_Index"
                            if type(row[cropped_col]) in [pd.Timestamp, datetime.datetime]:
                                add_to[col].append(int((row[cropped_col].date() - index_date).days))
                            else:
                                add_to[col].append(row[cropped_col])
                        else:
                            add_to[col].append(row[col])
                    add_to['Comments'].append(row['Comments'])
        while len(add_to['Visit_Number']) < len(add_to['Vaccination_Status']):
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(study)
            add_to['Visit_Number'].append(visit)
        if len(add_to['Visit_Number']) > len(add_to['Vaccination_Status']):
            add_to['Vaccination_Status'].append('Unvaccinated')
            for col in vax_cols[4:-1]:
                add_to[col].append('N/A')
            add_to['Comments'].append('')
        '''
        Medications
        '''
    # meds_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Health_Condition_Or_Disease', 'Treatment', 'Dosage', 'Dosage_Units', 'Dosage_Regimen', 'Start_Date_Duration_From_Index', 'Stop_Date_Duration_From_Index', 'Update', 'Comments']
        source_df = current_input['Meds']
        sec = 'Meds'
        add_to = future_output[sec]
        # if participant not in source_df['Participant ID'].unique():
        #     add_to['Health_Condition_Or_Disease'].append('')
        #     add_to['Treatment'].append('Not Reported')
        #     for col in meds_cols[5:-2]:
        #         add_to[col].append('')
        #     if visit == 'Baseline(1)':
        #         add_to['Update'].append('Baseline Information')
        #     else:
        #         add_to['Update'].append('Not Reported')
        #     add_to['Comments'].append('No Treatment Reported')
        for idx, row in source_df[(source_df['Participant ID'] == participant) & source_df['Report_Time'].apply(lambda val: type(val) != str)].sort_values('Report_Time').iterrows():
            try:
                if row['Reported'] != 'Yes' and row['Report_Time'].date() < visit_date:
                    for col in meds_cols[3:-2]:
                        if 'Date' in col:
                            cropped_col = col[:-20] # drop "Duration_From_Index"
                            if cropped_col == 'Start_Date' and source_df.loc[idx, 'Reported'] == 'Partial':
                                add_to[col].append('Ongoing')
                            elif cropped_col == 'Stop_Date' and type(row[cropped_col]) == datetime.datetime and row[cropped_col].date() > visit_date:
                                add_to[col].append('Ongoing')
                            else:
                                try:
                                    if type(row[cropped_col]) == datetime.datetime:
                                        add_to[col].append(int((row[cropped_col].date() - index_date).days))
                                    else:
                                        add_to[col].append(row[cropped_col])
                                except:
                                    print(participant, visit, sample_id, col, row[cropped_col].date())
                                    add_to[col].append(row[cropped_col])
                        else:
                            add_to[col].append(row[col])
                    if visit == 'Baseline(1)':
                        add_to['Update'].append('Baseline Information')
                    else:
                        add_to['Update'].append('Not Reported')
                    add_to['Comments'].append(row['Comments'])
                    if type(row['Stop_Date']) == datetime.datetime and row['Stop_Date'].date() <= visit_date:
                        source_df.loc[idx, 'Reported'] = 'Yes'
                    else:
                        source_df.loc[idx, 'Reported'] = 'Partial'
            except:
                print("Error processing treatment info for", participant, "taking", row['Treatment'])
        while len(add_to['Visit_Number']) < len(add_to['Treatment']):
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(study)
            add_to['Visit_Number'].append(visit)
        '''
        Autoimmune (IRIS)
        '''
    # auto_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Autoimmune_Condition', 'Autoimmune_Condition_ICD10_code', 'Year_Of_Diagnosis_Duration_to_Index', 'Antibody_Name', 'Antibody_Present', 'Update', 'Comments']
        if study == 'IRIS':
            source_df = current_input['Baseline']
            sec = 'Auto'
            add_to = future_output[sec]
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(study)
            add_to['Visit_Number'].append(visit)
            add_to['Autoimmune_Condition'].append(source_df.loc[seronet_id, 'IBD Description'])
            add_to['Autoimmune_Condition_ICD10_code'].append(source_df.loc[seronet_id, 'IBD ICD-10'])
            add_to['Year_Of_Diagnosis_Duration_to_Index'].append('Not Reported')
            add_to['Antibody_Name'].append('N/A')
            add_to['Antibody_Present'].append('N/A')
            if visit == 'Baseline(1)':
                add_to['Update'].append('Baseline Information')
            else:
                add_to['Update'].append('No Update Reported')
            add_to['Comments'].append('')
        '''
        Transplant (TITAN)
        '''
    # trans_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Organ Transplant', 'Organ_Transplant_Other', 'Number_of_Hematopoietic_Cell_Transplants', 'Number_Of_Solid_Organ_Transplants', 'Date_of_Latest_Hematopoietic_Cell_Transplant_Duration_From_Index', 'Date_of_Latest_Solid_Organ_Transplant_Duration_From_Index', 'Update', 'Comments']
        if study == 'TITAN':
            sec = 'Transplant'
            source_df = current_input[sec]
            add_to = future_output[sec]
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(study)
            add_to['Visit_Number'].append(visit)
            for col in trans_cols[3:-2]:
                if 'Date' in col:
                    cropped_col = col[:-20] # drop "Duration_From_Index"
                    if source_df.loc[seronet_id, cropped_col] == "N/A":
                        add_to[col].append("N/A")
                    else:
                        add_to[col].append(int((source_df.loc[seronet_id, cropped_col].date() - index_date).days))
                else:
                    add_to[col].append(source_df.loc[seronet_id, col])
            if visit == 'Baseline(1)':
                add_to['Update'].append('Baseline Information')
            else:
                add_to['Update'].append('No Update Reported')
            add_to['Comments'].append('')
    # cancer_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Cancer', 'ICD_10_Code', 'Year_Of_Diagnosis_Duration_From_Index', 'Cured', 'In_Remission', 'In_Unspecified_Therapy', 'Chemotherapy', 'Radiation Therapy', 'Surgery', 'Update', 'Comments']
        if study == 'MARS':
            sec = 'Cancer'
            source_df = current_input[sec]
            add_to = future_output[sec]
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(study)
            add_to['Visit_Number'].append(visit)
            for col in cancer_cols[3:-2]:
                if 'Year' in col:
                    cropped_col = col[:-20] # drop "Duration_From_Index"
                    if source_df.loc[seronet_id, cropped_col].lower() == "not reported":
                        add_to[col].append("Not Reported")
                    else:
                        add_to[col].append(int(source_df.loc[seronet_id, cropped_col]) - index_date.year)
                else:
                    add_to[col].append(source_df.loc[seronet_id, col])
            if visit == 'Baseline(1)':
                add_to['Update'].append('Baseline Information')
            else:
                add_to['Update'].append('No Update Reported')
            add_to['Comments'].append(source_df.loc[seronet_id, 'Comments'])

    output_folder = priority_folder
    writer = pd.ExcelWriter(output_folder + 'CLINICAL FORMS Task D4 P4 WIP.xlsx') # Output sheet
    for sname, df in future_output.items():
        print(sname)
        for k, vs in df.items():
            print(k, len(vs))
        print()
        pd.DataFrame(df).to_excel(writer, sheet_name=sname, index=False)
    writer.save()
