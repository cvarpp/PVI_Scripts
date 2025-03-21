import pandas as pd
import datetime
import argparse
import util
import warnings

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

def write_clinical(input_df, output_fname, debug=False):
    base_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Date_Duration_From_Index', 'Lost_to_Follow_Up',
                 'Final_Visit', 'Age', 'Sex_At_Birth', 'Race', 'Ethnicity', 'Height', 'Weight', 'BMI', 'Location',
                 'Biospecimens_Collected',
                 'Diabetes', 'Diabetes_Provenance', 'Diabetes_Description_Or_ICD10_codes', 'Diabetes_Description_Or_ICD10_codes_Provenance', 'Hypertension', 'Hypertension_Provenance', 'Hypertension_Description_Or_ICD10_codes', 'Hypertension_Description_Or_ICD10_codes_Provenance', 'Obesity', 'Obesity_Provenance', 'Obesity_Description_Or_ICD10_codes', 'Obesity_Description_Or_ICD10_codes_Provenance', 'Cardiovascular_Disease', 'Cardiovascular_Disease_Provenance', 'Cardiovascular_Disease_Description_Or_ICD10_codes', 'Cardiovascular_Disease_Description_Or_ICD10_codes_Provenance', 'Chronic_Lung_Disease', 'Chronic_Lung_Disease_Provenance', 'Chronic_Lung_Disease_Description_Or_ICD10_codes', 'Chronic_Lung_Disease_Description_Or_ICD10_codes_Provenance', 'Chronic_Kidney_Disease', 'Chronic_Kidney_Disease_Provenance', 'Chronic_Kidney_Disease_Description_Or_ICD10_codes', 'Chronic_Kidney_Disease_Description_Or_ICD10_codes_Provenance', 'Chronic_Liver_Disease', 'Chronic_Liver_Disease_Provenance', 'Chronic_Liver_Disease_Description_Or_ICD10_codes', 'Chronic_Liver_Disease_Description_Or_ICD10_codes_Provenance', 'Acute_Liver_Disease', 'Acute_Liver_Disease_Provenance', 'Acute_Liver_Disease_Description_Or_ICD10_codes', 'Acute_Liver_Disease_Description_Or_ICD10_codes_Provenance', 'Immunosuppressive_Condition', 'Immunosuppressive_Condition_Provenance', 'Immunosuppressive_Condition_Description_Or_ICD10_codes', 'Immunosuppressive_Condition_Description_Or_ICD10_codes_Provenance', 'Autoimmune_Disorder', 'Autoimmune_Disorder_Provenance', 'Autoimmune_Disorder_Description_Or_ICD10_codes', 'Autoimmune_Disorder_Description_Or_ICD10_codes_Provenance', 'Chronic_Neurological_Condition', 'Chronic_Neurological_Condition_Provenance', 'Chronic_Neurological_Condition_Description_Or_ICD10_codes', 'Chronic_Neurological_Condition_Description_Or_ICD10_codes_Provenance', 'Chronic_Oxygen_Requirement', 'Chronic_Oxygen_Requirement_Provenance', 'Chronic_Oxygen_Requirement_Description_Or_ICD10_codes', 'Chronic_Oxygen_Requirement_Description_Or_ICD10_codes_Provenance', 'Inflammatory_Disease', 'Inflammatory_Disease_Provenance', 'Inflammatory_Disease_Description_Or_ICD10_codes', 'Inflammatory_Disease_Description_Or_ICD10_codes_Provenance', 'Viral_Infection', 'Viral_Infection_ICD10_codes_Or_Agents', 'Bacterial_Infection', 'Bacterial_Infection_ICD10_codes_Or_Agents',
                 'Cancer', 'Cancer_Provenance', 'Cancer_Description_Or_ICD10_codes', 'Cancer_Description_Or_ICD10_codes_Provenance',
                 'Substance_Abuse_Disorder', 'Substance_Abuse_Disorder_Description_Or_ICD10_codes', 'Organ_Transplant_Recipient',
                 'Organ_Transplant_Description_Or_ICD10_codes', 'Other_Health_Condition_Description_Or_ICD10_codes', 'Other_Health_Condition_Description_Or_ICD10_codes_Provenance', 'ECOG_Status',
                 'Smoking_Or_Vaping_Status', 'Alcohol_Use', 'Drug_Type', 'Drug_Use', 'Vaccination_Record', 'Pregnancy_Status', 'Comments']
    follow_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Visit_Date_Duration_From_Index', 'Lost_to_Follow_Up', 'Final_Visit', 'Baseline_Visit', 'Number_of_Missed_Scheduled_Visits', 'Unscheduled_Visit', 'Biospecimens_Collected', 'Diabetes', 'Diabetes_Description_Or_ICD10_codes', 'Hypertension', 'Hypertension_Description_Or_ICD10_codes', 'Obesity', 'Obesity_Description_Or_ICD10_codes', 'Cardiovascular_Disease', 'Cardiovascular_Disease_Description_Or_ICD10_codes', 'Chronic_Lung_Disease', 'Chronic_Lung_Disease_Description_Or_ICD10_codes', 'Chronic_Kidney_Disease', 'Chronic_Kidney_Disease_Description_Or_ICD10_codes', 'Chronic_Liver_Disease', 'Chronic_Liver_Disease_Description_Or_ICD10_codes', 'Acute_Liver_Disease', 'Acute_Liver_Disease_Description_Or_ICD10_codes', 'Immunosuppressive_Condition', 'Immunosuppressive_Condition_Description_Or_ICD10_codes', 'Autoimmune_Disorder', 'Autoimmune_Disorder_Description_Or_ICD10_codes', 'Chronic_Neurological_Condition', 'Chronic_Neurological_Condition_Description_Or_ICD10_codes', 'Chronic_Oxygen_Requirement', 'Chronic_Oxygen_Requirement_Description_Or_ICD10_codes', 'Inflammatory_Disease', 'Inflammatory_Disease_Description_Or_ICD10_codes', 'Viral_Infection', 'Viral_Infection_ICD10_codes_Or_Agents', 'Bacterial_Infection', 'Bacterial_Infection_ICD10_codes_Or_Agents', 'Cancer', 'Cancer_Provenance', 'Cancer_Description_Or_ICD10_codes', 'Cancer_Description_Or_ICD10_codes_Provenance', 'Substance_Abuse_Disorder', 'Substance_Abuse_Disorder_Description_Or_ICD10_codes', 'Organ_Transplant_Recipient', 'Organ_Transplant_Description_Or_ICD10_codes', 'Other_Health_Condition_Description_Or_ICD10_codes', 'ECOG_Status', 'Smoking_Or_Vaping_Status', 'Alcohol_Use', 'Drug_Type', 'Drug_Use', 'Vaccination_Record', 'Pregnancy_Status', 'Comments']
    covid_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'COVID_Status', 'Breakthrough_COVID', 'SARS-CoV-2_Variant', 'PCR_Test_Date_Duration_From_Index', 'Rapid_Antigen_Test_Date_Duration_From_Index', 'Antibody_Test_Date_Duration_From_Index', 'Symptomatic_COVID', 'Recovered_From_COVID', 'Duration_of_Disease', 'Recovery_Date_Duration_From_Index', 'Disease_Severity', 'Level_Of_Care', 'Symptoms', 'Other_Symptoms', 'COVID_complications', 'Long_COVID_symptoms', 'Other_Long_COVID_symptoms', 'COVID_Therapy', 'Comments']
    vax_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Vaccination_Status', 'SARS-CoV-2_Vaccine_Lot_Number', 'SARS-CoV-2_Vaccine_Type', 'SARS-CoV-2_Vaccination_Date_Duration_From_Index', 'SARS-CoV-2_Vaccination_Side_Effects', 'Other_SARS-CoV-2_Vaccination_Side_Effects', 'Comments']
    meds_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Health_Condition_Or_Disease', 'Treatment', 'Dosage', 'Dosage_Units', 'Dosage_Regimen', 'Start_Date_Duration_From_Index', 'Stop_Date_Duration_From_Index', 'Update', 'Comments']
    auto_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Autoimmune_Condition', 'Autoimmune_Condition_ICD10_code', 'Year_Of_Diagnosis_Duration_to_Index', 'Antibody_Name', 'Antibody_Present', 'Update', 'Comments']
    trans_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Organ Transplant', 'Organ_Transplant_Other', 'Number_of_Hematopoietic_Cell_Transplants', 'Number_Of_Solid_Organ_Transplants', 'Date_of_Latest_Hematopoietic_Cell_Transplant_Duration_From_Index', 'Date_of_Latest_Solid_Organ_Transplant_Duration_From_Index', 'Update', 'Comments']
    cancer_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Cancer', 'Cancer_Provenance', 'ICD_10_Code', 'ICD_10_Code_Provenance', 'Year_Of_Diagnosis_Duration_From_Index', 'Year_Of_Diagnosis_Duration_From_Index_Provenance', 'Cured', 'Cured_Provenance', 'In_Remission', 'In_Remission_Provenance', 'In_Unspecified_Therapy', 'In_Unspecified_Therapy_Provenance', 'Chemotherapy', 'Chemotherapy_Provenance', 'Radiation Therapy', 'Radiation Therapy_Provenance', 'Surgery', 'Surgery_Provenance', 'Update', 'Comments']

    iris_data = util.iris_folder + 'IRIS for D4 Long.xlsx'
    mars_data = util.mars_folder + 'MARS for D4 Long.xlsx'
    titan_data = util.titan_folder + 'TITAN for D4 Long.xlsx'
    priority_data = util.prio_folder + 'PRIORITY for D4 Long.xlsx'
    gaea_data = util.gaea_folder + 'GAEA for D4 Long.xlsx'
    participant_study = input_df.drop_duplicates(subset='Participant ID').set_index('Participant ID')['Cohort']

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*", module='openpyxl')
        exclusions = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name='Exclusions')
    exclude_ppl = set(exclusions['Participant ID'].unique())
    exclude_ids = set(exclusions['Research_Participant_ID'].unique())
    exclude_filter = (input_df['Participant ID'].apply(lambda val: val not in exclude_ppl) & 
                        input_df['Research_Participant_ID'].apply(lambda val: val not in exclude_ids))
    all_samples = input_df[exclude_filter].copy()
    specimen_ids = all_samples['Biospecimen_ID'].unique() # samples to include 
    one_per = all_samples.drop_duplicates(subset=['Research_Participant_ID', 'Visit_Number']).copy()
    one_per['Serum'] = one_per.apply(lambda row: '{}_1{}'.format(row['Research_Participant_ID'], str(row['Visit_Number']).strip("bBaseline()").zfill(2)) in specimen_ids, axis=1)
    one_per['PBMC'] = one_per.apply(lambda row: '{}_2{}'.format(row['Research_Participant_ID'], str(row['Visit_Number']).strip("bBaseline()").zfill(2)) in specimen_ids, axis=1)
    one_per['Biospecimens_Collected'] = one_per.apply(specimenize, axis=1)
    one_per.set_index('Sample ID', inplace=True)

    future_output = {}
    future_output['Baseline'] = {col: [] for col in base_cols}
    future_output['FollowUp'] = {col: [] for col in follow_cols}
    future_output['COVID'] = {col: [] for col in covid_cols}
    future_output['Vax'] = {col: [] for col in vax_cols}
    future_output['Meds'] = {col: [] for col in meds_cols}
    future_output['Auto'] = {col: [] for col in auto_cols}
    future_output['Transplant'] = {col: [] for col in trans_cols}
    future_output['Cancer'] = {col: [] for col in cancer_cols}
    for df2b in future_output.values():
        df2b['Sample ID'] = []

    current_input = {}
    cohort_df_names = [iris_data, mars_data, titan_data, priority_data, gaea_data]

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*", module='openpyxl')
        current_input['Baseline'] = pd.concat([pd.read_excel(df_name, sheet_name='Baseline Info', keep_default_na=False).set_index('Research_Participant_ID') for df_name in cohort_df_names])
        current_input['COVID'] = pd.concat([pd.read_excel(df_name, sheet_name='COVID Infections', keep_default_na=False) for df_name in cohort_df_names]).reset_index()
        current_input['Vax'] = pd.concat([pd.read_excel(df_name, sheet_name='COVID Vaccinations', keep_default_na=False) for df_name in cohort_df_names]).reset_index()
        current_input['Meds'] = pd.concat([pd.read_excel(df_name, sheet_name='Medications', keep_default_na=False) for df_name in cohort_df_names]).reset_index()
        current_input['Transplant'] = pd.read_excel(titan_data, sheet_name='Transplant-Specific', keep_default_na=False).set_index('Research_Participant_ID')
        current_input['Cancer'] = pd.read_excel(mars_data, sheet_name='Cancer-specific', keep_default_na=False).set_index('Research_Participant_ID')
    current_input['Meds']['Reported'] = 'No'
    current_input['Vax']['Reported'] = 'No'
    current_input['COVID']['Reported'] = 'No'

    for sample_id, row_outer in one_per.sort_values('Date').iterrows():
        seronet_id = row_outer['Research_Participant_ID']
        visit = row_outer['Visit_Number']
        visit_date = row_outer['Date']
        participant = row_outer['Participant ID']
        study = participant_study[participant]
        # index_date = row_outer['Index Date']
        # days_from_index = int((visit_date - index_date).days)
        days_from_index = row_outer['Biospecimen_Collection_Date_Duration_From_Index']
        index_date = visit_date - datetime.timedelta(days=days_from_index)
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
        add_to['Sample ID'].append(sample_id)
        add_to['Research_Participant_ID'].append(seronet_id)
        add_to['Cohort'].append(study)
        add_to['Visit_Date_Duration_From_Index'].append(days_from_index)
        add_to['Lost_to_Follow_Up'].append('No')
        add_to['Final_Visit'].append('No')
        add_to['Biospecimens_Collected'].append(specimens)
        for col in base_cols[14:-3]: # hard coded
            if visit == 'Baseline(1)':
                if 'Provenance' in col:
                    if study == 'GAEA':
                        add_to[col].append('Self Reported')
                    else:
                        add_to[col].append('EMR')
                else:
                    add_to[col].append(source_df.loc[seronet_id, col])
            elif col in follow_cols:
                if 'Provenance' in col:
                    if study == 'GAEA':
                        add_to[col].append('Self Reported')
                    else:
                        add_to[col].append('EMR')
                elif "ICD10" in col:
                    if source_df.loc[seronet_id, col] == 'N/A':
                        add_to[col].append('Not Reported')
                    else:
                        add_to[col].append(source_df.loc[seronet_id, col])
                elif str(source_df.loc[seronet_id, col]).strip().upper() == 'YES':
                    add_to[col].append('Condition Status Unknown')
                else:
                    add_to[col].append('Not Reported')
        # This is the same so far, but my intuition is it may someday have to be different
        add_to['Vaccination_Record'].append('N/A')
        add_to['Pregnancy_Status'].append('Not Reported')
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
                if row['Reported'] == 'No' and pd.to_datetime(row['Report_Time'], errors='coerce')  < visit_date:
                    source_df.loc[idx, 'Reported'] = 'Yes'
                    for col in covid_cols[3:-1]:
                        if 'Date' in col:
                            cropped_col = col[:-20] # drop "Duration_From_Index"
                            try:
                                add_to[col].append(int((row[cropped_col] - index_date).days))
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
        while len(add_to['Research_Participant_ID']) > len(add_to['Sample ID']):
            add_to['Sample ID'].append(sample_id)
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
                if row['Reported'] == 'No' and pd.to_datetime(row['SARS-CoV-2_Vaccination_Date'], errors='coerce') < visit_date:
                    source_df.loc[idx, 'Reported'] = 'Yes'
                    for col in vax_cols[3:-1]:
                        if 'Date' in col:
                            cropped_col = col[:-20] # drop "_From_Index"
                            try:
                                add_to[col].append(int((row[cropped_col] - index_date).days))
                            except:
                                if row[cropped_col] in ["N/A", "Unknown", "Not Reported"]:
                                    add_to[col].append(row[cropped_col])
                                else:
                                    print(row[cropped_col], "invalid for", sample_id, ". Fatal error, exiting")
                                    exit(1)
                        else:
                            add_to[col].append(row[col])
                    add_to['Comments'].append(row['Comments'])
        while len(add_to['Visit_Number']) < len(add_to['Vaccination_Status']):
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(study)
            add_to['Visit_Number'].append(visit)
        if len(add_to['Visit_Number']) > len(add_to['Vaccination_Status']):
            if (len(add_to['Vaccination_Status']) == 0 or
                (add_to['Research_Participant_ID'][len(add_to['Vaccination_Status']) - 1] == seronet_id and add_to['Vaccination_Status'][-1] == 'Unvaccinated')):
                add_to['Vaccination_Status'].append('Unvaccinated')
            else:
                add_to['Vaccination_Status'].append('No vaccination event reported')
            for col in vax_cols[4:-1]:
                add_to[col].append('N/A')
            add_to['Comments'].append('')
        while len(add_to['Research_Participant_ID']) > len(add_to['Sample ID']):
            add_to['Sample ID'].append(sample_id)
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
                if row['Reported'] != 'Yes' and pd.to_datetime(row['Report_Time'], errors='coerce') < visit_date:
                    for col in meds_cols[3:-2]:
                        if 'Date' in col:
                            cropped_col = col[:-20] # drop "Duration_From_Index"
                            if cropped_col == 'Start_Date' and source_df.loc[idx, 'Reported'] == 'Partial':
                                add_to[col].append('Ongoing')
                            elif cropped_col == 'Stop_Date' and type(row[cropped_col]) == datetime.datetime and row[cropped_col] > visit_date:
                                add_to[col].append('Ongoing')
                            else:
                                try:
                                    if type(row[cropped_col]) == datetime.datetime:
                                        add_to[col].append(int((row[cropped_col] - index_date).days))
                                    else:
                                        add_to[col].append(row[cropped_col])
                                except:
                                    print(participant, visit, sample_id, col, row[cropped_col])
                                    add_to[col].append(row[cropped_col])
                        else:
                            add_to[col].append(row[col])
                    if visit == 'Baseline(1)':
                        add_to['Update'].append('Baseline Information')
                    else:
                        add_to['Update'].append('Not Reported')
                    add_to['Comments'].append(row['Comments'])
                    if type(row['Stop_Date']) == datetime.datetime and row['Stop_Date'] <= visit_date:
                        source_df.loc[idx, 'Reported'] = 'Yes'
                    else:
                        source_df.loc[idx, 'Reported'] = 'Partial'
            except Exception as e:
                print("Error processing treatment info for", participant, "taking", row['Treatment'])
                print("Fatal Error. Exiting...")
                exit(-1)
        while len(add_to['Visit_Number']) < len(add_to['Treatment']):
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(study)
            add_to['Visit_Number'].append(visit)
        while len(add_to['Research_Participant_ID']) > len(add_to['Sample ID']):
            add_to['Sample ID'].append(sample_id)
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
            add_to['Sample ID'].append(sample_id)
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
                    if str(source_df.loc[seronet_id, cropped_col]) == "N/A":
                        add_to[col].append("N/A")
                    else:
                        try:
                            add_to[col].append(int((pd.to_datetime(source_df.loc[seronet_id, cropped_col]) - index_date).days))
                        except:
                            add_to[col].append("N/A")
                else:
                    add_to[col].append(source_df.loc[seronet_id, col])
            if visit == 'Baseline(1)':
                add_to['Update'].append('Baseline Information')
            else:
                add_to['Update'].append('No Update Reported')
            add_to['Comments'].append('')
            add_to['Sample ID'].append(sample_id)
    # cancer_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Cancer', 'ICD_10_Code', 'Year_Of_Diagnosis_Duration_From_Index', 'Cured', 'In_Remission', 'In_Unspecified_Therapy', 'Chemotherapy', 'Radiation Therapy', 'Surgery', 'Update', 'Comments']
        if study == 'MARS':
            sec = 'Cancer'
            source_df = current_input[sec]
            add_to = future_output[sec]
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(study)
            add_to['Visit_Number'].append(visit)
            for col in cancer_cols[3:-2]:
                if 'Provenance' in col:
                    add_to[col].append('EMR')
                elif 'Year' in col:
                    cropped_col = col[:-20] # drop "Duration_From_Index"
                    if str(source_df.loc[seronet_id, cropped_col]).lower() == "not reported":
                        add_to[col].append("Not Reported")
                    else:
                        try:
                            add_to[col].append(int(source_df.loc[seronet_id, cropped_col]) - index_date.year)
                        except:
                            add_to[col].append('Not Reported')
                else:
                    add_to[col].append(source_df.loc[seronet_id, col])
            if visit == 'Baseline(1)':
                add_to['Update'].append('Baseline Information')
            else:
                add_to['Update'].append('No Update Reported')
            add_to['Comments'].append(source_df.loc[seronet_id, 'Comments'])
            add_to['Sample ID'].append(sample_id)

    output_folder = util.cross_d4
    output = {}
    for sname, df2b in future_output.items():
        try:
            df = pd.DataFrame(df2b)
            output[sname] = df
        except:
            print(sname)
            for k, vs in df2b.items():
                print(k, len(vs))
            print()
    if not debug:
        with pd.ExcelWriter(output_folder + '{}.xlsx'.format(output_fname)) as writer:
            for sname, df in output.items():
                df.to_excel(writer, sheet_name=sname, index=False, na_rep='N/A')
    return output


if __name__ == '__main__':
    
    argParser = argparse.ArgumentParser(description='Make clinical forms. input_df should be a fully specified filepath to a workbook containing a "Biospecimen" tab corresponding to the correct data submission.')
    argParser.add_argument('-i', '--input_df', action='store', required=True, type=lambda wb_path: pd.read_excel(wb_path, sheet_name='Biospecimen'))
    argParser.add_argument('-o', '--output_file', action='store', required=True)
    argParser.add_argument('-d', '--debug', action='store_true')
    args = argParser.parse_args()

    write_clinical(args.input_df, args.output_file, args.debug)
