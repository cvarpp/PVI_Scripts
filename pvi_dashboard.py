import pandas as pd
import numpy as np
from datetime import date
import util
from helpers import permissive_datemax, query_intake, query_dscf, immune_history
import warnings
import cam_convert

def load_and_clean_sheet(excel_file, sheet_name, index_col='Participant ID', **excel_kwargs):
    df = pd.read_excel(excel_file, sheet_name=sheet_name, **excel_kwargs)
    if index_col != 'Participant ID':
        df = df.rename(columns={index_col: 'Participant ID'})
    df = df.dropna(subset=['Participant ID'])
    df['Participant ID'] = df['Participant ID'].astype(str).str.strip().str.upper()
    return df.set_index('Participant ID')

def pull_cohorts():
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*")
        #CRP
        crp_data = (
            load_and_clean_sheet(util.crp_folder + 'CRP Patient Tracker no circ ref 7.7.23.xlsx', sheet_name='Tracker', header=4)
            .rename(columns={'COVID-19 Vaccine Type':'Dose #1 Type', 'Vaccine #1 Date':'Dose #1 Date', 'Vaccine #2 Date':'Dose #2 Date',
                                '3rd Dose Vaccine Type':'Dose #3 Type', '3rd Dose Vaccine Date':'Dose #3 Date', '4th Dose Vaccine Type':'Dose #4 Type',
                                '4th Dose Date':'Dose #4 Date', '5th Dose Vaccine Type':'Dose #5 Type', '5th Dose Date':'Dose #5 Date',
                                'Study Participation Status':'Status', 'Age at Visit':'Age'})
            .assign(**{'Study': 'CRP', 'Dose #2 Type': lambda df: df['Dose #1 Type']})
        )
        #ROBIN
        robin_data = pd.read_excel(util.umbrella + 'ROBIN/ROBIN Participant tracker.xlsx', sheet_name='Participant Info', index_col='Subject ID')
        #SHIELD
        shield_data = load_and_clean_sheet(util.projects +'SHIELD/SHIELD tracker.xlsx', sheet_name='Summary')
        shield_data=shield_data[shield_data.index.astype(str).str.startswith('1')].copy()
        #GAEA
        gaea_tracker = pd.ExcelFile(util.gaea_folder + 'GAEA Tracker.xlsx')
        gaea_data = load_and_clean_sheet(gaea_tracker, sheet_name='Summary')
        gaea_infections = load_and_clean_sheet(gaea_tracker, sheet_name='Infections')
        gaea_data = (
            gaea_data.join(gaea_infections, on='Participant ID', rsuffix='_inf')
            .rename(columns={'3rd Vaccine Date':'Dose #3 Date', '4th Vaccine Date':'Dose #4 Date', '5th Vaccine Date':'Dose #5 Date', 
                                '6th Vaccine Date':'Dose #6 Date', '7th Vaccine Date':'Dose #7 Date', '8th Vaccine Date':'Dose #8 Date',
                                '9th Vaccine Date':'Dose #9 Date', 'Vaccine Type':'Dose #1 Type', '3rd Vaccine Type':'Dose #3 Type', 
                                '4th Vaccine Type':'Dose #4 Type', '5th Vaccine Type':'Dose #5 Type', '6th Vaccine Type':'Dose #6 Type', 
                                '7th Vaccine Type':'Dose #7 Type', '8th Vaccine Type':'Dose #8 Type', '9th Vaccine Type':'Dose #9 Type',
                                'Symptom Onset Date 1':'Infection 1 Date', 'Symptom Onset Date 2':'Infection 2 Date', 'Symptom Onset Date 3':'Infection 3 Date',
                                'Age @ Baseline':'Age'})
            .assign(**{'Study': 'GAEA', 'Dose #2 Type': lambda df: df['Dose #1 Type']})
        )

        #MARS
        mars_tracker = pd.ExcelFile(util.mars_folder + 'MARS tracker.xlsx')
        mars_data = load_and_clean_sheet(mars_tracker, sheet_name='Pt List')
        mars_infections = load_and_clean_sheet(mars_tracker, sheet_name='COVID Infections')
        mars_data = (
            mars_data.join(mars_infections, on='Participant ID', rsuffix='inf')
            .rename(columns={'Vaccine #1 Date':'Dose #1 Date', '1st Vaccine Type':'Dose #1 Type', 'Vaccine #2 Date':'Dose #2 Date', 
                                '2nd Vaccine Type':'Dose #2 Type', '3rd Vaccine':'Dose #3 Date', '3rd Vaccine Type':'Dose #3 Type',
                                '4th vaccine':'Dose #4 Date', '4th Vaccine Type':'Dose #4 Type', '5th vaccine':'Dose #5 Date', 
                                '5th Vaccine Type':'Dose #5 Type', '6th vaccine':'Dose #6 Date', '6th vaccine type':'Dose #6 Type',
                                '7th vaccine':'Dose #7 Date', '7th vaccine type':'Dose #7 Type', '8th vaccine':'Dose #8 Date', 
                                '8th vaccine type':'Dose #8 Type', '9th vaccine':'Dose #9 Date', '9th vaccine type':'Dose #9 Type',
                                'Study Status':'Status', 'T1':'Baseline Date', 'T1 Sample ID':'Baseline Sample ID', 'Age @ time of consent':'Age'})
            .assign(Study='MARS')
        )
        #IRIS
        iris_data = (
            load_and_clean_sheet(util.iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4)
            .rename(columns={'First Dose Date':'Dose #1 Date', 'Which Vaccine?':'Dose #1 Type', 'Second Dose Date':'Dose #2 Date', 
                                'Third Dose Date':'Dose #3 Date', 'Third Dose Type':'Dose #3 Type',
                                'Fourth Dose Date':'Dose #4 Date', 'Fourth Dose Type':'Dose #4 Type', 'Fifth Dose Date':'Dose #5 Date', 
                                'Fifth Dose Type':'Dose #5 Type', 'Sixth Dose Date':'Dose #6 Date', 'Sixth Dose Type':'Dose #6 Type',
                                'Seventh Dose Date':'Dose #7 Date', 'Seventh Dose Type':'Dose #7 Type', 'Eighth Dose Date':'Dose #8 Date', 
                                'Eighth Dose Type':'Dose #8 Type', 'Ninth Dose Date':'Dose #9 Date', 'Ninth Dose Type':'Dose #9 Type',
                                'Symptom Onset':'Infection 1 Date', 'Symptom Onset 2':'Infection 2 Date', 'Symptom Onset 3':'Infection 3 Date',
                                'Sample ID':'Baseline Sample ID'})
            .assign(**{'Study': 'IRIS',
                       'Status': 'PULL FROM UMBRELLA TRACKER', #TODO: pull
                       'Date of Consent': 'PULL FROM UMBRELLA TRACKER', #TODO: pull
                       'Dose #2 Type': lambda df: df['Dose #1 Type']})
        )
        #TITAN
        titan_tracker = pd.ExcelFile(util.titan_tracker)
        titan_data = (
            load_and_clean_sheet(titan_tracker, sheet_name='Tracker', index_col='Umbrella Corresponding Participant ID', header=4)
            .rename(columns={'Gender':'Sex', 'Study Participation Status':'Status', 'Age at Enrollment':'Age'})
        )
        titan_infections = load_and_clean_sheet(titan_tracker, sheet_name='Infections Combined')
        titan_dems = load_and_clean_sheet(titan_tracker, sheet_name='Demographics & First Two Doses', index_col='Umbrella Participant ID', header=3)
        titan_data = (
            titan_data.join(titan_infections, on='Participant ID', rsuffix='inf')
            .join(titan_dems, on='Participant ID', rsuffix='dem')
            .rename(columns={'Vaccine #1 Date':'Dose #1 Date', 'Vaccine Type':'Dose #1 Type', 'Vaccine #2 Date':'Dose #2 Date', 
                                '3rd Dose Vaccine Date':'Dose #3 Date', '3rd Dose Vaccine Type':'Dose #3 Type',
                                'First Booster Dose Date (#4)':'Dose #4 Date', 'First Booster Vaccine Type (#4)':'Dose #4 Type', 'Second Booster Dose Date (#5)':'Dose #5 Date', 
                                'Second Booster Vaccine Type (#5)':'Dose #5 Type', 'Third Booster Dose Date (#6)':'Dose #6 Date', 'Third Booster Vaccine Type (#6)':'Dose #6 Type'})
            .assign(**{'Study': 'TITAN', 'Dose #2 Type': lambda df: df['Dose #1 Type']})
        )

    cohort_data = {
        'TITAN': titan_data,
        'IRIS': iris_data,
        'MARS': mars_data,
        'GAEA': gaea_data,
        'CRP': crp_data,
        'ROBIN': robin_data,
        'SHIELD': shield_data,
    }
    return cohort_data

def pull_trackers():
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*")

        # APOLLO
        apollo_tracker = pd.ExcelFile(util.apollo_tracker)
        # Load and clean each sheet
        apollo_data = load_and_clean_sheet(apollo_tracker, 'Summary')
        apollo_vaccines = load_and_clean_sheet(apollo_tracker, 'Vaccinations')
        apollo_infections = load_and_clean_sheet(apollo_tracker, 'Infections')
        # Join the dataframes
        apollo_data = (
            apollo_data
            .join(apollo_vaccines, how='inner', rsuffix='_vax')
            .join(apollo_infections, how='inner', rsuffix='_inf')
            .assign(Study='APOLLO')
        )

        apollo_data.rename(columns={'Age @ Enrollment':'Age', 'Number of Vaccine Doses':'Total Vaccine Doses', 
                                    'Number of SARS-CoV-2 Infections':'Total Infections','Symptom Onset Date 1':'Infection 1 Date', 
                                    'Symptom Onset Date 2':'Infection 2 Date', 'Symptom Onset Date 3':'Infection 3 Date', 
                                    'Symptom Onset Date 4':'Infection 4 Date'}, inplace=True)
    
        # PARIS
        paris_tracker = pd.ExcelFile(util.paris + 'Patient Tracking - PARIS.xlsx')
        paris_data = load_and_clean_sheet(paris_tracker, sheet_name='Subgroups', header=4)
        paris_main = load_and_clean_sheet(paris_tracker, sheet_name='Main', index_col='Subject ID', header=8)
        paris_dems = load_and_clean_sheet(util.projects + 'PARIS/Demographics.xlsx', sheet_name='inputs', index_col='Subject ID')
        paris_data = (
            paris_data.rename(columns={'Study Status': 'Status'})
            .join(paris_dems, how='left', rsuffix='_dem')
            .join(paris_main, how='left', rsuffix='_main')
            .assign(Study='PARIS')
        )

        paris_data.rename(columns={'E-mail_main':'Email', 'Date of Birth':'DOB', 'Gender':'Sex', 'Status':'Status_sg', 'Study Status':'Status',
                                'Sample ID_main':'Baseline Sample ID', 'First Dose Date':'Dose #1 Date', 'Second Dose Date':'Dose #2 Date', 
                                'Boost Date':'Dose #3 Date', 'Boost 2 Date':'Dose #4 Date', 'Boost 3 Date':'Dose #5 Date', 
                                'Boost 4 Date':'Dose #6 Date', 'Boost 5 Date':'Dose #7 Date', 'Boost 6 Date':'Dose #8 Date',
                                'Vaccine Type':'Dose #1 Type', 'Boost Type':'Dose #3 Type', 'Boost 2 Type':'Dose #4 Type', 'Boost 3 Type':'Dose #5 Type', 
                                'Boost 4 Type':'Dose #6 Type', 'Boost 5 Type':'Dose #7 Type', 'Boost 6 Type':'Dose #8 Type'}, inplace=True)
        paris_data['Dose #2 Type'] = paris_data['Dose #1 Type'].copy()
        paris_data['Ethnicity'] = paris_data['Ethnicity: Hispanic or Latino'].apply(lambda val: 'Hispanic or Latino' if val == 'Yes' else 'Not Hispanic or Latino')

        # HD
        hd_data = (
            load_and_clean_sheet(util.clin_ops + 'Healthy donors/Healthy Donors Participant Tracker.xlsx', sheet_name='Participants')
            .rename(columns={'Date of consent':'Date of Consent', 'Age at Baseline':'Age'})
            .assign(**{'Study': 'Healthy Donors', 'Baseline Date': lambda df: df['Date of Consent']})
        )
        # DOVE
        dove_data= load_and_clean_sheet(util.clin_ops + 'Healthy donors/DOVE/DOVE participant tracker.xlsx', sheet_name='Participant info')
        hd_data = hd_data.join(dove_data, rsuffix='_dove')
        #UMBRELLA
        umbrella_tracker = pd.ExcelFile(util.umbrella_tracker)
        umb_data = load_and_clean_sheet(umbrella_tracker, sheet_name='Summary', index_col='Subject ID')
        umb_vaccines = load_and_clean_sheet(umbrella_tracker, sheet_name='COVID-19 Vaccine Type & Dates!!')
        cohorts = pull_cohorts()
        umb_data = (umb_data
                .rename(columns={'Cohort':'Study', 'Study Status':'Status', 'Date of ICF':'Date of Consent', 'Baseline Visit Date':'Baseline Date'})
                .join(umb_vaccines)
                .join(cohorts['ROBIN'], rsuffix='_robin')
                .join(cohorts['SHIELD'], rsuffix='_shield')
    )
    studies = {
        'APOLLO': apollo_data,
        'PARIS': paris_data,
        'Healthy Donors': hd_data,
        'Umbrella': umb_data,
    }
    return studies, cohorts

def make_dems_sheet(studies):
    all_participants = pd.concat(studies.values())
    dem_cols = ['Participant ID', 'Name', 'Email', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'DOB', 'Age', 'Sex', 'Gender', 'Race', 'Ethnicity']
    return all_participants.reset_index().loc[:, dem_cols].copy()

def make_covid_sheet(studies, cohorts):
    shared_cols = ['Participant ID', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                'Sex', 'Race', 'Ethnicity',
                'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type',
                'Infection 1 Date', 'Infection 2 Date']
    paris_cols = shared_cols + ['Dose #6 Date', 'Dose #6 Type', 'Dose #7 Date', 'Dose #7 Type',
                                'Dose #8 Date', 'Dose #8 Type', 'Infection 3 Date', 'Infection 4 Date']
    paris_output = studies['PARIS'].reset_index().loc[:, paris_cols]
    crp_cols= shared_cols + ['Infection 3 Date']
    crp_output = cohorts['CRP'].reset_index().loc[:, crp_cols]
    apollo_cols= shared_cols + ['Gender', 'Total Vaccine Doses', 'Dose #6 Date', 'Dose #6 Type',
                                'Dose #7 Date', 'Dose #7 Type', 'Dose #8 Date', 'Dose #8 Type', 'Dose #9 Date', 'Dose #9 Type',
                                'Dose #10 Date', 'Dose #10 Type', 'Total Infections', 'Infection 3 Date', 'Infection 4 Date']
    apollo_output = studies['APOLLO'].reset_index().loc[:, apollo_cols]
    titan_cols= shared_cols + ['Dose #6 Date', 'Dose #6 Type']
    titan_output = cohorts['TITAN'].reset_index().loc[:, titan_cols]
    iris_cols = shared_cols + ['Dose #6 Date', 'Dose #6 Type', 'Dose #7 Date', 'Dose #7 Type',
                'Dose #8 Date', 'Dose #8 Type', 'Dose #9 Date', 'Dose #9 Type', 'Infection 3 Date']
    iris_output = cohorts['IRIS'].reset_index().loc[:, iris_cols]
    mars_cols = shared_cols + ['Dose #6 Date', 'Dose #6 Type', 'Dose #7 Date', 'Dose #7 Type',
                               'Dose #8 Date', 'Dose #8 Type', 'Dose #9 Date', 'Dose #9 Type',
                               'Infection 3 Date', 'Infection 4 Date']
    mars_output = cohorts['MARS'].reset_index().loc[:, mars_cols]
    gaea_cols = shared_cols + ['Dose #6 Date', 'Dose #6 Type', 'Dose #7 Date', 'Dose #7 Type',
                'Dose #8 Date', 'Dose #8 Type', 'Dose #9 Date', 'Dose #9 Type',
                'Infection 3 Date']
    gaea_output = cohorts['GAEA'].reset_index().loc[:, gaea_cols]
    covid_dfs=[apollo_output, paris_output, crp_output, gaea_output, mars_output, iris_output, titan_output]
    covid_data = pd.concat(covid_dfs)
    immune_event_cols = ['Dose #1 Date', 'Dose #2 Date', 'Dose #3 Date', 'Dose #4 Date', 'Dose #5 Date', 
                                'Dose #6 Date', 'Dose #7 Date', 'Dose #8 Date', 'Dose #9 Date', 'Dose #10 Date',
                                'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    vaccine_date_cols = ['Dose #1 Date', 'Dose #2 Date', 'Dose #3 Date', 'Dose #4 Date', 'Dose #5 Date', 
                                'Dose #6 Date', 'Dose #7 Date', 'Dose #8 Date', 'Dose #9 Date', 'Dose #10 Date']
    vaccine_type_cols = ['Dose #1 Type', 'Dose #2 Type', 'Dose #3 Type', 'Dose #4 Type', 'Dose #5 Type', 
                                'Dose #6 Type', 'Dose #7 Type', 'Dose #8 Type', 'Dose #9 Type', 'Dose #10 Type']
    infection_date_cols = ['Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    covid_data['Date of Last Immune Event'] = covid_data.apply(lambda row: permissive_datemax([row[immune_event] for immune_event in immune_event_cols], pd.Timestamp.today()), axis=1)
    covid_data['Immune History'] = covid_data.apply(lambda row: immune_history(row[vaccine_date_cols], row[vaccine_type_cols], row[infection_date_cols], pd.Timestamp.today()), axis=1)
    covid_data['Total Immune Events'] = covid_data['Total Vaccine Doses'].fillna(0).astype(int) + covid_data['Total Infections'].fillna(0).astype(int) #TODO: Update to be correct for non-APOLLO
    covid_data['Last Immune Event'] = covid_data['Immune History'].str.endswith('I').apply(lambda val: 'Infection' if val else 'Vaccination')
    return covid_data.drop_duplicates(subset='Participant ID', keep='last')

def make_lookup_sheet():
    lookup = {'Participant ID': ['']*100}
    lookup_df = pd.DataFrame(lookup)
    for col in ['Name', 'Email', 'Study', 'Status', 'Date of Consent',
                'Baseline Date', 'DOB', 'Age', 'Sex', 'Gender', 'Race', 'Ethnicity']:
        lookup_df[col] = ''
    dem_cols = {
        'Email': 'C:C',
        'Name': 'B:B',
        'Study': 'D:D',
        'Status': 'E:E',
        'Date of Consent': 'F:F',
        'Baseline Date': 'G:G',
        'DOB': 'H:H',
        'Age': 'I:I',
        'Sex': 'J:J',
        'Gender': 'K:K',
        'Race': 'L:L',
        'Ethnicity': 'M:M',
    }
    for dem_name, dem_col in dem_cols.items():
        lookup_df[dem_name] = [f"=IF($A{idx+2}>0, XLOOKUP($A{idx+2}" for idx in lookup_df.index]
        lookup_df[dem_name] += f",'All Demographics'!$A:$A,'All Demographics'!{dem_col}), \"\")"
    return lookup_df

def aggregate_sample_data():
     with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*")
        cam_data=cam_convert.transform_cam()
        cam_data.rename(columns={'sample_id':'Sample ID'}, inplace=True)
        cam_data.set_index('Sample ID')
        intake = query_intake(include_research=True).reset_index()
        intake.rename(columns={'sample_id':'Sample ID'}, inplace=True)
        intake.set_index('Sample ID')
        proc_data = query_dscf().reset_index()
        proc_data.rename(columns={'sample_id':'Sample ID'}, inplace=True)
        proc_data.set_index('Sample ID')
        sample_data=(intake.join(cam_data, rsuffix='_cam')
                        .join(proc_data, rsuffix='_proc'))
        sample_cols = []
        return sample_data

def make_sheet(col_list):
    all_studies = pd.concat(studies.values())
    all_cohorts = pd.concat(cohorts.values())
    all_data = all_studies.join(all_cohorts, on='Participant ID', rsuffix='_cohort')
    sheet = all_data.reset_index().loc[:, col_list]
    return pd.DataFrame(sheet)

if __name__ == '__main__':
    studies, cohorts = pull_trackers()
    dems = make_dems_sheet(studies)
    covid_data = make_covid_sheet(studies, cohorts)
    lookup_df = make_lookup_sheet()
    sample_data = aggregate_sample_data()

    output_filename = util.cross_project + 'PVI Dashboard {}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    with pd.ExcelWriter(output_filename) as writer:
        dems.to_excel(writer, index = False, sheet_name='All Demographics', freeze_panes=(1,1))
        lookup_df.to_excel(writer, index=False, sheet_name='Demographic Lookup', freeze_panes=(1,1))
        covid_data.to_excel(writer, index = False, sheet_name='SARS-CoV-2 Immune Histories', freeze_panes=(1,1))
        sample_data.to_excel(writer, sheet_name='Sample Data', freeze_panes=(1,1))
        studies['APOLLO'].to_excel(writer, sheet_name='All APOLLO', freeze_panes=(1,1))
        studies['PARIS'].to_excel(writer, sheet_name='All PARIS', freeze_panes=(1,1))
        studies['Healthy Donors'].to_excel(writer, sheet_name='All Healthy Donors', freeze_panes=(1,1))
        studies['Umbrella'].to_excel(writer, sheet_name='All Umbrella', freeze_panes=(1,1))
        print("Written to {}".format(output_filename))
