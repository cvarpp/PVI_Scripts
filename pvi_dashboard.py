import pandas as pd
import numpy as np
from datetime import date
from dateutil import parser
import util
import argparse
import sys
import PySimpleGUI as sg
from helpers import try_datediff, permissive_datemax, query_intake, map_dates, immune_history, ValuesToClass
import warnings

with warnings.catch_warnings():
    warnings.filterwarnings("ignore", message=".*extension is not supported.*", module='openpyxl')
#Make a Massive Tracker
#APOLLO
    apollo_data = pd.read_excel(util.apollo_tracker, sheet_name='Summary').dropna(subset=['Participant ID'])
    apollo_data['Participant ID'] = apollo_data['Participant ID'].apply(lambda val: val.strip().upper())
    apollo_vaccines = pd.read_excel(util.apollo_tracker, sheet_name='Vaccinations').dropna(subset=['Participant ID'])
    apollo_infections = pd.read_excel(util.apollo_tracker, sheet_name='Infections').dropna(subset=['Participant ID'])
    participants = apollo_data['Participant ID'].unique()
    apollo_data.set_index('Participant ID', inplace=True)
    apollo_vaccines.set_index('Participant ID', inplace=True)
    apollo_infections.set_index('Participant ID', inplace=True)
    apollo_data = (apollo_data.join(apollo_vaccines, on='Participant ID', how = 'inner')
                              .join(apollo_infections, on='Participant ID', how = 'inner')
                              .copy())
    apollo_data['Study'] = 'APOLLO'
    apollo_data.rename(columns={'Age @ Enrollment':'Age', 'Number of Vaccine Doses':'Total Vaccine Doses', 
                                'Number of SARS-CoV-2 Infections':'Total Infections','Symptom Onset Date 1':'Infection 1 Date', 
                                'Symptom Onset Date 2':'Infection 2 Date', 'Symptom Onset Date 3':'Infection 3 Date', 
                                'Symptom Onset Date 4':'Infection 4 Date'}, inplace=True)

    apollo_cols= ['Participant ID', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                  'Sex', 'Gender', 'Race', 'Ethnicity', 'Total Vaccine Doses', 
                  'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                  'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type', 'Dose #6 Date', 'Dose #6 Type', 
                  'Dose #7 Date', 'Dose #7 Type', 'Dose #8 Date', 'Dose #8 Type', 'Dose #9 Date', 'Dose #9 Type', 
                  'Dose #10 Date', 'Dose #10 Type', 'Total Infections', 'Infection 1 Date', 'Infection 2 Date', 
                  'Infection 3 Date', 'Infection 4 Date']
    apollo_output = apollo_data.reset_index().loc[:, apollo_cols]
    
    #PARIS
    paris_data = pd.read_excel(util.paris + 'Patient Tracking - PARIS.xlsx', sheet_name='Subgroups', header=4).dropna(subset=['Participant ID'])
    paris_main = pd.read_excel(util.paris + 'Patient Tracking - PARIS.xlsx', sheet_name='Main', header=8).dropna(subset=['Subject ID']).set_index('Subject ID')
    paris_dems = pd.read_excel(util.projects + 'PARIS/Demographics.xlsx', sheet_name='inputs').dropna(subset=['Subject ID']).set_index('Subject ID')
    paris_data['Participant ID'] = paris_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, paris_data['Participant ID'].unique())
    paris_data.set_index('Participant ID', inplace=True)
    paris_data.rename(columns={'Study Status':'Status'}, inplace=True)
    paris_data = (paris_data.join(paris_dems, on='Participant ID', rsuffix='_dem')
                            .join(paris_main, on='Participant ID', rsuffix='_main')
                            .copy())
    paris_data.rename(columns={'E-mail_main':'Email', 'Date of Birth':'DOB', 'Gender':'Sex', 'Status':'Status_sg', 'Study Status':'Status',
                               'Sample ID_main':'Baseline Sample ID', 'First Dose Date':'Dose #1 Date', 'Second Dose Date':'Dose #2 Date', 
                               'Boost Date':'Dose #3 Date', 'Boost 2 Date':'Dose #4 Date', 'Boost 3 Date':'Dose #5 Date', 
                               'Boost 4 Date':'Dose #6 Date', 'Boost 5 Date':'Dose #7 Date', 'Boost 6 Date':'Dose #8 Date',
                               'Vaccine Type':'Dose #1 Type', 'Boost Type':'Dose #3 Type', 'Boost 2 Type':'Dose #4 Type', 'Boost 3 Type':'Dose #5 Type', 
                               'Boost 4 Type':'Dose #6 Type', 'Boost 5 Type':'Dose #7 Type', 'Boost 6 Type':'Dose #8 Type'}, inplace=True)
    paris_data['Dose #2 Type'] = paris_data['Dose #1 Type'].copy()
    paris_data['Study'] = 'PARIS'
    paris_data['Ethnicity'] = paris_data['Ethnicity: Hispanic or Latino'].apply(lambda val: 'Hispanic or Latino' if val == 'Yes' else 'Not Hispanic or Latino')

    paris_cols= ['Participant ID', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                  'Sex', 'Race', 'Ethnicity', 
                  'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                  'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type', 'Dose #6 Date', 'Dose #6 Type', 
                  'Dose #7 Date', 'Dose #7 Type', 'Dose #8 Date', 'Dose #8 Type', 
                  'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    paris_output = paris_data.reset_index().loc[:, paris_cols]

    #HD
    hd_data = pd.read_excel(util.clin_ops + 'Healthy donors/Healthy Donors Participant Tracker.xlsx', sheet_name='Participants').dropna(subset=['Participant ID'])
    hd_data['Participant ID'] = hd_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, hd_data['Participant ID'].unique())
    hd_data.set_index('Participant ID', inplace=True)
    hd_data.rename(columns={'Date of consent':'Date of Consent', 'Age at Baseline':'Age'}, inplace=True)
    hd_data['Baseline Date'] = hd_data['Date of Consent'].copy()
    hd_data['Study']='Healthy Donors'
    #DOVE
    dove_data=pd.read_excel(util.clin_ops + 'Healthy donors/DOVE/DOVE participant tracker.xlsx', sheet_name='Participant info').dropna(subset=['Participant ID'])
    dove_data['Participant ID'] = dove_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, dove_data['Participant ID'].unique())
    dove_data.set_index('Participant ID', inplace=True)
    hd_data = hd_data.join(dove_data, rsuffix='_dove').copy()

    #UMBRELLA
    umb_data=pd.read_excel(util.umbrella_tracker, sheet_name='Summary').dropna(subset=['Subject ID'])
    umb_vaccines=pd.read_excel(util.umbrella_tracker, sheet_name='COVID-19 Vaccine Type & Dates!!')
    umb_data.rename(columns={'Subject ID':'Participant ID'}, inplace=True)
    umb_data['Participant ID'] = umb_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, umb_data['Participant ID'].unique())
    umb_data.set_index('Participant ID', inplace=True)
    umb_vaccines.set_index('Participant ID', inplace=True)
    umb_data.rename(columns={'Cohort':'Study', 'Study Status':'Status', 'Date of ICF':'Date of Consent', 'Baseline Visit Date':'Baseline Date'}, inplace=True)
    #CRP
    crp_data = pd.read_excel(util.crp_folder + 'CRP Patient Tracker no circ ref 7.7.23.xlsx', sheet_name='Tracker', header=4).dropna(subset=['Participant ID'])
    crp_data['Participant ID'] = crp_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, crp_data['Participant ID'].unique())
    crp_data.set_index('Participant ID', inplace=True)
    
    crp_data.rename(columns={'COVID-19 Vaccine Type':'Dose #1 Type', 'Vaccine #1 Date':'Dose #1 Date', 'Vaccine #2 Date':'Dose #2 Date',
                             '3rd Dose Vaccine Type':'Dose #3 Type', '3rd Dose Vaccine Date':'Dose #3 Date', '4th Dose Vaccine Type':'Dose #4 Type',
                             '4th Dose Date':'Dose #4 Date', '5th Dose Vaccine Type':'Dose #5 Type', '5th Dose Date':'Dose #5 Date',
                             'Study Participation Status':'Status', 'Age at Visit':'Age'}, inplace=True)
    crp_data['Dose #2 Type'] = crp_data['Dose #1 Type'].copy()
    crp_data['Study'] = 'CRP'

    crp_cols= ['Participant ID', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                  'Sex', 'Gender', 'Race', 'Ethnicity',
                  'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                  'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type', 
                  'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date']
    crp_output = crp_data.reset_index().loc[:, crp_cols]
    #umb_data = umb_data.join(crp_data, rsuffix='_crp').copy()
    #ROBIN
    robin_data = pd.read_excel(util.umbrella + 'ROBIN/ROBIN Participant tracker.xlsx', sheet_name='Participant Info').dropna(subset=['Subject ID'])
    robin_data.rename(columns={'Subject ID':'Participant ID'}, inplace=True)
    robin_data['Participant ID'] = robin_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, robin_data['Participant ID'].unique())
    robin_data.set_index('Participant ID', inplace=True)
    umb_data = umb_data.join(robin_data, rsuffix='_robin').copy()
    #SHIELD
    shield_data = pd.read_excel(util.projects +'SHIELD/SHIELD tracker.xlsx', sheet_name='Summary').dropna(subset=['Participant ID'])
    shield_data['Participant ID'] = shield_data['Participant ID'].astype(str).apply(lambda val: val.strip().upper())
    shield_data=shield_data[shield_data['Participant ID'].str.startswith('1')]
    participants=np.append(participants, shield_data['Participant ID'].unique())
    shield_data.set_index('Participant ID', inplace=True)
    umb_data = umb_data.join(shield_data, rsuffix='_shield').copy()
    #GAEA
    gaea_data=pd.read_excel(util.gaea_folder + 'GAEA Tracker.xlsx', sheet_name='Summary').dropna(subset=['Participant ID'])
    gaea_infections=pd.read_excel(util.gaea_folder + 'GAEA Tracker.xlsx', sheet_name='Infections').dropna(subset=['Participant ID'])
    gaea_data['Participant ID'] = gaea_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, gaea_data['Participant ID'].unique())
    gaea_data.set_index('Participant ID', inplace=True)
    gaea_infections.set_index('Participant ID', inplace=True)
    gaea_data = gaea_data.join(gaea_infections, on='Participant ID', rsuffix='inf')

    gaea_data.rename(columns={'3rd Vaccine Date':'Dose #3 Date', '4th Vaccine Date':'Dose #4 Date', '5th Vaccine Date':'Dose #5 Date', 
                              '6th Vaccine Date':'Dose #6 Date', '7th Vaccine Date':'Dose #7 Date', '8th Vaccine Date':'Dose #8 Date',
                              '9th Vaccine Date':'Dose #9 Date', 'Vaccine Type':'Dose #1 Type', '3rd Vaccine Type':'Dose #3 Type', 
                              '4th Vaccine Type':'Dose #4 Type', '5th Vaccine Type':'Dose #5 Type', '6th Vaccine Type':'Dose #6 Type', 
                              '7th Vaccine Type':'Dose #7 Type', '8th Vaccine Type':'Dose #8 Type', '9th Vaccine Type':'Dose #9 Type',
                              'Symptom Onset Date 1':'Infection 1 Date', 'Symptom Onset Date 2':'Infection 2 Date', 'Symptom Onset Date 3':'Infection 3 Date',
                              'Age @ Baseline':'Age'}, inplace=True)
    gaea_data['Dose #2 Type'] = gaea_data['Dose #1 Type'].copy()
    gaea_data['Study'] = 'GAEA'

    gaea_cols= ['Participant ID', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                  'Sex', 'Race', 'Ethnicity',
                  'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                  'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type', 'Dose #6 Date', 'Dose #6 Type', 
                  'Dose #7 Date', 'Dose #7 Type', 'Dose #8 Date', 'Dose #8 Type', 'Dose #9 Date', 'Dose #9 Type', 
                  'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date']
    gaea_output = gaea_data.reset_index().loc[:, gaea_cols]
    #umb_data = umb_data.join(gaea_data, rsuffix='_gaea').copy()
    #MARS
    mars_data=pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID'])
    mars_infections=pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name='COVID Infections').dropna(subset=['Participant ID'])
    mars_data['Participant ID'] = mars_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, mars_data['Participant ID'].unique())
    mars_data.set_index('Participant ID', inplace=True)
    mars_infections.set_index('Participant ID', inplace=True)
    mars_data = mars_data.join(mars_infections, on='Participant ID', rsuffix='inf')

    mars_data.rename(columns={'Vaccine #1 Date':'Dose #1 Date', '1st Vaccine Type':'Dose #1 Type', 'Vaccine #2 Date':'Dose #2 Date', 
                              '2nd Vaccine Type':'Dose #2 Type', '3rd Vaccine':'Dose #3 Date', '3rd Vaccine Type':'Dose #3 Type',
                              '4th vaccine':'Dose #4 Date', '4th Vaccine Type':'Dose #4 Type', '5th vaccine':'Dose #5 Date', 
                              '5th Vaccine Type':'Dose #5 Type', '6th vaccine':'Dose #6 Date', '6th vaccine type':'Dose #6 Type',
                              '7th vaccine':'Dose #7 Date', '7th vaccine type':'Dose #7 Type', '8th vaccine':'Dose #8 Date', 
                              '8th vaccine type':'Dose #8 Type', '9th vaccine':'Dose #9 Date', '9th vaccine type':'Dose #9 Type',
                              'Study Status':'Status', 'T1':'Baseline Date', 'T1 Sample ID':'Baseline Sample ID', 'Age @ time of consent':'Age'}, inplace=True)
    mars_data['Study'] = 'MARS'

    mars_cols= ['Participant ID', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                  'Sex', 'Race', 'Ethnicity',
                  'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                  'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type', 'Dose #6 Date', 'Dose #6 Type', 
                  'Dose #7 Date', 'Dose #7 Type', 'Dose #8 Date', 'Dose #8 Type', 'Dose #9 Date', 'Dose #9 Type', 
                  'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    mars_output = mars_data.reset_index().loc[:, mars_cols]
    #umb_data = umb_data.join(mars_data, rsuffix='_mars').copy()
    #IRIS
    iris_data = pd.read_excel(util.iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID'])
    iris_data['Participant ID'] = iris_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, iris_data['Participant ID'].unique())
    iris_data.set_index('Participant ID', inplace=True)

    iris_data.rename(columns={'First Dose Date':'Dose #1 Date', 'Which Vaccine?':'Dose #1 Type', 'Second Dose Date':'Dose #2 Date', 
                              'Third Dose Date':'Dose #3 Date', 'Third Dose Type':'Dose #3 Type',
                              'Fourth Dose Date':'Dose #4 Date', 'Fourth Dose Type':'Dose #4 Type', 'Fifth Dose Date':'Dose #5 Date', 
                              'Fifth Dose Type':'Dose #5 Type', 'Sixth Dose Date':'Dose #6 Date', 'Sixth Dose Type':'Dose #6 Type',
                              'Seventh Dose Date':'Dose #7 Date', 'Seventh Dose Type':'Dose #7 Type', 'Eighth Dose Date':'Dose #8 Date', 
                              'Eighth Dose Type':'Dose #8 Type', 'Ninth Dose Date':'Dose #9 Date', 'Ninth Dose Type':'Dose #9 Type',
                              'Symptom Onset':'Infection 1 Date', 'Symptom Onset 2':'Infection 2 Date', 'Symptom Onset 3':'Infection 3 Date',
                              'Sample ID':'Baseline Sample ID'}, inplace=True)
    iris_data['Dose #2 Type'] = iris_data['Dose #1 Type'].copy()
    iris_data['Study'] = 'IRIS'
    iris_data['Status'] = 'PULL FROM UMBRELLA TRACKER' #do this eventually
    iris_data['Date of Consent'] = 'PULL FROM UMBRELLA TRACKER' #do this eventually
    
    iris_cols= ['Participant ID', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                  'Sex', 'Race', 'Ethnicity',
                  'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                  'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type', 'Dose #6 Date', 'Dose #6 Type', 
                  'Dose #7 Date', 'Dose #7 Type', 'Dose #8 Date', 'Dose #8 Type', 'Dose #9 Date', 'Dose #9 Type', 
                  'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date']
    iris_output = iris_data.reset_index().loc[:, iris_cols]
    #umb_data = umb_data.join(iris_data, rsuffix='_iris').copy()
    #TITAN
    titan_data = pd.read_excel(util.titan_tracker, sheet_name='Tracker',header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_data.rename(columns={'Umbrella Corresponding Participant ID':'Participant ID', 'Gender':'Sex', 'Study Participation Status':'Status',
                               'Age at Enrollment':'Age'}, inplace=True)
    titan_data['Participant ID'] = titan_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, titan_data['Participant ID'].unique())
    titan_data.set_index('Participant ID', inplace=True)
    titan_infections = pd.read_excel(util.titan_tracker, sheet_name='Infections Combined').dropna(subset=['Participant ID'])
    titan_infections.set_index('Participant ID', inplace=True)
    titan_dems = pd.read_excel(util.titan_tracker, sheet_name='Demographics & First Two Doses', header=3).dropna(subset='Umbrella Participant ID')
    titan_dems.set_index('Umbrella Participant ID', inplace=True)
    titan_data = (titan_data.join(titan_infections, on='Participant ID', rsuffix='inf')
                           .join(titan_dems, on='Participant ID', rsuffix='dem').copy())

    titan_data.rename(columns={'Vaccine #1 Date':'Dose #1 Date', 'Vaccine Type':'Dose #1 Type', 'Vaccine #2 Date':'Dose #2 Date', 
                              '3rd Dose Vaccine Date':'Dose #3 Date', '3rd Dose Vaccine Type':'Dose #3 Type',
                              'First Booster Dose Date (#4)':'Dose #4 Date', 'First Booster Vaccine Type (#4)':'Dose #4 Type', 'Second Booster Dose Date (#5)':'Dose #5 Date', 
                              'Second Booster Vaccine Type (#5)':'Dose #5 Type', 'Third Booster Dose Date (#6)':'Dose #6 Date', 'Third Booster Vaccine Type (#6)':'Dose #6 Type'}, inplace=True)
    titan_data['Dose #2 Type'] = titan_data['Dose #1 Type'].copy()
    titan_data['Study'] = 'TITAN'

    titan_cols= ['Participant ID', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                  'Sex', 'Race', 'Ethnicity',
                  'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                  'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type', 'Dose #6 Date', 'Dose #6 Type', 
                  'Infection 1 Date', 'Infection 2 Date']
    titan_output = titan_data.reset_index().loc[:, titan_cols]
    #umb_data = umb_data.join(titan_data, rsuffix='_titan').copy()

    all_studies = [apollo_data, paris_data, hd_data, umb_data]
    all_participants = pd.concat(all_studies)

    dem_cols = ['Participant ID', 'Name', 'Email', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'DOB', 'Age', 'Sex', 'Gender', 'Race', 'Ethnicity']
    dems = all_participants.reset_index().loc[:, dem_cols]

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
    covid_data['Date of Last Immune Event'] = covid_data.apply(lambda row: permissive_datemax([row[immune_event] for immune_event in immune_event_cols], pd.Timestamp.today() ), axis=1)
    covid_data['Immune History'] = covid_data.apply(lambda row: immune_history(row[vaccine_date_cols], row[vaccine_type_cols], row[infection_date_cols], pd.Timestamp.today() ), axis=1)
    covid_data['Total Immune Events'] = covid_data['Total Vaccine Doses'].fillna(0).astype(int) + covid_data['Total Infections'].fillna(0).astype(int)
    covid_data['Last Immune Event'] = covid_data['Immune History'].str.endswith('I').apply(lambda val: 'Infection' if val else 'Vaccination')

    output_filename = util.cross_project + 'PVI Sampling Dashboard_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    with pd.ExcelWriter(output_filename) as writer:
        dems.to_excel(writer, index = False, sheet_name='All Demographics', freeze_panes=(1,1))
        covid_data.to_excel(writer, index = False, sheet_name='SARS-CoV-2 Immune Histories', freeze_panes=(1,1))
        apollo_data.to_excel(writer, sheet_name='All APOLLO', freeze_panes=(1,1))
        paris_data.to_excel(writer, sheet_name='All PARIS', freeze_panes=(1,1))
        hd_data.to_excel(writer, sheet_name='All Healthy Donors', freeze_panes=(1,1))
        umb_data.to_excel(writer, sheet_name='All Umbrella', freeze_panes=(1,1))
        print("Written to {}".format(output_filename))

if __name__ == '__main__':
    pass