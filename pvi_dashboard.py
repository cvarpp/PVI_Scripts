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
    apollo_immune_event_cols = ['Dose #1 Date', 'Dose #2 Date', 'Dose #3 Date', 'Dose #4 Date', 'Dose #5 Date', 
                     'Dose #6 Date', 'Dose #7 Date', 'Dose #8 Date', 'Dose #9 Date', 
                     'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    apollo_vaccine_date_cols = ['Dose #1 Date', 'Dose #2 Date', 'Dose #3 Date', 'Dose #4 Date', 'Dose #5 Date', 
                     'Dose #6 Date', 'Dose #7 Date', 'Dose #8 Date', 'Dose #9 Date']
    apollo_vaccine_type_cols = ['Dose #1 Type', 'Dose #2 Type', 'Dose #3 Type', 'Dose #4 Type', 'Dose #5 Type', 
                     'Dose #6 Type', 'Dose #7 Type', 'Dose #8 Type', 'Dose #9 Type']
    apollo_infection_date_cols = ['Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    apollo_data['Date of Last Immune Event'] = apollo_data.apply(lambda row: permissive_datemax([row[immune_event] for immune_event in apollo_immune_event_cols], pd.Timestamp.today() ), axis=1)
    apollo_data['Immune History'] = apollo_data.apply(lambda row: immune_history(row[apollo_vaccine_date_cols], row[apollo_vaccine_type_cols], row[apollo_infection_date_cols], pd.Timestamp.today() ), axis=1)
    apollo_data['Total Immune Events'] = apollo_data['Total Vaccine Doses'].astype(int) + apollo_data['Total Infections'].astype(int)
    apollo_data['Last Immune Event'] = apollo_data['Immune History'].str.endswith('I').apply(lambda val: 'Infection' if val else 'Vaccination')

    apollo_cols= ['Participant ID', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                  'Sex', 'Gender', 'Race', 'Ethnicity', 'Immune History', 'Total Immune Events', 'Total Vaccine Doses', 
                  'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                  'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type', 'Dose #6 Date', 'Dose #6 Type', 
                  'Dose #7 Date', 'Dose #7 Type', 'Dose #8 Date', 'Dose #8 Type', 'Dose #9 Date', 'Dose #9 Type', 
                  'Total Infections', 'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
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
    paris_drop_cols = ['Unnamed: 0', 'Rought Last Name']
    paris_immune_event_cols = ['Dose #1 Date', 'Dose #2 Date', 'Dose #3 Date', 'Dose #4 Date', 'Dose #5 Date', 
                     'Dose #6 Date', 'Dose #7 Date', 'Dose #8 Date', 
                     'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    paris_vaccine_date_cols = ['Dose #1 Date', 'Dose #2 Date', 'Dose #3 Date', 'Dose #4 Date', 'Dose #5 Date', 
                     'Dose #6 Date', 'Dose #7 Date', 'Dose #8 Date']
    paris_vaccine_type_cols = ['Dose #1 Type', 'Dose #2 Type', 'Dose #3 Type', 'Dose #4 Type', 'Dose #5 Type', 
                     'Dose #6 Type', 'Dose #7 Type', 'Dose #8 Type']
    paris_infection_date_cols = ['Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    paris_data['Date of Last Immune Event'] = paris_data.apply(lambda row: permissive_datemax([row[immune_event] for immune_event in paris_immune_event_cols], pd.Timestamp.today()), axis=1)
    paris_data['Immune History'] = paris_data.apply(lambda row: immune_history(row[paris_vaccine_date_cols], row[paris_vaccine_type_cols], row[paris_infection_date_cols], pd.Timestamp.today()), axis=1)
    paris_data['Total Infections'] = paris_data['Immune History'].str.count('I')
    paris_data['Total Immune Events'] = paris_data['Immune History'].str.count('-') + 1
    paris_data['Total Vaccine Doses'] = paris_data['Total Immune Events'].astype(int) - paris_data['Total Infections'].astype(int)
    paris_data['Last Immune Event'] = paris_data['Immune History'].str.endswith('I').apply(lambda val: 'Infection' if val else 'Vaccination')

    paris_cols= ['Participant ID', 'Status', 'Date of Consent', 'Baseline Date', 'Baseline Sample ID', 'DOB', 'Age', 
                  'Sex', 'Race', 'Ethnicity', 'Immune History', 'Total Immune Events', 'Total Vaccine Doses', 
                  'Dose #1 Date', 'Dose #1 Type', 'Dose #2 Date', 'Dose #2 Type', 'Dose #3 Date', 'Dose #3 Type', 
                  'Dose #4 Date', 'Dose #4 Type', 'Dose #5 Date', 'Dose #5 Type', 'Dose #6 Date', 'Dose #6 Type', 
                  'Dose #7 Date', 'Dose #7 Type', 'Dose #8 Date', 'Dose #8 Type', 
                  'Total Infections', 'Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
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
    umb_data = umb_data.join(crp_data, rsuffix='_crp').copy()
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
    umb_data = umb_data.join(gaea_data, rsuffix='_gaea').copy()
    #MARS
    mars_data=pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID'])
    mars_infections=pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name='COVID_Infections').dropna(subset=['Participant ID'])
    mars_data['Participant ID'] = mars_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, mars_data['Participant ID'].unique())
    mars_data.set_index('Participant ID', inplace=True)
    mars_infections.set_index('Participant ID', inplace=True)
    umb_data = umb_data.join(mars_data, rsuffix='_mars').copy()
    #IRIS
    iris_data = pd.read_excel(util.iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID'])
    iris_data['Participant ID'] = iris_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, iris_data['Participant ID'].unique())
    iris_data.set_index('Participant ID', inplace=True)
    umb_data = umb_data.join(iris_data, rsuffix='_iris').copy()
    #TITAN
    titan_data = pd.read_excel(util.titan_tracker, sheet_name='Tracker',header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_data.rename(columns={'Umbrella Corresponding Participant ID':'Participant ID'}, inplace=True)
    titan_data['Participant ID'] = titan_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants=np.append(participants, titan_data['Participant ID'].unique())
    titan_data.set_index('Participant ID', inplace=True)
    umb_data = umb_data.join(titan_data, rsuffix='_titan').copy()

    all_studies = [apollo_data, paris_data, hd_data, umb_data]
    all_participants = pd.concat(all_studies)

    date_cols = ['Date of Consent', 'Baseline Date', 'DOB']

    dem_cols = ['Participant ID', 'Name', 'Email', 'Study', 'Status', 'Date of Consent', 'Baseline Date', 'DOB', 'Age', 'Sex', 'Gender', 'Race', 'Ethnicity']
    dems = all_participants.reset_index().loc[:, dem_cols]

    covid_dfs=[apollo_output,paris_output]
    all_covid = pd.concat(covid_dfs)

    output_filename = util.pvi + 'Secret Sheets/Jessica/PVI Dashboard Testing/PVI Sampling Dashboard_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    with pd.ExcelWriter(output_filename) as writer:
        dems.to_excel(writer, index = False, sheet_name='All Demographics', freeze_panes=(1,1))
        all_covid.to_excel(writer, index = False, sheet_name='All SARS-CoV-2', freeze_panes=(1,1))
        apollo_data.to_excel(writer, sheet_name='All APOLLO', freeze_panes=(1,1))
        paris_data.to_excel(writer, sheet_name='All PARIS', freeze_panes=(1,1))
        hd_data.to_excel(writer, sheet_name='All Healthy Donors', freeze_panes=(1,1))
        umb_data.to_excel(writer, sheet_name='All Umbrella', freeze_panes=(1,1))
        print("Written to {}".format(output_filename))

if __name__ == '__main__':
    pass