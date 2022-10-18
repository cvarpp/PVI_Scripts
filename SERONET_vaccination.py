import pandas as pd
import datetime
import util

if __name__ == '__main__':
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    priority_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/PRIORITY/'
    prio_data_vax = pd.read_excel(priority_folder + 'Priority tracker.xlsx', sheet_name='Vaccinations').dropna(subset=['Record ID'])
    prio_data_main = pd.read_excel(priority_folder + 'Priority tracker.xlsx', sheet_name='Participants tracker').dropna(subset=['Record ID']).set_index('Record ID')
    participants = prio_data_vax['Record ID'].to_numpy()
    prio_data = prio_data_vax.join(prio_data_main, on='Record ID', rsuffix='_ignore').set_index('Record ID')
    mars_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/MARS/'
    mars_data = pd.read_excel(mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID'])
    mars_data['Participant ID'] = mars_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = mars_data['Participant ID'].unique()
    mars_data.set_index('Participant ID', inplace=True)
    titan_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/TITAN/'
    titan_data = pd.read_excel(titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_data['Participant ID'] = titan_data['Umbrella Corresponding Participant ID'].apply(lambda val: val.strip().upper())
    participants = titan_data['Participant ID'].unique()
    titan_data.set_index('Participant ID', inplace=True)
    iris_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/IRIS/'
    iris_data = pd.read_excel(iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID'])
    iris_data['Participant ID'] = iris_data['Participant ID'].apply(lambda val: val.strip().upper())
    iris_data.set_index('Participant ID', inplace=True)
    samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Sample Tracking/Sample Intake Log.xlsx', sheet_name='Sample Intake Log', header=util.header_intake, dtype=str)
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    samplesClean = samples.dropna(subset=['Participant ID'])
    vaccine_stuff = {}
    columns = ['Participant ID', 'Timepoint', 'Vaccine Type', 'Vaccine Date']
    """
    IRIS Stuff
    """
    for col in columns:
        vaccine_stuff[col] = []
    for participant in iris_data.index.to_numpy():
        try:
            index_date = iris_data.loc[participant, 'Second Dose Date'].date()
        except:
            print("Participant {} ({}) is missing vaccine dose 2 date (dose 1 on {})".format(participant, 'IRIS', iris_data.loc[participant, 'First Dose Date']))
            index_date = 'N/A'
        vaccine = iris_data.loc[participant, 'Which Vaccine?']
        date_1 = iris_data.loc[participant, 'First Dose Date']
        if type(date_1) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if type(vaccine) == str and vaccine[:1].upper() == 'J':
                vaccine_stuff['Timepoint'].append('Dose 1 of 1')
            else:
                vaccine_stuff['Timepoint'].append('Dose 1 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_1)
        date_2 = iris_data.loc[participant, 'Second Dose Date']
        if type(date_2) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            vaccine_stuff['Timepoint'].append('Dose 2 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_2)
        date_3 = iris_data.loc[participant, 'Third Dose Date']
        boost = iris_data.loc[participant, 'Third Dose Type']
        if type(date_3) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            vaccine_stuff['Timepoint'].append('Dose 3')
            vaccine_stuff['Vaccine Type'].append(boost)
            vaccine_stuff['Vaccine Date'].append(date_3)
    pd.DataFrame(vaccine_stuff).to_excel(util.script_output + 'iris_vaccines.xlsx', index=False)
    """
    TITAN Stuff
    """
    for col in columns:
        vaccine_stuff[col] = []
    for participant in titan_data.index.to_numpy():
        try:
            index_date = titan_data.loc[participant, '3rd Dose Vaccine Date'].date()
        except:
            print("Participant {}  ({}) is missing vaccine dose 3 date (dose 1 on {}, dose 2 on {})".format(participant, 'TITAN', titan_data.loc[participant, 'Vaccine #1 Date'], titan_data.loc[participant, 'Vaccine #2 Date']))
            index_date = 'N/A'
        vaccine = titan_data.loc[participant, 'Vaccine Type']
        date_1 = titan_data.loc[participant, 'Vaccine #1 Date']
        if type(date_1) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if type(vaccine) == str and vaccine[:1].upper() == 'J':
                vaccine_stuff['Timepoint'].append('Dose 1 of 1')
            else:
                vaccine_stuff['Timepoint'].append('Dose 1 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_1)
        date_2 = titan_data.loc[participant, 'Vaccine #2 Date']
        if type(date_2) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            vaccine_stuff['Timepoint'].append('Dose 2 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_2)
        date_3 = titan_data.loc[participant, '3rd Dose Vaccine Date']
        boost = titan_data.loc[participant, '3rd Dose Vaccine Type']
        if type(date_3) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            vaccine_stuff['Timepoint'].append('Dose 3')
            vaccine_stuff['Vaccine Type'].append(boost)
            vaccine_stuff['Vaccine Date'].append(date_3)
    pd.DataFrame(vaccine_stuff).to_excel(util.script_output + 'titan_vaccines.xlsx', index=False)
    """
    MARS Stuff
    """
    for col in columns:
        vaccine_stuff[col] = []
    for participant in mars_data.index.to_numpy():
        try:
            index_date = mars_data.loc[participant, 'Vaccine #2 Date'].date()
        except:
            print("Participant {} ({}) is missing vaccine dose 2 date (dose 1 on {})".format(participant, 'MARS', mars_data.loc[participant, 'Vaccine #1 Date']))
            index_date = 'N/A'
        vaccine = mars_data.loc[participant, 'Vaccine Name']
        date_1 = mars_data.loc[participant, 'Vaccine #1 Date']
        if type(date_1) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if type(vaccine) == str and vaccine[:1].upper() == 'J':
                vaccine_stuff['Timepoint'].append('Dose 1 of 1')
            else:
                vaccine_stuff['Timepoint'].append('Dose 1 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_1)
        date_2 = mars_data.loc[participant, 'Vaccine #2 Date']
        if type(date_2) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            vaccine_stuff['Timepoint'].append('Dose 2 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_2)
        boost_cols = [['3rd Vaccine', '3rd Vaccine Type '], ['4th vaccine', '4th Vaccine Type'], ['5th vaccine', '5th Vaccine Type']]
        timepoint_cols = ['Dose 3', 'Booster 1', 'Booster 2']
        for i, vals in enumerate(boost_cols):
            dt, tp = vals
            date_boost = mars_data.loc[participant, dt]
            boost_type = mars_data.loc[participant, tp]
            if type(date_boost) in [datetime.datetime, pd.Timestamp]:
                vaccine_stuff['Participant ID'].append(participant)
                vaccine_stuff['Timepoint'].append(timepoint_cols[i])
                vaccine_stuff['Vaccine Type'].append(boost_type)
                vaccine_stuff['Vaccine Date'].append(date_boost)
    pd.DataFrame(vaccine_stuff).to_excel(util.script_output + 'mars_vaccines.xlsx', index=False)
    """
    PRIORITY Stuff
    """
    for col in columns:
        vaccine_stuff[col] = []
    for participant in prio_data.index.to_numpy():
        try:
            index_date = prio_data.loc[participant, 'Baseline date'].date()
        except:
            print("Participant {} ({}) is missing baseline date".format(participant, 'PRIORITY'))
            index_date = 'N/A'
        vaccine = prio_data.loc[participant, 'Vaccine Type']
        date_1 = prio_data.loc[participant, 'Vaccine Dose 1']
        if type(date_1) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if type(vaccine) == str and vaccine[:1].upper() == 'J':
                vaccine_stuff['Timepoint'].append('Dose 1 of 1')
            else:
                vaccine_stuff['Timepoint'].append('Dose 1 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_1)
        date_2 = prio_data.loc[participant, 'Vaccine Dose 2']
        if type(date_2) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            vaccine_stuff['Timepoint'].append('Dose 2 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_2)
        date_3 = prio_data.loc[participant, 'Boost Date']
        boost = prio_data.loc[participant, 'Boost Type']
        if type(date_3) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if vaccine_stuff['Timepoint'][-1] == 'Dose 1 of 1':
                vaccine_stuff['Timepoint'].append('Dose 2')
            else:
                vaccine_stuff['Timepoint'].append('Dose 3')
            vaccine_stuff['Vaccine Type'].append(boost)
            vaccine_stuff['Vaccine Date'].append(date_3)
    pd.DataFrame(vaccine_stuff).to_excel(util.script_output + 'prio_vaccines.xlsx', index=False)