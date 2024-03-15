import pandas as pd
import datetime
import util

if __name__ == '__main__':
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    gaea_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/GAEA/'
    gaea_data = pd.read_excel(gaea_folder + 'GAEA tracker.xlsx', sheet_name='Summary').dropna(subset=['Participant ID']).set_index('Participant ID')
    priority_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/PRIORITY/'
    prio_data_vax = pd.read_excel(priority_folder + 'Priority tracker.xlsx', sheet_name='Vaccinations').dropna(subset=['Record ID'])
    prio_data_main = pd.read_excel(priority_folder + 'Priority tracker.xlsx', sheet_name='Participants tracker').dropna(subset=['Record ID']).set_index('Record ID')
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
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    vaccine_stuff = {}
    columns = ['Participant ID', 'Timepoint', 'Vaccine Type', 'Vaccine Date']
    exclusions = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name='Exclusions')
    exclude_ppl = set(exclusions['Participant ID'].unique())
    """
    IRIS Stuff
    """
    for col in columns:
        vaccine_stuff[col] = []
    for participant in iris_data.index.to_numpy():
        if participant in exclude_ppl:
            continue
        try:
            index_date = iris_data.loc[participant, 'Second Dose Date'].date()
        except:
            print("Participant {} ({}) is missing vaccine dose 2 date (dose 1 on {})".format(participant, 'IRIS', iris_data.loc[participant, 'First Dose Date']))
            index_date = 'N/A'
        vaccine = iris_data.loc[participant, 'Which Vaccine?']
        if str(vaccine).upper()[:1] == 'J' and str(vaccine).upper()[:2] != 'JA':
            vaccine = 'Johnson & Johnson'
        date_1 = iris_data.loc[participant, 'First Dose Date']
        if type(date_1) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if type(vaccine) == str and vaccine[:2].upper() == 'JO':
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
            if type(vaccine) == str and vaccine[:2].upper() == 'JO':
                vaccine_stuff['Timepoint'].append('Dose 2')
            else:
                vaccine_stuff['Timepoint'].append('Dose 3')
            vaccine_stuff['Vaccine Type'].append(boost)
            vaccine_stuff['Vaccine Date'].append(date_3)
        boost_cols = [['Fourth Dose Date', 'Fourth Dose Type'], ['Fifth Dose Date', 'Fifth Dose Type']]
        timepoint_cols = ['Booster 1', 'Booster 2']
        for i, vals in enumerate(boost_cols):
            dt, tp = vals
            date_boost = iris_data.loc[participant, dt]
            boost_type = iris_data.loc[participant, tp]
            if 'bivalent' in str(boost_type).lower():
                boost_type = boost_type.split()[0]
                addendum = ':Bivalent'
            else:
                addendum = ''
            if type(date_boost) in [datetime.datetime, pd.Timestamp]:
                vaccine_stuff['Participant ID'].append(participant)
                vaccine_stuff['Timepoint'].append(timepoint_cols[i] + addendum)
                vaccine_stuff['Vaccine Type'].append(boost_type)
                vaccine_stuff['Vaccine Date'].append(date_boost)
    pd.DataFrame(vaccine_stuff).to_excel(util.seronet_vax + 'iris_vaccines.xlsx', index=False)
    """
    TITAN Stuff
    """
    for col in columns:
        vaccine_stuff[col] = []
    for participant in titan_data.index.to_numpy():
        if participant in exclude_ppl:
            continue
        try:
            index_date = titan_data.loc[participant, '3rd Dose Vaccine Date'].date()
        except:
            print("Participant {}  ({}) is missing vaccine dose 3 date (dose 1 on {}, dose 2 on {})".format(participant, 'TITAN', titan_data.loc[participant, 'Vaccine #1 Date'], titan_data.loc[participant, 'Vaccine #2 Date']))
            index_date = 'N/A'
        vaccine = titan_data.loc[participant, 'Vaccine Type']
        if str(vaccine).upper()[:1] == 'J' and str(vaccine).upper()[:2] != 'JA':
            vaccine = 'Johnson & Johnson'
        date_1 = titan_data.loc[participant, 'Vaccine #1 Date']
        if type(date_1) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if type(vaccine) == str and vaccine[:2].upper() == 'JO':
                vaccine_stuff['Timepoint'].append('Dose 1 of 1')
                vaccine_stuff['Vaccine Type'].append('Johnson & Johnson')
            else:
                vaccine_stuff['Timepoint'].append('Dose 1 of 2')
                vaccine_stuff['Vaccine Type'].append(str(vaccine).capitalize().strip())
            vaccine_stuff['Vaccine Date'].append(date_1)
        date_2 = titan_data.loc[participant, 'Vaccine #2 Date']
        if type(date_2) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            vaccine_stuff['Timepoint'].append('Dose 2 of 2')
            vaccine_stuff['Vaccine Type'].append(str(vaccine).capitalize().strip())
            vaccine_stuff['Vaccine Date'].append(date_2)
        date_3 = titan_data.loc[participant, '3rd Dose Vaccine Date']
        boost = titan_data.loc[participant, '3rd Dose Vaccine Type']
        if type(date_3) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if 'bivalent' in str(boost).lower():
                boost = boost.split()[0]
                addendum = ':Bivalent'
            elif 'monovalent' in str(boost).lower():
                boost = boost.split()[0]
                addendum = ':Monovalent XBB.1.5'
            else:
                addendum = ''
            if type(vaccine) == str and vaccine[:2].upper() == 'JO':
                vaccine_stuff['Timepoint'].append('Dose 2' + addendum)
            else:
                vaccine_stuff['Timepoint'].append('Dose 3' + addendum)
            vaccine_stuff['Vaccine Type'].append(boost)
            vaccine_stuff['Vaccine Date'].append(date_3)
        boost_cols = [['First Booster Dose Date (#4)', 'First Booster Vaccine Type (#4)'], ['Second Booster Dose Date (#5)', 'Second Booster Vaccine Type (#5)'], ['Third Booster\nDose Date (#6)', 'Third Booster\nVaccine Type (#6)']]
        timepoint_cols = ['Booster 1', 'Booster 2', 'Booster 3']
        for i, vals in enumerate(boost_cols):
            dt, tp = vals
            date_boost = titan_data.loc[participant, dt]
            boost_type = titan_data.loc[participant, tp]
            if 'bivalent' in str(boost_type).lower():
                boost_type = boost_type.split()[0]
                addendum = ':Bivalent'
            elif 'monovalent' in str(boost_type).lower():
                boost_type = boost_type.split()[0]
                addendum = ':Monovalent XBB.1.5'
            else:
                addendum = ''
            if type(date_boost) in [datetime.datetime, pd.Timestamp]:
                vaccine_stuff['Participant ID'].append(participant)
                vaccine_stuff['Timepoint'].append(timepoint_cols[i] + addendum)
                vaccine_stuff['Vaccine Type'].append(str(boost_type).capitalize().strip())
                vaccine_stuff['Vaccine Date'].append(date_boost)
    from_tracker = pd.DataFrame(vaccine_stuff).rename(columns={'Timepoint': 'Vaccination_Status'}).set_index(['Participant ID', 'Vaccination_Status'])
    from_long = pd.read_excel(titan_folder + 'TITAN for D4 Long.xlsx', sheet_name='COVID Vaccinations').set_index(['Participant ID', 'Vaccination_Status'])
    in_both = from_long.join(from_tracker, how='inner', lsuffix='_long', rsuffix='_tracker')
    in_both_correct = in_both[(in_both['SARS-CoV-2_Vaccination_Date'] == in_both['Vaccine Date']) &
                                (in_both['SARS-CoV-2_Vaccine_Type'] == in_both['Vaccine Type'])]
    date_issue = in_both[(in_both['SARS-CoV-2_Vaccination_Date'] != in_both['Vaccine Date'])]
    type_issue = in_both[(in_both['SARS-CoV-2_Vaccine_Type'] != in_both['Vaccine Type'])]
    add_to_tracker = from_long[~from_long.index.isin(from_tracker.index)]
    add_to_long = from_tracker[~from_tracker.index.isin(from_long.index)]
    with pd.ExcelWriter(util.seronet_vax + 'titan_vaccines.xlsx') as writer:
        add_to_long.reset_index().to_excel(writer, sheet_name='Add to Long', index=False)
        date_issue.reset_index().to_excel(writer, sheet_name='Date Discrepancy', index=False)
        type_issue.reset_index().to_excel(writer, sheet_name='Type Discrepancy', index=False)
        add_to_tracker.reset_index().to_excel(writer, sheet_name='Add to Tracker', index=False)
        in_both_correct.reset_index().to_excel(writer, sheet_name='In Both (agreed)', index=False)
        from_tracker.reset_index().to_excel(writer, sheet_name='Tracker Source', index=False)
        from_long.reset_index().to_excel(writer, sheet_name='Long D4 Source', index=False)
    """
    MARS Stuff
    """
    for col in columns:
        vaccine_stuff[col] = []
    for participant in mars_data.index.to_numpy():
        if participant in exclude_ppl:
            continue
        try:
            index_date = mars_data.loc[participant, 'Vaccine #2 Date'].date()
        except:
            print("Participant {} ({}) is missing vaccine dose 2 date (dose 1 on {})".format(participant, 'MARS', mars_data.loc[participant, 'Vaccine #1 Date']))
            index_date = 'N/A'
        vaccine = mars_data.loc[participant, 'Vaccine Name']
        if str(vaccine).upper()[:1] == 'J' and str(vaccine).upper()[:2] != 'JA':
            vaccine = 'Johnson & Johnson'
        date_1 = mars_data.loc[participant, 'Vaccine #1 Date']
        if type(date_1) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if type(vaccine) == str and vaccine[:2].upper() == 'JO':
                vaccine_stuff['Timepoint'].append('Dose 1 of 1')
                vaccine_stuff['Vaccine Type'].append('Johnson & Johnson')
            else:
                vaccine_stuff['Timepoint'].append('Dose 1 of 2')
                vaccine_stuff['Vaccine Type'].append(str(vaccine).capitalize().strip())
            vaccine_stuff['Vaccine Date'].append(date_1)
        date_2 = mars_data.loc[participant, 'Vaccine #2 Date']
        if type(date_2) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            vaccine_stuff['Timepoint'].append('Dose 2 of 2')
            vaccine_stuff['Vaccine Type'].append(str(vaccine).capitalize().strip())
            vaccine_stuff['Vaccine Date'].append(date_2)
        boost_cols = [['3rd Vaccine', '3rd Vaccine Type '], ['4th vaccine', '4th Vaccine Type'], ['5th vaccine', '5th Vaccine Type'],
                      ['6th vaccine', '6th vaccine type'], ['7th vaccine', '7th vaccine type'], ['8th vaccine', '8th vaccine type']]
        timepoint_cols = ['Dose 3', 'Booster 1', 'Booster 2', 'Booster 3', 'Booster 4', 'Booster 5']
        for i, vals in enumerate(boost_cols):
            dt, tp = vals
            date_boost = mars_data.loc[participant, dt]
            boost_type = mars_data.loc[participant, tp]
            if 'bivalent' in str(boost_type).lower():
                boost_type = boost_type.split()[0]
                addendum = ':Bivalent'
            elif 'XBB' in str(boost_type).upper():
                boost_type = boost_type.split()[0]
                addendum = ':Monovalent XBB.1.5'
            else:
                addendum = ''
            if type(date_boost) in [datetime.datetime, pd.Timestamp]:
                if type(boost_type) == str and boost_type[:2].upper() == 'JO':
                    vaccine_stuff['Vaccine Type'].append('Johnson & Johnson')
                else:
                    vaccine_stuff['Vaccine Type'].append(str(boost_type).capitalize().strip())
                vaccine_stuff['Participant ID'].append(participant)
                if type(vaccine) == str and vaccine[:2].upper() == 'JO' and timepoint_cols[i] == 'Dose 3':
                    vaccine_stuff['Timepoint'].append('Dose 2' + addendum)
                else:
                    vaccine_stuff['Timepoint'].append(timepoint_cols[i] + addendum)
                vaccine_stuff['Vaccine Date'].append(date_boost)
    # pd.DataFrame(vaccine_stuff).to_excel(util.seronet_vax + 'mars_vaccines.xlsx', index=False)
    from_tracker = pd.DataFrame(vaccine_stuff).rename(columns={'Timepoint': 'Vaccination_Status'}).set_index(['Participant ID', 'Vaccination_Status'])
    from_long = pd.read_excel(mars_folder + 'MARS for D4 Long.xlsx', sheet_name='COVID Vaccinations').set_index(['Participant ID', 'Vaccination_Status'])
    in_both = from_long.join(from_tracker, how='inner', lsuffix='_long', rsuffix='_tracker')
    in_both_correct = in_both[(in_both['SARS-CoV-2_Vaccination_Date'] == in_both['Vaccine Date']) &
                                (in_both['SARS-CoV-2_Vaccine_Type'] == in_both['Vaccine Type'])]
    date_issue = in_both[(in_both['SARS-CoV-2_Vaccination_Date'] != in_both['Vaccine Date'])]
    type_issue = in_both[(in_both['SARS-CoV-2_Vaccine_Type'] != in_both['Vaccine Type'])]
    add_to_tracker = from_long[~from_long.index.isin(from_tracker.index)]
    add_to_long = from_tracker[~from_tracker.index.isin(from_long.index)]
    with pd.ExcelWriter(util.seronet_vax + 'mars_vaccines.xlsx') as writer:
        add_to_long.reset_index().to_excel(writer, sheet_name='Add to Long', index=False)
        date_issue.reset_index().to_excel(writer, sheet_name='Date Discrepancy', index=False)
        type_issue.reset_index().to_excel(writer, sheet_name='Type Discrepancy', index=False)
        add_to_tracker.reset_index().to_excel(writer, sheet_name='Add to Tracker', index=False)
        in_both_correct.reset_index().to_excel(writer, sheet_name='In Both (agreed)', index=False)
        from_tracker.reset_index().to_excel(writer, sheet_name='Tracker Source', index=False)
        from_long.reset_index().to_excel(writer, sheet_name='Long D4 Source', index=False)
    """
    PRIORITY Stuff
    """
    for col in columns:
        vaccine_stuff[col] = []
    for participant in prio_data.index.to_numpy():
        if participant in exclude_ppl:
            continue
        try:
            index_date = prio_data.loc[participant, 'Baseline date'].date()
        except:
            print("Participant {} ({}) is missing baseline date".format(participant, 'PRIORITY'))
            index_date = 'N/A'
        vaccine = prio_data.loc[participant, 'Vaccine Type']
        if str(vaccine).upper()[:1] == 'J' and str(vaccine).upper()[:2] != 'JA':
            vaccine = 'Johnson & Johnson'
        date_1 = prio_data.loc[participant, 'Vaccine Dose 1']
        if type(date_1) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if type(vaccine) == str and vaccine[:2].upper() == 'JO':
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
        boost_cols = [['Boost Date', 'Boost Type'], ['Vaccine Dose 4', 'Vaccine Type Dose 4'], ['Vaccine Dose 5', 'Vaccine Type Dose 5']]
        timepoint_cols = ['Booster 1', 'Booster 2', 'Booster 3']
        for i, vals in enumerate(boost_cols):
            dt, tp = vals
            date_boost = prio_data.loc[participant, dt]
            boost_type = prio_data.loc[participant, tp]
            if 'bivalent' in str(boost_type).lower():
                boost_type = boost_type.split()[0]
                addendum = ':Bivalent'
            else:
                addendum = ''
            if type(date_boost) in [datetime.datetime, pd.Timestamp]:
                vaccine_stuff['Participant ID'].append(participant)
                vaccine_stuff['Timepoint'].append(timepoint_cols[i] + addendum)
                vaccine_stuff['Vaccine Type'].append(boost_type)
                vaccine_stuff['Vaccine Date'].append(date_boost)
    pd.DataFrame(vaccine_stuff).to_excel(util.seronet_vax + 'prio_vaccines.xlsx', index=False)
    """
    GAEA Stuff
    """
    for col in columns:
        vaccine_stuff[col] = []
    for participant in gaea_data.index.to_numpy():
        if participant in exclude_ppl:
            continue
        try:
            index_date = gaea_data.loc[participant, 'Baseline Date'].date()
        except:
            print("Participant {} ({}) is missing baseline date".format(participant, 'GAEA'))
            index_date = 'N/A'
        vaccine = gaea_data.loc[participant, 'Vaccine Type']
        if str(vaccine).upper()[:1] == 'J' and str(vaccine).upper()[:2] != 'JA':
            vaccine = 'Johnson & Johnson'
        date_1 = gaea_data.loc[participant, 'Dose #1 Date']
        if type(date_1) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            if type(vaccine) == str and vaccine[:2].upper() == 'JO':
                vaccine_stuff['Timepoint'].append('Dose 1 of 1')
            else:
                vaccine_stuff['Timepoint'].append('Dose 1 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_1)
        date_2 = gaea_data.loc[participant, 'Dose #2 Date']
        if type(date_2) in [datetime.datetime, pd.Timestamp]:
            vaccine_stuff['Participant ID'].append(participant)
            vaccine_stuff['Timepoint'].append('Dose 2 of 2')
            vaccine_stuff['Vaccine Type'].append(vaccine)
            vaccine_stuff['Vaccine Date'].append(date_2)
        boost_cols = [['3rd Vaccine Date', '3rd Vaccine Type '], ['4th Vaccine Date', '4th Vaccine Type'], ['5th Vaccine Date', '5th Vaccine Type']]
        timepoint_cols = ['Booster 1', 'Booster 2', 'Booster 3']
        for i, vals in enumerate(boost_cols):
            dt, tp = vals
            date_boost = gaea_data.loc[participant, dt]
            boost_type = gaea_data.loc[participant, tp]
            if 'bivalent' in str(boost_type).lower():
                boost_type = boost_type.split()[0]
                addendum = ':Bivalent'
            else:
                addendum = ''
            if type(date_boost) in [datetime.datetime, pd.Timestamp]:
                vaccine_stuff['Participant ID'].append(participant)
                vaccine_stuff['Timepoint'].append(timepoint_cols[i] + addendum)
                vaccine_stuff['Vaccine Type'].append(boost_type)
                vaccine_stuff['Vaccine Date'].append(date_boost)
    pd.DataFrame(vaccine_stuff).to_excel(util.seronet_vax + 'gaea_vaccines.xlsx', index=False)