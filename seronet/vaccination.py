import pandas as pd
import datetime
import util

if __name__ == '__main__':
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    gaea_data = pd.read_excel(util.gaea_folder + 'GAEA tracker.xlsx', sheet_name='Summary').dropna(subset=['Participant ID']).set_index('Participant ID')
    mars_data = pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID'])
    mars_data['Participant ID'] = mars_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = mars_data['Participant ID'].unique()
    mars_data.set_index('Participant ID', inplace=True)
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    vaccine_stuff = {}
    columns = ['Participant ID', 'Timepoint', 'Vaccine Type', 'Vaccine Date']
    exclusions = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name='Exclusions')
    exclude_ppl = set(exclusions['Participant ID'].unique())
    
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
        vaccine = mars_data.loc[participant, '1st Vaccine Type']
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
                      ['6th vaccine', '6th vaccine type'], ['7th vaccine', '7th vaccine type'], ['8th vaccine', '8th vaccine type'],
                      ['9th vaccine', '9th vaccine type']]
        timepoint_cols = ['Dose 3', 'Booster 1', 'Booster 2', 'Booster 3', 'Booster 4', 'Booster 5', 'Booster 6']
        for i, vals in enumerate(boost_cols):
            dt, tp = vals
            date_boost = mars_data.loc[participant, dt]
            date_boost_dt = pd.to_datetime(date_boost, errors='coerce')
            boost_type = mars_data.loc[participant, tp]
            if date_boost_dt >= pd.to_datetime('2024-08-29'):
                boost_type = str(boost_type).split()[0]
                if 'novavax' in str(boost_type).lower():
                    addendum = ':Monovalent JN.1'
                else:
                    addendum = ':Monovalent KP.2'
            elif date_boost_dt >= pd.to_datetime('2023-08-29'):
                boost_type = str(boost_type).split()[0]
                addendum = ':Monovalent XBB.1.5'
            elif date_boost_dt >= pd.to_datetime('2022-08-29'):
                boost_type = str(boost_type).split()[0]
                addendum = ':Bivalent'
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
    from_long = pd.read_excel(util.mars_folder + 'MARS for D4 Long.xlsx', sheet_name='COVID Vaccinations').set_index(['Participant ID', 'Vaccination_Status'])
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
        boost_cols = [['3rd Vaccine Date', '3rd Vaccine Type '], ['4th Vaccine Date', '4th Vaccine Type'], ['5th Vaccine Date', '5th Vaccine Type'],
                      ['6th Vaccine Date', '6th Vaccine Type'], ['7th Vaccine Date', '7th Vaccine Type'], ['8th Vaccine Date', '8th Vaccine Type'],
                      ['9th Vaccine Date', '9th Vaccine Type']]
        timepoint_cols = ['Booster 1', 'Booster 2', 'Booster 3', 'Booster 4', 'Booster 5', 'Booster 6', 'Booster 7']
        for i, vals in enumerate(boost_cols):
            dt, tp = vals
            date_boost = gaea_data.loc[participant, dt]
            date_boost_dt = pd.to_datetime(date_boost, errors='coerce')
            boost_type = gaea_data.loc[participant, tp]
            if date_boost_dt >= pd.to_datetime('2024-08-29'):
                boost_type = str(boost_type).split()[0]
                if 'novavax' in str(boost_type).lower():
                    addendum = ':Monovalent JN.1'
                else:
                    addendum = ':Monovalent KP.2'
            elif date_boost_dt >= pd.to_datetime('2023-08-29'):
                boost_type = str(boost_type).split()[0]
                addendum = ':Monovalent XBB.1.5'
            elif date_boost_dt >= pd.to_datetime('2022-08-29'):
                boost_type = boost_type.split()[0]
                addendum = ':Bivalent'
            else:
                addendum = ''
            if type(date_boost) in [datetime.datetime, pd.Timestamp]:
                vaccine_stuff['Participant ID'].append(participant)
                vaccine_stuff['Timepoint'].append(timepoint_cols[i] + addendum)
                vaccine_stuff['Vaccine Type'].append(boost_type)
                vaccine_stuff['Vaccine Date'].append(date_boost)
    #pd.DataFrame(vaccine_stuff).to_excel(util.seronet_vax + 'gaea_vaccines.xlsx', index=False)
    from_tracker = pd.DataFrame(vaccine_stuff).rename(columns={'Timepoint': 'Vaccination_Status'}).set_index(['Participant ID', 'Vaccination_Status'])
    from_long = pd.read_excel(util.gaea_folder + 'GAEA for D4 Long.xlsx', sheet_name='COVID Vaccinations').set_index(['Participant ID', 'Vaccination_Status'])
    in_both = from_long.join(from_tracker, how='inner', lsuffix='_long', rsuffix='_tracker')
    in_both_correct = in_both[(in_both['SARS-CoV-2_Vaccination_Date'] == in_both['Vaccine Date']) &
                                (in_both['SARS-CoV-2_Vaccine_Type'] == in_both['Vaccine Type'])]
    date_issue = in_both[(in_both['SARS-CoV-2_Vaccination_Date'] != in_both['Vaccine Date'])]
    type_issue = in_both[(in_both['SARS-CoV-2_Vaccine_Type'] != in_both['Vaccine Type'])]
    add_to_tracker = from_long[~from_long.index.isin(from_tracker.index)]
    add_to_long = from_tracker[~from_tracker.index.isin(from_long.index)]
    with pd.ExcelWriter(util.seronet_vax + 'gaea_vaccines.xlsx') as writer:
        add_to_long.reset_index().to_excel(writer, sheet_name='Add to Long', index=False)
        date_issue.reset_index().to_excel(writer, sheet_name='Date Discrepancy', index=False)
        type_issue.reset_index().to_excel(writer, sheet_name='Type Discrepancy', index=False)
        add_to_tracker.reset_index().to_excel(writer, sheet_name='Add to Tracker', index=False)
        in_both_correct.reset_index().to_excel(writer, sheet_name='In Both (agreed)', index=False)
        from_tracker.reset_index().to_excel(writer, sheet_name='Tracker Source', index=False)
        from_long.reset_index().to_excel(writer, sheet_name='Long D4 Source', index=False)