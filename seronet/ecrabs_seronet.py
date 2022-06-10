import pandas as pd
import numpy as np
from datetime import date
import datetime
from dateutil import parser
from dateutil.relativedelta import relativedelta
from bisect import bisect_left
import sys

def process_lots():
    lots = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Processing Team/Data Sample Collection Form.xlsx', sheet_name='Lot # Sheet')
    lot_log = {}
    for _, row in lots.sort_values('Date Used').iterrows():
        mat = row['Material']
        open_date = row['Date Used'].date()
        lot = row['Lot Number']
        exp = row['EXP Date']
        cat = row['Catalog Number']
        if mat not in lot_log.keys():
            lot_log[mat] = [(open_date, lot, exp, cat)]
        else:
            lot_log[mat].append((open_date, lot, exp, cat))
    return lot_log



def get_catalog_lot_exp(coll_date, material, lot_log):
    unknown = ("Please enter into Lot Sheet", "-", "-", "-")
    if material not in lot_log.keys():
        return unknown
    else:
        lots = list(filter(lambda vals: vals[0] <= coll_date, lot_log[material]))
        if len(lots) > 0:
            return lots[-1]
        else:
            return unknown


if __name__ == '__main__':
    output_fname = 'tmp'
    if len(sys.argv) != 2:
        print("usage: python ecrabs_seronet.py output_file_name")
        exit(1)
    else:
        output_fname = sys.argv[1]
    script_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Scripts/'
    script_input = script_folder + 'input/'
    script_output = script_folder + 'output/'

    equip_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Equipment_ID', 'Equipment_Type', 'Equipment_Calibration_Due_Date', 'Comments']
    consum_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Consumable_Name', 'Consumable_Catalog_Number', 'Consumable_Lot_Number', 'Consumable_Expiration_Date', 'Comments']
    reag_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Reagent_Name', 'Reagent_Catalog_Number', 'Reagent_Lot_Number', 'Reagent_Expiration_Date', 'Comments']
    aliq_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Aliquot_ID', 'Aliquot_Volume', 'Aliquot_Units', 'Aliquot_Concentration', 'Aliquot_Initials', 'Aliquot_Tube_Type', 'Aliquot_Tube_Type_Catalog_Number', 'Aliquot_Tube_Type_Lot_Number', 'Aliquot_Tube_Type_Expiration_Date', 'Comments']
    biospec_cols = ['Participant ID', 'Sample ID', 'Date', 'Proc Comments', 'Time Collected', 'Time Received', 'Time Processed', 'Time Serum Frozen', 'Time Cells Frozen', 'Research_Participant_ID', 'Cohort', 'Visit_Number', 'Biospecimen_ID', 'Biospecimen_Type', 'Biospecimen_Collection_Date_Duration_From_Index', 'Biospecimen_Processing_Batch_ID', 'Initial_Volume_of_Biospecimen', 'Biospecimen_Collection_Company_Clinic', 'Biospecimen_Collector_Initials', 'Biospecimen_Collection_Year', 'Collection_Tube_Type', 'Collection_Tube_Type_Catalog_Number', 'Collection_Tube_Type_Lot_Number', 'Collection_Tube_Type_Expiration_Date', 'Storage_Time_at_2_8_Degrees_Celsius', 'Storage_Start_Time_at_2-8_Initials', 'Storage_End_Time_at_2-8_Initials', 'Biospecimen_Processing_Company_Clinic', 'Biospecimen_Processor_Initials', 'Biospecimen_Collection_to_Receipt_Duration', 'Biospecimen_Receipt_to_Storage_Duration', 'Centrifugation_Time', 'RT_Serum_Clotting_Time', 'Live_Cells_Hemocytometer_Count', 'Total_Cells_Hemocytometer_Count', 'Viability_Hemocytometer_Count', 'Live_Cells_Automated_Count', 'Total_Cells_Automated_Count', 'Viability_Automated_Count', 'Storage_Time_in_Mr_Frosty', 'Comments']
    ship_cols = ['Participant ID', 'Sample ID', 'Date', 'Study ID', 'Current Label', 'Material Type', 'Volume', 'Volume Unit', 'Volume Estimate', 'Vial Type', 'Vial Warnings', 'Vial Modifiers']

    iris_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/IRIS/'
    titan_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/TITAN/'
    mars_folder = '~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Clinical Research Study Operations/Umbrella Viral Sample Collection Protocol/MARS/'
    
    # include_file = mars_folder + 'MARS for D4 Long.xlsx'
    # include_sheet = 'Source Short' # Participants to include
    
    # include_ppl = pd.read_excel(include_file, sheet_name=include_sheet)['Participant ID'].unique()
    include_ppl = []
    first_date = parser.parse('1/1/2021').date()
    last_date = parser.parse('12/31/2021').date()

    ecrabs_sheets = ['Equipment', 'Consumables', 'Reagent', 'Aliquot', 'Biospecimen', 'Shipping Manifest']
    necessities_df = pd.read_excel(script_input + 'ECRAB_SERONET.xlsx', sheet_name=None)

    lot_log = process_lots()

    CRF = "Controlled-Rate Freezer"

    visit_type = "Visit Type / Samples Needed"
    iris_data = pd.read_excel(iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID'])
    iris_data['Participant ID'] = iris_data['Participant ID'].apply(lambda val: val.strip().upper())
    titan_data = pd.read_excel(titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_data['Participant ID'] = titan_data['Umbrella Corresponding Participant ID'].apply(lambda val: val.strip().upper())
    visit_type = "Visit Type / Samples Needed"
    mars_data = pd.read_excel(mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID'])
    mars_data['Participant ID'] = mars_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = [p for p in mars_data['Participant ID'].unique()] + [p for p in titan_data['Participant ID'].unique()] + [p for p in iris_data['Participant ID'].unique()]
    participant_study = {p: 'MARS' for p in mars_data['Participant ID'].unique()}
    participant_study.update({p: 'TITAN' for p in titan_data['Participant ID'].unique()})
    participant_study.update({p: 'IRIS' for p in iris_data['Participant ID'].unique()})
    mars_data.set_index('Participant ID', inplace=True)
    iris_data.set_index('Participant ID', inplace=True)
    titan_data.set_index('Participant ID', inplace=True)
    samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Sample Tracking/Sample Intake Log.xlsx', sheet_name='Sample Intake Log', header=6, dtype=str)
    collection_log = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Sample Tracking/Sample Intake Log.xlsx', sheet_name='Collection Log')
    collection_log['idx'] = collection_log['Sample ID'].apply(lambda val: str(val).strip().upper())
    collection_log.set_index('idx', inplace=True)
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    samplesClean = samples.dropna(subset=['Participant ID'])
    participant_samples = {participant: [] for participant in participants if participant in include_ppl}
    samples_of_interest = set()
    for _, sample in samplesClean.iterrows():
        sample_id = str(sample['Sample ID']).strip().upper()
        if len(sample_id) != 5:
            continue
        participant = str(sample['Participant ID']).strip().upper()
        if str(sample['Study']).strip().upper() == 'PRIORITY':
            if participant not in participant_samples.keys():
                participant_samples[participant] = []
                participant_study[participant] = 'PRIORITY'
        if participant in participant_samples.keys():
            try:
                sample_date = parser.parse(str(sample['Date Collected'])).date()
                if sample_date > last_date or sample_date < first_date:
                    continue
            except:
                print(sample['Date Collected'])
                sample_date = parser.parse('1/1/1900').date()
                continue # because we're in TIMP, we don't want dateless samples
            if participant_study[participant] == "MARS":
                if type(mars_data.loc[participant, 'Vaccine #1 Date']) != datetime.datetime or (
                    type(mars_data.loc[participant, 'Vaccine #2 Date']) != datetime.datetime
                ):
                    print("Try to find vaccine dates for", participant)
                    continue
                elif mars_data.loc[participant, 'Vaccine #1 Date'].date() < sample_date and (
                     mars_data.loc[participant, 'Vaccine #2 Date'].date() >= sample_date):
                    continue
            if participant_study[participant] == "IRIS":
                if type(iris_data.loc[participant, 'First Dose Date']) != datetime.datetime or (
                    type(iris_data.loc[participant, 'Second Dose Date']) != datetime.datetime
                ):
                    print("Try to find vaccine dates for", participant)
                    continue
                elif iris_data.loc[participant, 'First Dose Date'].date() < sample_date and (
                    iris_data.loc[participant, 'Second Dose Date'].date() >= sample_date):
                    continue
            participant_samples[participant].append((sample_date,sample_id))
            samples_of_interest.add(sample_id)
    bsl2p_samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Processing Team/Data Sample Collection Form.xlsx', sheet_name='BSL2+ Samples', header=1)
    bsl2_samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Processing Team/Data Sample Collection Form.xlsx', sheet_name='BSL2 Samples')
    shared_samples = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - Sample Tracking/Released Samples/Collaborator Samples Tracker.xlsx', sheet_name='Released Samples')
    no_pbmcs = set([str(sid).strip() for sid in shared_samples[shared_samples['Sample Type'] == 'PBMC']['Sample ID'].unique()])
    missing_info = {'Sample ID': [], 'Serum Volume': [], 'PBMC Conc': [], 'PBMC Vial Count': []}
    sample_info = {}
    for _, sample in bsl2p_samples.iterrows():
        sample_id = str(sample['Sample ID']).strip().upper()
        if sample_id not in samples_of_interest:
            continue
        try:
            serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("mlML ").split()[0])
        except:
            try:
                serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("ulUL ").split()[0]) / 1000.
            except:
                print(sample_id, sample['Total volume of serum (ml)'], type(serum_volume))
                if type(sample['Total volume of serum (ml)']) == str:
                    serum_volume = 0
                else:
                    serum_volume = sample['Total volume of serum (ml)']
        if type(serum_volume) != str and serum_volume > 4.5:
            serum_volume = 4.5
        if sample_id in no_pbmcs:
            pbmc_conc = 0
            vial_count = 0
        else:
            pbmc_conc = sample['# cells per aliquot']
            vial_count = sample['# of aliquots frozen']
            if (type(pbmc_conc) != float or pd.isna(pbmc_conc)) and type(pbmc_conc) != int:
                pbmc_conc = 0
                vial_count = 0
        if type(vial_count) in [int, float] and vial_count > 2:
            vial_count = 2
        if sample_id in collection_log.index:
            coll_time = collection_log.loc[sample_id, 'Time Collected']
        else:
            coll_time = '?'


        rec_time = sample['Time Received']
        proc_time = '???'
        serum_freeze_time = sample['Time put in -80: SERUM']
        cell_freeze_time = sample['Time put in -80: PBMC']
        proc_inits = sample['Processed by (initials)']
        viability = sample['% Viability']
        cpt_vol = sample['CPT/EDTA VOL']
        sst_vol = sample['SST VOL']
        comment = sample['Comments']

        if not ((pd.isna(serum_volume) or serum_volume == 0) and (type(pbmc_conc) == str or pd.isna(pbmc_conc) or pbmc_conc == 0)):
            sample_info[sample_id] = [serum_volume, pbmc_conc, vial_count, coll_time, rec_time, proc_time, serum_freeze_time, cell_freeze_time, proc_inits, viability, cpt_vol, sst_vol, comment]
        else:
            missing_info['Sample ID'].append(sample_id)
            missing_info['Serum Volume'].append(serum_volume)
            missing_info['PBMC Conc'].append(pbmc_conc)
            missing_info['PBMC Vial Count'].append(vial_count)
    for _, sample in bsl2_samples.iterrows():
        sample_id = str(sample['Sample ID']).strip().upper()
        if sample_id not in samples_of_interest:
            continue
        try:
            serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("mlML ").split()[0])
        except:
            try:
                serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("ulUL ").split()[0]) / 1000.
            except:
                print(sample_id, sample['Total volume of serum (ml)'], type(serum_volume))
                if type(sample['Total volume of serum (ml)']) == str:
                    serum_volume = 0
                else:
                    serum_volume = sample['Total volume of serum (ml)']
        if type(serum_volume) != str and serum_volume > 4.5:
            serum_volume = 4.5
        if sample_id in no_pbmcs:
            pbmc_conc = 0
            vial_count = 0
        else:
            pbmc_conc = sample['# cell per aliquot']
            vial_count = sample['# of aliquots frozen']
            if (type(pbmc_conc) != float or pd.isna(pbmc_conc)) and type(pbmc_conc) != int:
                pbmc_conc = 0
                vial_count = 0
        if type(vial_count) in [int, float] and vial_count > 2:
            vial_count = 2
        if sample_id in collection_log.index:
            coll_time = collection_log.loc[sample_id, 'Time Collected']
        else:
            coll_time = sample['Time collected']
        rec_time = sample['Time Received']
        proc_time = sample['Time Started Processing']
        serum_freeze_time = sample['Time put in -80: SERUM']
        cell_freeze_time = sample['Time put in -80: PBMC']
        proc_inits = sample['Processed by (initials)'] # is this our best column?
        viability = sample['% Viability']
        cpt_vol = sample['CPT/EDTA VOL']
        sst_vol = sample['SST VOL']
        comment = sample['COMMENTS']
        if not ((pd.isna(serum_volume) or serum_volume == 0) and (type(pbmc_conc) == str or pd.isna(pbmc_conc) or pbmc_conc == 0)):
            sample_info[sample_id] = [serum_volume, pbmc_conc, vial_count, coll_time, rec_time, proc_time, serum_freeze_time, cell_freeze_time, proc_inits, viability, cpt_vol, sst_vol, comment]
        else:
            missing_info['Sample ID'].append(sample_id)
            missing_info['Serum Volume'].append(serum_volume)
            missing_info['PBMC Conc'].append(pbmc_conc)
            missing_info['PBMC Vial Count'].append(vial_count)
    mit_days = [0, 30, 60, 90, 180, 300, 540, 720]
    pri_months = [0, 6, 12]
    data = {'Cohort': [], 'SERONET ID': [], 'Days from Index': [], 'Vaccine': [], '1st Dose Date': [], 'Days to 1st': [], '2nd Dose Date': [], 'Days to 2nd': [], 'Boost Vaccine': [], 'Boost Date': [], 'Days to Boost': [], 'Participant ID': [], 'Date': [], 'Post-Baseline': [], 'Sample ID': [], 'Volume of Serum Collected (mL)': [], 'PBMC concentration per mL (x10^6)': [], '# of PBMC vials': [], 'coll_time': [], 'rec_time': [], 'proc_time': [], 'serum_freeze_time': [], 'cell_freeze_time': [], 'proc_inits': [], 'viability': [], 'cpt_vol': [], 'sst_vol': [], 'proc_comment': []}
    short_window = 14
    long_window = 21
    participant_data = {}
    for participant, samples in participant_samples.items():
        try:
            samples.sort(key=lambda x: x[0])
        except:
            print(participant)
            print(samples)
            print()
            continue
        if len(samples) < 1:
            print(participant)
            print()
            continue
        baseline = samples[0][0]
        participant_data[participant] = {}
        times = [d for d in mit_days]
        if participant_study[participant] == 'TITAN':
            seronet_id = "14_T{}".format(samples[0][1])
            index_date = titan_data.loc[participant, '3rd Dose Vaccine Date'].date()
            participant_data[participant]['Vaccine'] = titan_data.loc[participant, 'Vaccine Type']
            try:
                participant_data[participant]['1st Dose Date'] = titan_data.loc[participant, 'Vaccine #1 Date'].date()
            except:
                participant_data[participant]['1st Dose Date'] = titan_data.loc[participant, 'Vaccine #1 Date']
            try:
                participant_data[participant]['2nd Dose Date'] = titan_data.loc[participant, 'Vaccine #2 Date'].date()
            except:
                participant_data[participant]['2nd Dose Date'] = titan_data.loc[participant, 'Vaccine #2 Date']
            try:
                participant_data[participant]['Boost Date'] = titan_data.loc[participant, '3rd Dose Vaccine Date'].date()
            except:
                participant_data[participant]['Boost Date'] = titan_data.loc[participant, '3rd Dose Vaccine Date']
            participant_data[participant]['Boost Vaccine'] = titan_data.loc[participant, '3rd Dose Vaccine Type']
        elif participant_study[participant] == 'IRIS':
            seronet_id = "14_I{}".format(samples[0][1])
            try:
                index_date = iris_data.loc[participant, 'Second Dose Date'].date()
            except:
                print("Participant {} is missing vaccine dose 2 date (dose 1 on {})".format(participant, iris_data.loc[participant, 'First Dose Date']))
                index_date = 'N/A'
            participant_data[participant]['Vaccine'] = iris_data.loc[participant, 'Which Vaccine?']
            try:
                participant_data[participant]['1st Dose Date'] = iris_data.loc[participant, 'First Dose Date'].date()
            except:
                participant_data[participant]['1st Dose Date'] = iris_data.loc[participant, 'First Dose Date']
            try:
                participant_data[participant]['2nd Dose Date'] = iris_data.loc[participant, 'Second Dose Date'].date()
            except:
                participant_data[participant]['2nd Dose Date'] = iris_data.loc[participant, 'Second Dose Date']
            try:
                participant_data[participant]['Boost Date'] = iris_data.loc[participant, 'Third Dose Date'].date()
            except:
                participant_data[participant]['Boost Date'] = iris_data.loc[participant, 'Third Dose Date']
            participant_data[participant]['Boost Vaccine'] = iris_data.loc[participant, 'Third Dose Type']
        elif participant_study[participant] == 'MARS':
            seronet_id = "14_M{}".format(samples[0][1])
            try:
                index_date = mars_data.loc[participant, 'Vaccine #2 Date'].date()
            except:
                print("Participant {} is missing vaccine dose 2 date (dose 1 on {})".format(participant, mars_data.loc[participant, 'Vaccine #1 Date']))
                index_date = 'N/A'
            participant_data[participant]['Vaccine'] = mars_data.loc[participant, 'Vaccine Name']
            try:
                participant_data[participant]['1st Dose Date'] = mars_data.loc[participant, 'Vaccine #1 Date'].date()
            except:
                participant_data[participant]['1st Dose Date'] = mars_data.loc[participant, 'Vaccine #1 Date']
            try:
                participant_data[participant]['2nd Dose Date'] = mars_data.loc[participant, 'Vaccine #2 Date'].date()
            except:
                participant_data[participant]['2nd Dose Date'] = mars_data.loc[participant, 'Vaccine #2 Date']
            try:
                participant_data[participant]['Boost Date'] = mars_data.loc[participant, '3rd Vaccine'].date()
            except:
                participant_data[participant]['Boost Date'] = mars_data.loc[participant, '3rd Vaccine']
            participant_data[participant]['Boost Vaccine'] = mars_data.loc[participant, '3rd Vaccine Type ']
        elif participant_study[participant] == 'PRIORITY':
            seronet_id = "14_P{}".format(samples[0][1])
            index_date = baseline
            times = [m for m in pri_months]
            participant_data[participant]['Vaccine'] = ''
            participant_data[participant]['1st Dose Date'] = ''
            participant_data[participant]['2nd Dose Date'] = ''
            participant_data[participant]['Boost Date'] = ''
            participant_data[participant]['Boost Vaccine'] = ''
        dts = [s[0] for s in samples]
        if type(index_date) == datetime.date:
            start_idx = bisect_left(dts, index_date)
        else:
            start_idx = 0
        if start_idx > 0:
            start_idx -= 1
        samples = samples[start_idx:] # This is our messy "remove all but one pre-index sample"
        one_in_window = False
        for date_, sample_id in samples:
            include = False
            while True: # Very slapdash
                if len(times) == 0:
                    break
                if participant_study[participant] == 'PRIORITY':
                    timediff = relativedelta(months=times[0])
                    window = long_window
                else:
                    timediff = relativedelta(days=times[0])
                    if times[0] < 100:
                        window = short_window
                    else:
                        window = long_window
                if type(index_date) == datetime.date and times[0] == 0: # We may later have to add TIM baselines here, try to distinguish
                    if date_ <= index_date:
                        include = True
                        one_in_window = True
                        times.pop(0)
                        break
                    else:
                        times.pop(0) # Hopefully we're TIM and we missed the baseline
                elif type(index_date) == datetime.date:
                    if abs(int((index_date + timediff - date_).days)) <= window:
                        one_in_window = True
                        include = True
                        times.pop(0)
                        break
                    elif date_ >= index_date + timediff + relativedelta(days=window):
                        # samples_to_write.extend(['No Visit', 'No Visit', 'No Visit'])
                        times.pop(0)
                    else:
                        break
                else:
                    print(index_date)
                    print()
                    for time in times:
                        pass
                        # samples_to_write.extend(['-', '-', '-'])
                    times = []
                    break
            if not include:
                continue
            if type(index_date) != datetime.date:
                days_from_ind = '??? Double check!'
            else:
                days_from_ind = int((date_ - index_date).days)
            sample_id = str(sample_id).strip().upper()
            data['Cohort'].append(participant_study[participant])
            data['SERONET ID'].append(seronet_id)
            data['Days from Index'].append(days_from_ind)
            data['Vaccine'].append(participant_data[participant]['Vaccine'])
            data['1st Dose Date'].append(participant_data[participant]['1st Dose Date'])
            try:
                data['Days to 1st'].append(int((date_ - data['1st Dose Date'][-1]).days))
            except:
                data['Days to 1st'].append('')
            data['2nd Dose Date'].append(participant_data[participant]['2nd Dose Date'])
            try:
                data['Days to 2nd'].append(int((date_ - data['2nd Dose Date'][-1]).days))
            except:
                data['Days to 2nd'].append('')
            data['Boost Date'].append(participant_data[participant]['Boost Date'])
            try:
                data['Days to Boost'].append(int((date_ - data['Boost Date'][-1]).days))
            except:
                data['Days to Boost'].append('')
            data['Boost Vaccine'].append(participant_data[participant]['Boost Vaccine'])
            data['Participant ID'].append(participant)
            data['Date'].append(date_)
            data['Post-Baseline'].append((date_ - baseline).days)
            data['Sample ID'].append(sample_id)
            unknown = ['?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', 'Processing Data Unavailable']
            if sample_id in sample_info.keys():
                serum_volume, pbmc_conc, vial_count, coll_time, rec_time, proc_time, serum_freeze_time, cell_freeze_time, proc_inits, viability, cpt_vol, sst_vol, comment = sample_info[sample_id]
            else:
                serum_volume, pbmc_conc, vial_count, coll_time, rec_time, proc_time, serum_freeze_time, cell_freeze_time, proc_inits, viability, cpt_vol, sst_vol, comment = unknown
            if participant_study[participant] == 'PRIORITY':
                pbmc_conc = 0 # May have to change to "N/A"
                vial_count = 0 # May have to change to "N/A"
            data['Volume of Serum Collected (mL)'].append(serum_volume)
            data['PBMC concentration per mL (x10^6)'].append(pbmc_conc)
            data['# of PBMC vials'].append(vial_count)
            data['coll_time'].append(coll_time)
            data['rec_time'].append(rec_time)
            data['proc_time'].append(proc_time)
            data['serum_freeze_time'].append(serum_freeze_time)
            data['cell_freeze_time'].append(cell_freeze_time)
            data['proc_inits'].append(proc_inits)
            data['viability'].append(viability)
            data['cpt_vol'].append(cpt_vol)
            data['sst_vol'].append(sst_vol)
            data['proc_comment'].append(comment)
    report = pd.DataFrame(data)

    future_output = {}
    future_output['Equipment'] = {col: [] for col in equip_cols}
    future_output['Consumables'] = {col: [] for col in consum_cols}
    future_output['Reagent'] = {col: [] for col in reag_cols}
    future_output['Aliquot'] = {col: [] for col in aliq_cols}
    future_output['Biospecimen'] = {col: [] for col in biospec_cols}
    future_output['Shipping Manifest'] = {col: [] for col in ship_cols}
    participant_visits = {}
    for _, row in report.iterrows():
        participant = row['Participant ID']
        sample_id = row['Sample ID']
        seronet_id = row['SERONET ID']
        if participant not in participant_visits.keys():
            participant_visits[participant] = 1
        else:
            participant_visits[participant] += 1
        visit_num = participant_visits[participant]
        if visit_num == 1:
            visit = 'Baseline(1)'
        else:
            visit = visit_num
        serum_id = '{}_1{}'.format(seronet_id, str(visit_num).zfill(2))
        cells_id = '{}_2{}'.format(seronet_id, str(visit_num).zfill(2))
        serum_vol = row['Volume of Serum Collected (mL)']
        cell_count = row['PBMC concentration per mL (x10^6)']
        vial_count = row['# of PBMC vials']
        rec_time = row['rec_time']
        proc_time = row['proc_time']
        coll_time = row['coll_time']
        serum_freeze_time = row['serum_freeze_time']
        cell_freeze_time = row['cell_freeze_time']
        proc_inits = row['proc_inits']
        viability = row['viability']
        cpt_vol = row['cpt_vol']
        sst_vol = row['sst_vol']
        proc_comment = row['proc_comment']
        try:
            if cell_count > 20.:
                aliq_cells = 10.
            elif type(vial_count) != str:
                aliq_cells = cell_count / vial_count
            else:
                aliq_cells = cell_count
            live_cells = cell_count * (10**6)
            try:
                all_cells = live_cells * 100 / viability
            except:
                all_cells = "Viability needs fixing"
        except:
            aliq_cells = '???'
            live_cells = cell_count
            all_cells = '???'
            proc_comment = str(proc_comment) + '; cell count is not number, please fix'
            print(sample_id, cell_count)
        '''
        Equipment First
        '''
        # equip_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Equipment_ID', 'Equipment_Type', 'Equipment_Calibration_Due_Date', 'Comments']
        sec = 'Equipment'
        df = necessities_df[sec]
        add_to = future_output[sec]
        if type(serum_vol) == str or (type(serum_vol) == float and pd.isna(serum_vol)):
            print(sample_id, serum_vol, " is not formatted well. Please fix")
        if type(vial_count) == str or (type(vial_count) == float and pd.isna(vial_count)):
            print(sample_id, vial_count, " PBMCs is not formatted well. Please fix")
        for _, mat_row in df.iterrows():
            if type(vial_count) == str or (type(vial_count) == float and pd.isna(vial_count)):
                # print(sample_id, vial_count, " PBMCs is not formatted well. Please fix")
                continue
            elif row['# of PBMC vials'] > 0 and (mat_row['Stype'] == 'Cells' or mat_row['Stype'] == 'Both'):
                add_to['Participant ID'].append(participant)
                add_to['Sample ID'].append(sample_id)
                add_to['Date'].append(row['Date'])
                add_to['Biospecimen_ID'].append(cells_id)
                add_to['Equipment_ID'].append(mat_row['Equipment_ID'])
                add_to['Equipment_Type'].append(mat_row['Equipment_Type'])
                add_to['Equipment_Calibration_Due_Date'].append(mat_row['Equipment_Calibration_Due_Date'])
                add_to['Comments'].append(mat_row['Comments'])
        '''
        Consumables Next
        '''
        # consum_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Consumable_Name', 'Consumable_Catalog_Number', 'Consumable_Lot_Number', 'Consumable_Expiration_Date', 'Comments']
        # lot_log[mat] = [(open_date, lot, exp, cat)] Format of lot_log
        sec = 'Consumables'
        df = necessities_df[sec]
        add_to = future_output[sec]
        for _, mat_row in df.iterrows():
            if type(vial_count) == str or (type(vial_count) == float and pd.isna(vial_count)):
                # print(sample_id, vial_count, " serum is not formatted well. Please fix")
                continue
            elif row['# of PBMC vials'] > 0 and (mat_row['Stype'] == 'Cells' or mat_row['Stype'] == 'Both'):
                add_to['Participant ID'].append(participant)
                add_to['Sample ID'].append(sample_id)
                add_to['Date'].append(row['Date'])
                add_to['Biospecimen_ID'].append(cells_id)
                cname = mat_row['Consumable_Name']
                add_to['Consumable_Name'].append(cname)
                odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], cname, lot_log)
                add_to['Consumable_Catalog_Number'].append(cat)
                add_to['Consumable_Lot_Number'].append(lot)
                add_to['Consumable_Expiration_Date'].append(exp)
                add_to['Comments'].append('')
        '''
        Reagents Next
        '''
        # reag_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Reagent_Name', 'Reagent_Catalog_Number', 'Reagent_Lot_Number', 'Reagent_Expiration_Date', 'Comments']
        # lot_log[mat] = [(open_date, lot, exp, cat)] Format of lot_log
        sec = 'Reagent'
        df = necessities_df[sec]
        add_to = future_output[sec]
        for _, mat_row in df.iterrows():
            if type(vial_count) == str or (type(vial_count) == float and pd.isna(vial_count)):
                # print(sample_id, vial_count, " serum is not formatted well. Please fix")
                continue
            elif row['# of PBMC vials'] > 0 and (mat_row['Stype'] == 'Cells' or mat_row['Stype'] == 'Both'):
                add_to['Participant ID'].append(participant)
                add_to['Sample ID'].append(sample_id)
                add_to['Date'].append(row['Date'])
                add_to['Biospecimen_ID'].append(cells_id)
                rname = mat_row['Reagent_Name']
                add_to['Reagent_Name'].append(rname)
                odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], rname, lot_log)
                add_to['Reagent_Catalog_Number'].append(cat)
                add_to['Reagent_Lot_Number'].append(lot)
                add_to['Reagent_Expiration_Date'].append(exp)
                add_to['Comments'].append('')
        '''
        Aliquot Penultimate
        '''
        # aliq_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Aliquot_ID', 'Aliquot_Volume', 'Aliquot_Units', 'Aliquot_Concentration', 'Aliquot_Initials', 'Aliquot_Tube_Type', 'Aliquot_Tube_Type_Catalog_Number', 'Aliquot_Tube_Type_Lot_Number', 'Aliquot_Tube_Type_Expiration_Date', 'Comments']
        sec = 'Aliquot'
        df = necessities_df[sec]
        add_to = future_output[sec]
        proc_inits = row['proc_inits']
        if type(serum_vol) != str and not (type(serum_vol) == float and pd.isna(serum_vol)) and serum_vol > 0:
            add_to['Participant ID'].append(participant)
            add_to['Sample ID'].append(sample_id)
            add_to['Date'].append(row['Date'])
            add_to['Biospecimen_ID'].append(serum_id)
            add_to['Aliquot_ID'].append("{}_1".format(serum_id))
            add_to['Aliquot_Volume'].append(serum_vol)
            add_to['Aliquot_Units'].append('mL')
            add_to['Aliquot_Concentration'].append("N/A")
            add_to['Aliquot_Initials'].append(proc_inits)
            tname = 'CRYOTUBE 4.5ML'
            add_to['Aliquot_Tube_Type'].append(tname)
            odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], tname, lot_log)
            add_to['Aliquot_Tube_Type_Catalog_Number'].append(cat)
            add_to['Aliquot_Tube_Type_Lot_Number'].append(lot)
            add_to['Aliquot_Tube_Type_Expiration_Date'].append(exp)
            add_to['Comments'].append('')
        if not (type(vial_count) == str or (type(vial_count) == float and pd.isna(vial_count))):
            for i in range(min(int(vial_count), 2)):
                add_to['Participant ID'].append(participant)
                add_to['Sample ID'].append(sample_id)
                add_to['Date'].append(row['Date'])
                add_to['Biospecimen_ID'].append(cells_id)
                add_to['Aliquot_ID'].append("{}_{}".format(cells_id, i + 1))
                add_to['Aliquot_Volume'].append(1)
                add_to['Aliquot_Units'].append('mL')
                add_to['Aliquot_Concentration'].append("{:.2f}".format(aliq_cells * 1000000.0))
                add_to['Aliquot_Initials'].append(proc_inits)
                tname = 'CRYOTUBE 1.8ML'
                add_to['Aliquot_Tube_Type'].append(tname)
                odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], tname, lot_log)
                add_to['Aliquot_Tube_Type_Catalog_Number'].append(cat)
                add_to['Aliquot_Tube_Type_Lot_Number'].append(lot)
                add_to['Aliquot_Tube_Type_Expiration_Date'].append(exp)
                add_to['Comments'].append('')
        '''
        Biospecimen Last
        '''
        sec = 'Biospecimen'
        df = necessities_df[sec]
        add_to = future_output[sec]
        # biospec_cols = ['Participant ID', 'Sample ID', 'Date', 'Proc Comments', 'Research_Participant_ID', 'Cohort', 'Visit_Number', 'Biospecimen_ID', 'Biospecimen_Type', 'Biospecimen_Collection_Date_Duration_From_Index', 'Biospecimen_Processing_Batch_ID', 'Initial_Volume_of_Biospecimen', 'Biospecimen_Collection_Company_Clinic', 'Biospecimen_Collector_Initials', 'Biospecimen_Collection_Year', 'Collection_Tube_Type', 'Collection_Tube_Type_Catalog_Number', 'Collection_Tube_Type_Lot_Number', 'Collection_Tube_Type_Expiration_Date', 'Storage_Time_at_2_8_Degrees_Celsius', 'Storage_Start_Time_at_2-8_Initials', 'Storage_End_Time_at_2-8_Initials', 'Biospecimen_Processing_Company_Clinic', 'Biospecimen_Processor_Initials', 'Biospecimen_Collection_to_Receipt_Duration', 'Biospecimen_Receipt_to_Storage_Duration', 'Centrifugation_Time', 'RT_Serum_Clotting_Time', 'Live_Cells_Hemocytometer_Count', 'Total_Cells_Hemocytometer_Count', 'Viability_Hemocytometer_Count', 'Live_Cells_Automated_Count', 'Total_Cells_Automated_Count', 'Viability_Automated_Count', 'Storage_Time_in_Mr_Frosty', 'Comments']
        if type(serum_vol) != str and not (type(serum_vol) == float and pd.isna(serum_vol)) and serum_vol > 0:
            add_to['Participant ID'].append(participant)
            add_to['Sample ID'].append(sample_id)
            add_to['Date'].append(row['Date'])
            add_to['Proc Comments'].append(proc_comment)
            add_to['Time Collected'].append(coll_time)
            add_to['Time Received'].append(rec_time)
            add_to['Time Processed'].append(proc_time)
            add_to['Time Serum Frozen'].append(serum_freeze_time)
            add_to['Time Cells Frozen'].append(cell_freeze_time)
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(participant_study[participant])
            add_to['Visit_Number'].append(visit)
            add_to['Biospecimen_ID'].append(serum_id)
            add_to['Biospecimen_Type'].append('Serum')
            add_to['Biospecimen_Collection_Date_Duration_From_Index'].append(row['Days from Index'])
            add_to['Biospecimen_Processing_Batch_ID'].append('???') # decide how to handle
            add_to['Initial_Volume_of_Biospecimen'].append(sst_vol)
            add_to['Biospecimen_Collection_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            if sample_id in collection_log.index:
                add_to['Biospecimen_Collector_Initials'].append(collection_log.loc[sample_id, 'Collected By'])
            else:
                add_to['Biospecimen_Collector_Initials'].append('???')
            add_to['Biospecimen_Collection_Year'].append(2021)
            tname = 'SST'
            add_to['Collection_Tube_Type'].append(tname)
            odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], tname, lot_log)
            add_to['Collection_Tube_Type_Catalog_Number'].append(cat)
            add_to['Collection_Tube_Type_Lot_Number'].append(lot)
            add_to['Collection_Tube_Type_Expiration_Date'].append(exp)
            add_to['Storage_Time_at_2_8_Degrees_Celsius'].append('N/A')
            add_to['Storage_Start_Time_at_2-8_Initials'].append('N/A')
            add_to['Storage_End_Time_at_2-8_Initials'].append('N/A')
            add_to['Biospecimen_Processing_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            add_to['Biospecimen_Processor_Initials'].append(proc_inits)
            try:
                add_to['Biospecimen_Collection_to_Receipt_Duration'].append((datetime.datetime.combine(date.min, rec_time) - datetime.datetime.combine(date.min, coll_time)) / datetime.timedelta(hours=1))
            except Exception as e:
                print(e)
                print(coll_time, type(coll_time))
                print(rec_time, type(rec_time))
                print()
                add_to['Biospecimen_Collection_to_Receipt_Duration'].append('???')
            try:
                add_to['Biospecimen_Receipt_to_Storage_Duration'].append((datetime.datetime.combine(date.min, serum_freeze_time) - datetime.datetime.combine(date.min, rec_time)) / datetime.timedelta(hours=1))
            except Exception as e:
                print(e)
                add_to['Biospecimen_Receipt_to_Storage_Duration'].append('???')
                print(serum_freeze_time, type(serum_freeze_time))
                print(rec_time, type(rec_time))
                print()
            add_to['Centrifugation_Time'].append(40)
            try:
                add_to['RT_Serum_Clotting_Time'].append((datetime.datetime.combine(date.min, proc_time) - datetime.datetime.combine(date.min, coll_time)) / datetime.timedelta(hours=1))
            except Exception as e:
                print(e)
                print(proc_time, type(proc_time))
                print(coll_time, type(coll_time))
                print()
                add_to['RT_Serum_Clotting_Time'].append('???')
            add_to['Live_Cells_Hemocytometer_Count'].append('N/A')
            add_to['Total_Cells_Hemocytometer_Count'].append('N/A')
            add_to['Viability_Hemocytometer_Count'].append('N/A')
            add_to['Live_Cells_Automated_Count'].append('N/A')
            add_to['Total_Cells_Automated_Count'].append('N/A')
            add_to['Viability_Automated_Count'].append('N/A')
            add_to['Storage_Time_in_Mr_Frosty'].append('N/A')
            add_to['Comments'].append('')
        if not (type(vial_count) == str or (type(vial_count) == float and pd.isna(vial_count))) and row['# of PBMC vials'] > 0:
            add_to['Participant ID'].append(participant)
            add_to['Sample ID'].append(sample_id)
            add_to['Date'].append(row['Date'])
            add_to['Proc Comments'].append(proc_comment)
            add_to['Time Collected'].append(coll_time)
            add_to['Time Received'].append(rec_time)
            add_to['Time Processed'].append(proc_time)
            add_to['Time Serum Frozen'].append(serum_freeze_time)
            add_to['Time Cells Frozen'].append(cell_freeze_time)
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(participant_study[participant])
            add_to['Visit_Number'].append(visit)
            add_to['Biospecimen_ID'].append(cells_id)
            add_to['Biospecimen_Type'].append('PBMC')
            add_to['Biospecimen_Collection_Date_Duration_From_Index'].append(row['Days from Index'])
            add_to['Biospecimen_Processing_Batch_ID'].append('???') # decide how to handle
            add_to['Initial_Volume_of_Biospecimen'].append(sst_vol)
            add_to['Biospecimen_Collection_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            if sample_id in collection_log.index:
                add_to['Biospecimen_Collector_Initials'].append(collection_log.loc[sample_id, 'Collected By'])
            else:
                add_to['Biospecimen_Collector_Initials'].append('???')
            add_to['Biospecimen_Collection_Year'].append(2021)
            tname = 'CPT'
            add_to['Collection_Tube_Type'].append(tname)
            odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], tname, lot_log)
            add_to['Collection_Tube_Type_Catalog_Number'].append(cat)
            add_to['Collection_Tube_Type_Lot_Number'].append(lot)
            add_to['Collection_Tube_Type_Expiration_Date'].append(exp)
            add_to['Storage_Time_at_2_8_Degrees_Celsius'].append('N/A')
            add_to['Storage_Start_Time_at_2-8_Initials'].append('N/A')
            add_to['Storage_End_Time_at_2-8_Initials'].append('N/A')
            add_to['Biospecimen_Processing_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            add_to['Biospecimen_Processor_Initials'].append(proc_inits)
            try:
                add_to['Biospecimen_Collection_to_Receipt_Duration'].append((datetime.datetime.combine(date.min, rec_time) - datetime.datetime.combine(date.min, coll_time)) / datetime.timedelta(hours=1))
            except Exception as e:
                print(e)
                print(coll_time, type(coll_time))
                print(rec_time, type(rec_time))
                print()
                add_to['Biospecimen_Collection_to_Receipt_Duration'].append('???')
            try:
                add_to['Biospecimen_Receipt_to_Storage_Duration'].append((datetime.datetime.combine(date.min, cell_freeze_time) - datetime.datetime.combine(date.min, rec_time)) / datetime.timedelta(hours=1))
            except Exception as e:
                print(e)
                print(cell_freeze_time, type(cell_freeze_time))
                print(rec_time, type(rec_time))
                print()
                add_to['Biospecimen_Receipt_to_Storage_Duration'].append('???')
            add_to['Centrifugation_Time'].append(60)
            add_to['RT_Serum_Clotting_Time'].append('N/A')
            add_to['Live_Cells_Hemocytometer_Count'].append('N/A')
            add_to['Total_Cells_Hemocytometer_Count'].append('N/A')
            add_to['Viability_Hemocytometer_Count'].append('N/A')
            add_to['Live_Cells_Automated_Count'].append(live_cells)
            add_to['Total_Cells_Automated_Count'].append(all_cells)
            add_to['Viability_Automated_Count'].append(viability)
            if live_cells >= 19_999_999:
                add_to['Storage_Time_in_Mr_Frosty'].append('N/A')
            else:
                try:
                    add_to['Storage_Time_in_Mr_Frosty'].append((datetime.datetime.combine(date.min + datetime.timedelta(days=1), datetime.time(10,30,0)) - datetime.datetime.combine(date.min, cell_freeze_time)) / datetime.timedelta(hours=1))
                except:
                    add_to['Storage_Time_in_Mr_Frosty'].append('PLEASE FILL')
            add_to['Comments'].append('')
        '''
        Just Kidding Shipping Manifest
        '''
        # ship_cols = ['Participant ID', 'Sample ID', 'Date', 'Study ID', 'Current Label', 'Material Type', 'Volume', 'Volume Unit', 'Volume Estimate', 'Vial Type', 'Vial Warnings', 'Vial Modifiers']
        sec = 'Shipping Manifest'
        add_to = future_output[sec]
        study_id = 'LP003'
        if type(serum_vol) != str and not (type(serum_vol) == float and pd.isna(serum_vol)) and serum_vol > 0:
            add_to['Participant ID'].append(participant)
            add_to['Sample ID'].append(sample_id)
            add_to['Date'].append(row['Date'])
            add_to['Study ID'].append(study_id)
            add_to['Current Label'].append("{}_1".format(serum_id))
            add_to['Material Type'].append('Serum')
            add_to['Volume'].append(serum_vol)
            add_to['Volume Unit'].append('mL')
            add_to['Volume Estimate'].append('actual')
            tname = 'CRYOTUBE 4.5ML'
            add_to['Vial Type'].append(tname)
            add_to['Vial Warnings'].append('')
            add_to['Vial Modifiers'].append('')

    writer = pd.ExcelWriter('~/The Mount Sinai Hospital/Simon Lab - Processing Team/{}.xlsx'.format(output_fname))
    for sname, df in future_output.items():
        pd.DataFrame(df).to_excel(writer, sheet_name=sname, index=False)
    writer.save()
    report.to_excel(script_output + 'SERONET_In_Window_Data_biospecimen_companion.xlsx', index=False)

