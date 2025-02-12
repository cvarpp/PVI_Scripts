import pandas as pd
# import datetime
import argparse
import util
from seronet.d4_all_data import pull_from_source
from helpers import clean_sample_id
from lots_consolidated import compress_lots

PVI_list = ['MVK','NL','ACA','JN']

def process_lots():
    equip_lots = compress_lots().set_index(['Date Used', 'Material'])
    return equip_lots

def get_catalog_lot_exp(coll_date, material, lot_log):
    unknown = ("Please enter into Lot Sheet", "-", "-", "-")
    if (pd.to_datetime(coll_date), material.title()) not in lot_log.index:
        return unknown
    else:
        row = lot_log.loc[(pd.to_datetime(coll_date), material.title()), :]
        return [row['odate'], row['Lot Number'], row['EXP Date'], row['Catalog Number']]

def time_diff_wrapper(t_start, t_end, annot):
    try:
        time_diff = (pd.to_datetime(str(t_end), errors='coerce') - pd.to_datetime(str(t_start), errors='coerce'))
        if pd.isna(time_diff):
            time_diff = 'Missing'
        else:
            time_diff /= pd.Timedelta(hours=1)
    except Exception as e:
        if not (pd.isna(t_end) or pd.isna(t_start)):
            print(annot, e)
            print(t_start, type(t_start))
            print(t_end, type(t_end))
            print()
        time_diff = 'Missing'
    return time_diff

def make_ecrabs(source, first_date=pd.to_datetime('1/1/2021'), last_date=pd.to_datetime('1/1/2040'), output_fname='tmp', debug=False):
    issues = set()
    first_date = first_date.date()
    last_date = last_date.date()
    equip_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Equipment_ID', 'Equipment_Type', 'Equipment_Calibration_Due_Date', 'Comments']
    consum_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Consumable_Name', 'Consumable_Catalog_Number', 'Consumable_Lot_Number', 'Consumable_Expiration_Date', 'Comments']
    reag_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Reagent_Name', 'Reagent_Catalog_Number', 'Reagent_Lot_Number', 'Reagent_Expiration_Date', 'Comments']
    aliq_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Aliquot_ID', 'Aliquot_Volume', 'Aliquot_Units', 'Aliquot_Concentration', 'Aliquot_Initials', 'Aliquot_Tube_Type', 'Aliquot_Tube_Type_Catalog_Number', 'Aliquot_Tube_Type_Lot_Number', 'Aliquot_Tube_Type_Expiration_Date', 'Comments']
    biospec_cols = ['Participant ID', 'Sample ID', 'Date', 'Proc Comments', 'Time Collected', 'Time Received', 'Time Processed', 'Time Serum Frozen', 'Time Cells Frozen', 'Time in LN', 'Research_Participant_ID', 'Cohort', 'Visit_Number', 'Biospecimen_ID', 'Biospecimen_Type', 'Biospecimen_Collection_Date_Duration_From_Index', 'Biospecimen_Processing_Batch_ID', 'Initial_Volume_of_Biospecimen', 'Biospecimen_Collection_Company_Clinic', 'Biospecimen_Collector_Initials', 'Biospecimen_Collection_Year', 'Collection_Tube_Type', 'Collection_Tube_Type_Catalog_Number', 'Collection_Tube_Type_Lot_Number', 'Collection_Tube_Type_Expiration_Date', 'Storage_Time_at_2_8_Degrees_Celsius', 'Storage_Start_Time_at_2-8_Initials', 'Storage_End_Time_at_2-8_Initials', 'Biospecimen_Processing_Company_Clinic', 'Biospecimen_Processor_Initials', 'Biospecimen_Collection_to_Receipt_Duration', 'Biospecimen_Receipt_to_Storage_Duration', 'Centrifugation_Time', 'RT_Serum_Clotting_Time', 'Live_Cells_Hemocytometer_Count', 'Total_Cells_Hemocytometer_Count', 'Viability_Hemocytometer_Count', 'Live_Cells_Automated_Count', 'Total_Cells_Automated_Count', 'Viability_Automated_Count', 'Storage_Time_in_Mr_Frosty', 'Comments']
    ship_cols = ['Participant ID', 'Sample ID', 'Date', 'Study ID', 'Current Label', 'Material Type', 'Volume', 'Volume Unit', 'Volume Estimate', 'Vial Type', 'Vial Warnings', 'Vial Modifiers']

    exclusions = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name='Exclusions')
    exclude_ppl = set(exclusions['Participant ID'].unique())
    exclude_ids = set(exclusions['Research_Participant_ID'].unique())
    # ecrabs_sheets = ['Equipment', 'Consumables', 'Reagent', 'Aliquot', 'Biospecimen', 'Shipping Manifest']
    necessities_df = pd.read_excel(util.script_input + 'ECRAB_SERONET.xlsx', sheet_name=None)

    lot_log = process_lots()

    future_output = {}
    future_output['Equipment'] = {col: [] for col in equip_cols}
    future_output['Consumables'] = {col: [] for col in consum_cols}
    future_output['Reagent'] = {col: [] for col in reag_cols}
    future_output['Aliquot'] = {col: [] for col in aliq_cols}
    future_output['Biospecimen'] = {col: [] for col in biospec_cols}
    future_output['Shipping Manifest'] = {col: [] for col in ship_cols}
    participant_visits = {}
    sample_visits = pd.read_excel(util.seronet_data + 'SERONET Key.xlsx', sheet_name='Source').drop_duplicates(subset='Sample ID').assign(sample_id=clean_sample_id).set_index('sample_id')
    sample_visits['Visit Num'] = sample_visits['Biospecimen_ID'].str[-2:].apply(lambda val: int(val) if val[-1].isdigit() else val)
    for _, row in source.iterrows():
        proc_comment = row['proc_comment']
        participant = row['Participant ID']
        sample_id = str(row['Sample ID'])
        seronet_id = row['Seronet ID']
        cohort = row['Cohort']
        if participant in exclude_ppl or seronet_id in exclude_ids:
            continue
        if sample_id in sample_visits.index:
            visit_num = sample_visits.loc[sample_id, 'Visit Num']
        else:
            visit_num = 0
        if type(visit_num) == int:
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
        total_cell_count = row['Total Cell Count (x10^6)']
        aliq_cell_count = row['PBMC concentration per mL (x10^6)']
        last_aliq_cell_count = row['Cells in Last Aliquot']
        vial_count = row['# of PBMC vials']
        rec_time = row['rec_time']
        proc_time = row['proc_time']
        coll_time = row['coll_time']
        #print(row['coll_inits'], type(row['coll_inits']))
        coll_inits = str(row['coll_inits']).strip().upper()
        if coll_inits not in PVI_list:
            coll_inits = "CCT (Clinical Care Team)"

        serum_freeze_time = pd.to_datetime(str(row['serum_freeze_time']), errors='coerce')
        cell_freeze_time = pd.to_datetime(str(row['cell_freeze_time']), errors='coerce')
        if pd.isna(serum_freeze_time):
            serum_freeze_time = cell_freeze_time
        if pd.isna(cell_freeze_time):
            cell_freeze_time = serum_freeze_time
        ln_time = pd.to_datetime(str(row['Time in LN']), errors='coerce')

        proc_inits = row['proc_inits']
        viability = row['viability']
        cpt_vol = row['cpt_vol']
        sst_vol = row['sst_vol']
        try:
            total_live_cells = total_cell_count * (10**6)
            aliq_live_cells = aliq_cell_count * (10**6)
            last_aliq_live_cells = last_aliq_cell_count * (10**6)
            try:
                total_all_cells = total_live_cells * 100 / viability
            except:
                total_all_cells = "Viability needs fixing"
                if vial_count != 0:
                    issues.add(sample_id)
        except Exception as e:
            print(e)
            total_live_cells = "Missing"
            aliq_live_cells = "Missing"
            last_aliq_live_cells = "Missing"
            total_all_cells = "Missing"
            proc_comment = str(proc_comment) + '; cell count is not number, please fix'
            issues.add(sample_id)
            print(sample_id, "has invalid cell count:", total_cell_count)
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
                continue
            elif row['# of PBMC vials'] > 0 and (mat_row['Stype'] == 'Cells' or mat_row['Stype'] == 'Both'):
                add_to['Participant ID'].append(participant)
                add_to['Sample ID'].append(sample_id)
                add_to['Date'].append(row['Date'])
                add_to['Biospecimen_ID'].append(cells_id)
                add_to['Equipment_ID'].append(mat_row['Equipment_ID'])
                add_to['Equipment_Type'].append(mat_row['Equipment_Type'])
                add_to['Equipment_Calibration_Due_Date'].append('N/A')
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
                continue
            elif row['# of PBMC vials'] > 0 and (mat_row['Stype'] == 'Cells' or mat_row['Stype'] == 'Both'):
                add_to['Participant ID'].append(participant)
                add_to['Sample ID'].append(sample_id)
                add_to['Date'].append(row['Date'])
                add_to['Biospecimen_ID'].append(cells_id)
                cname = mat_row['Consumable_Name']
                add_to['Consumable_Name'].append(cname)
                _odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], cname, lot_log)
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
                continue
            elif row['# of PBMC vials'] > 0 and (mat_row['Stype'] == 'Cells' or mat_row['Stype'] == 'Both'):
                add_to['Participant ID'].append(participant)
                add_to['Sample ID'].append(sample_id)
                add_to['Date'].append(row['Date'])
                add_to['Biospecimen_ID'].append(cells_id)
                rname = mat_row['Reagent_Name']
                add_to['Reagent_Name'].append(rname)
                _odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], rname, lot_log)
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
            _odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], tname, lot_log)
            add_to['Aliquot_Tube_Type_Catalog_Number'].append(cat)
            add_to['Aliquot_Tube_Type_Lot_Number'].append(lot)
            add_to['Aliquot_Tube_Type_Expiration_Date'].append(exp)
            add_to['Comments'].append('')
        else:
            print("{} for {} has no serum".format(sample_id, participant))
        if not (type(vial_count) == str or (type(vial_count) == float and pd.isna(vial_count))):
            for i in range(min(int(vial_count), 2)):
                add_to['Participant ID'].append(participant)
                add_to['Sample ID'].append(sample_id)
                add_to['Date'].append(row['Date'])
                add_to['Biospecimen_ID'].append(cells_id)
                add_to['Aliquot_ID'].append("{}_{}".format(cells_id, i + 1))
                add_to['Aliquot_Volume'].append(1)
                add_to['Aliquot_Units'].append('mL')
                try:
                    if i + 1 == int(vial_count):
                        add_to['Aliquot_Concentration'].append("{:.0f}".format(last_aliq_live_cells))
                    else:
                        add_to['Aliquot_Concentration'].append("{:.0f}".format(aliq_live_cells))
                except:
                    print("Issue with aliquot for", sample_id)
                    add_to['Aliquot_Concentration'].append("{:.0f}".format(aliq_live_cells))
                add_to['Aliquot_Initials'].append(proc_inits)
                tname = 'CRYOTUBE 1.8ML'
                add_to['Aliquot_Tube_Type'].append(tname)
                _odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], tname, lot_log)
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
            add_to['Time in LN'].append(ln_time)
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(cohort)
            add_to['Visit_Number'].append(visit)
            add_to['Biospecimen_ID'].append(serum_id)
            add_to['Biospecimen_Type'].append('Serum')
            add_to['Biospecimen_Collection_Date_Duration_From_Index'].append(row['Days from Index'])
            add_to['Biospecimen_Processing_Batch_ID'].append('Please calculate using "Seronet Batch ID.xlsx"')
            add_to['Initial_Volume_of_Biospecimen'].append(sst_vol)
            add_to['Biospecimen_Collection_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            add_to['Biospecimen_Collector_Initials'].append(coll_inits)
            add_to['Biospecimen_Collection_Year'].append(row['Date'].year)
            tname = 'SST'
            add_to['Collection_Tube_Type'].append(tname)
            _odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], tname, lot_log)
            add_to['Collection_Tube_Type_Catalog_Number'].append(cat)
            add_to['Collection_Tube_Type_Lot_Number'].append(lot)
            add_to['Collection_Tube_Type_Expiration_Date'].append(exp)
            add_to['Storage_Time_at_2_8_Degrees_Celsius'].append('N/A')
            add_to['Storage_Start_Time_at_2-8_Initials'].append('N/A')
            add_to['Storage_End_Time_at_2-8_Initials'].append('N/A')
            add_to['Biospecimen_Processing_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            add_to['Biospecimen_Processor_Initials'].append(proc_inits)
            time_diff = time_diff_wrapper(coll_time, rec_time, "collection to receipt {}".format(sample_id))
            add_to['Biospecimen_Collection_to_Receipt_Duration'].append(time_diff)
            time_diff = time_diff_wrapper(rec_time, serum_freeze_time, "receipt to serum storage {}".format(sample_id))
            add_to['Biospecimen_Receipt_to_Storage_Duration'].append(time_diff)
            time_diff = time_diff_wrapper(coll_time, proc_time, "clot time {}".format(sample_id))
            if time_diff != 'Missing':
                time_diff = round(time_diff * 60., 2)
            add_to['RT_Serum_Clotting_Time'].append(time_diff)
            add_to['Centrifugation_Time'].append(10)
            for col in ['Live_Cells_Hemocytometer_Count', 'Total_Cells_Hemocytometer_Count', 'Viability_Hemocytometer_Count', 'Live_Cells_Automated_Count', 'Total_Cells_Automated_Count', 'Viability_Automated_Count', 'Storage_Time_in_Mr_Frosty']:
                add_to[col].append('N/A')
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
            add_to['Time in LN'].append(ln_time)
            add_to['Research_Participant_ID'].append(seronet_id)
            add_to['Cohort'].append(cohort)
            add_to['Visit_Number'].append(visit)
            add_to['Biospecimen_ID'].append(cells_id)
            add_to['Biospecimen_Type'].append('PBMC')
            add_to['Biospecimen_Collection_Date_Duration_From_Index'].append(row['Days from Index'])
            add_to['Biospecimen_Processing_Batch_ID'].append('Missing') # decide how to handle
            add_to['Initial_Volume_of_Biospecimen'].append(cpt_vol)
            add_to['Biospecimen_Collection_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            add_to['Biospecimen_Collector_Initials'].append(coll_inits)
            add_to['Biospecimen_Collection_Year'].append(row['Date'].year)
            tname = 'CPT'
            add_to['Collection_Tube_Type'].append(tname)
            _odate, lot, exp, cat = get_catalog_lot_exp(row['Date'], tname, lot_log)
            add_to['Collection_Tube_Type_Catalog_Number'].append(cat)
            add_to['Collection_Tube_Type_Lot_Number'].append(lot)
            add_to['Collection_Tube_Type_Expiration_Date'].append(exp)
            add_to['Storage_Time_at_2_8_Degrees_Celsius'].append('N/A')
            add_to['Storage_Start_Time_at_2-8_Initials'].append('N/A')
            add_to['Storage_End_Time_at_2-8_Initials'].append('N/A')
            add_to['Biospecimen_Processing_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            add_to['Biospecimen_Processor_Initials'].append(proc_inits)
            time_diff = time_diff_wrapper(coll_time, rec_time, "collection to receipt {}".format(sample_id))
            add_to['Biospecimen_Collection_to_Receipt_Duration'].append(time_diff)
            time_diff = time_diff_wrapper(rec_time, cell_freeze_time, "receipt to cell storage {}".format(sample_id))
            add_to['Biospecimen_Receipt_to_Storage_Duration'].append(time_diff)
            add_to['Centrifugation_Time'].append("N/A")
            add_to['RT_Serum_Clotting_Time'].append('N/A')
            add_to['Live_Cells_Hemocytometer_Count'].append('N/A')
            add_to['Total_Cells_Hemocytometer_Count'].append('N/A')
            add_to['Viability_Hemocytometer_Count'].append('N/A')
            add_to['Live_Cells_Automated_Count'].append(total_live_cells)
            add_to['Total_Cells_Automated_Count'].append(total_all_cells)
            add_to['Viability_Automated_Count'].append(viability)
            if row['Freezing Method'] != 'Mr. Frosty':
                add_to['Storage_Time_in_Mr_Frosty'].append('N/A')
            else:
                time_diff = time_diff_wrapper(cell_freeze_time, ln_time + pd.Timedelta(hours=24),"Time in Mr Frosty {}".format(sample_id))
                add_to['Storage_Time_in_Mr_Frosty'].append(time_diff)
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

    output = {}
    for sname, df2b in future_output.items():
        df = pd.DataFrame(df2b)
        output[sname] = df
    if not debug:
        with pd.ExcelWriter(util.proc_d4 + '{}.xlsx'.format(output_fname)) as writer:
            for sname, df in output.items():
                df[df['Date'].apply(lambda val: first_date <= val <= last_date)].to_excel(writer, sheet_name=sname, index=False)
        source.to_excel(util.script_output + 'SERONET_In_Window_Data_biospecimen_companion.xlsx', index=False)
        with open(util.script_output + 'trouble.csv', 'w+') as f:
            print("Sample ID", file=f)
            for sample in issues:
                print(sample, file=f)
    return output



if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Create processing information sheets for SERONET data submissions.')
    argParser.add_argument('-c', '--use_cache', action='store_true')
    argParser.add_argument('-o', '--output_file', action='store', default='tmp')
    argParser.add_argument('-s', '--start', action='store', type=pd.to_datetime, default=pd.to_datetime('1/1/2021'))
    argParser.add_argument('-e', '--end', action='store', type=pd.to_datetime, default=pd.to_datetime('12/31/2025'))
    argParser.add_argument('-x', '--override', action='store', type=pd.read_excel)
    argParser.add_argument('-d', '--debug', action='store_true')
    args = argParser.parse_args()
    if args.override is not None:
        make_ecrabs(args.override, args.start, args.end, args.output_file, args.debug)
        exit(0)
    if not args.use_cache:
        source = pull_from_source(args.debug)
    else:
        source = pd.read_excel(util.unfiltered)
    make_ecrabs(source, args.start, args.end, args.output_file, args.debug)
