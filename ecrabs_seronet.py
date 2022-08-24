import pandas as pd
import numpy as np
from datetime import date
import datetime
from dateutil import parser
import sys
import argparse
import util
from seronet_data import filter_windows
from d4_all_data import pull_from_source

def process_lots():
    lots = pd.read_excel(util.dscf, sheet_name='Lot # Sheet')
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

def make_ecrabs(source, first_date='1/1/2021', last_date='12/31/2025', output_fname='tmp'):
    issues = set()
    first_date = parser.parse(first_date).date()
    last_date = parser.parse(last_date).date()
    equip_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Equipment_ID', 'Equipment_Type', 'Equipment_Calibration_Due_Date', 'Comments']
    consum_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Consumable_Name', 'Consumable_Catalog_Number', 'Consumable_Lot_Number', 'Consumable_Expiration_Date', 'Comments']
    reag_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Reagent_Name', 'Reagent_Catalog_Number', 'Reagent_Lot_Number', 'Reagent_Expiration_Date', 'Comments']
    aliq_cols = ['Participant ID', 'Sample ID', 'Date', 'Biospecimen_ID', 'Aliquot_ID', 'Aliquot_Volume', 'Aliquot_Units', 'Aliquot_Concentration', 'Aliquot_Initials', 'Aliquot_Tube_Type', 'Aliquot_Tube_Type_Catalog_Number', 'Aliquot_Tube_Type_Lot_Number', 'Aliquot_Tube_Type_Expiration_Date', 'Comments']
    biospec_cols = ['Participant ID', 'Sample ID', 'Date', 'Proc Comments', 'Time Collected', 'Time Received', 'Time Processed', 'Time Serum Frozen', 'Time Cells Frozen', 'Research_Participant_ID', 'Cohort', 'Visit_Number', 'Biospecimen_ID', 'Biospecimen_Type', 'Biospecimen_Collection_Date_Duration_From_Index', 'Biospecimen_Processing_Batch_ID', 'Initial_Volume_of_Biospecimen', 'Biospecimen_Collection_Company_Clinic', 'Biospecimen_Collector_Initials', 'Biospecimen_Collection_Year', 'Collection_Tube_Type', 'Collection_Tube_Type_Catalog_Number', 'Collection_Tube_Type_Lot_Number', 'Collection_Tube_Type_Expiration_Date', 'Storage_Time_at_2_8_Degrees_Celsius', 'Storage_Start_Time_at_2-8_Initials', 'Storage_End_Time_at_2-8_Initials', 'Biospecimen_Processing_Company_Clinic', 'Biospecimen_Processor_Initials', 'Biospecimen_Collection_to_Receipt_Duration', 'Biospecimen_Receipt_to_Storage_Duration', 'Centrifugation_Time', 'RT_Serum_Clotting_Time', 'Live_Cells_Hemocytometer_Count', 'Total_Cells_Hemocytometer_Count', 'Viability_Hemocytometer_Count', 'Live_Cells_Automated_Count', 'Total_Cells_Automated_Count', 'Viability_Automated_Count', 'Storage_Time_in_Mr_Frosty', 'Comments']
    ship_cols = ['Participant ID', 'Sample ID', 'Date', 'Study ID', 'Current Label', 'Material Type', 'Volume', 'Volume Unit', 'Volume Estimate', 'Vial Type', 'Vial Warnings', 'Vial Modifiers']

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
    for _, row in source.iterrows():
        participant = row['Participant ID']
        sample_id = row['Sample ID']
        seronet_id = row['Seronet ID']
        cohort = row['Cohort']
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
        coll_inits = row['coll_inits']
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
            elif vial_count == 0:
                aliq_cells = 0
            elif type(vial_count) != str:
                aliq_cells = cell_count / vial_count
            else:
                aliq_cells = cell_count
            live_cells = cell_count * (10**6)
            try:
                all_cells = live_cells * 100 / viability
            except:
                all_cells = "Viability needs fixing"
                if vial_count != 0:
                    issues.add(sample_id)
        except Exception as e:
            print(e)
            aliq_cells = '???'
            live_cells = cell_count
            all_cells = '???'
            proc_comment = str(proc_comment) + '; cell count is not number, please fix'
            issues.add(sample_id)
            print(sample_id, "has invalid cell count:", cell_count)
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
            add_to['Cohort'].append(cohort)
            add_to['Visit_Number'].append(visit)
            add_to['Biospecimen_ID'].append(serum_id)
            add_to['Biospecimen_Type'].append('Serum')
            add_to['Biospecimen_Collection_Date_Duration_From_Index'].append(row['Days from Index'])
            add_to['Biospecimen_Processing_Batch_ID'].append('???') # decide how to handle
            add_to['Initial_Volume_of_Biospecimen'].append(sst_vol)
            add_to['Biospecimen_Collection_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            add_to['Biospecimen_Collector_Initials'].append(coll_inits)
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
                add_to['Biospecimen_Collection_to_Receipt_Duration'].append(
                    (datetime.datetime.combine(date.min, parser.parse(rec_time).time()) -
                    datetime.datetime.combine(date.min, parser.parse(coll_time).time())) / datetime.timedelta(hours=1))
            except Exception as e:
                if not (pd.isna(rec_time) or pd.isna(coll_time)):
                    print("col to rec", e)
                    print(coll_time, type(coll_time))
                    print(rec_time, type(rec_time))
                    print()
                issues.add(sample_id)
                add_to['Biospecimen_Collection_to_Receipt_Duration'].append('???')
            try:
                add_to['Biospecimen_Receipt_to_Storage_Duration'].append(
                    (datetime.datetime.combine(date.min, parser.parse(serum_freeze_time).time()) -
                    datetime.datetime.combine(date.min, parser.parse(rec_time).time())) / datetime.timedelta(hours=1))
            except Exception as e:
                if not (pd.isna(rec_time) or pd.isna(serum_freeze_time)):
                    print("rec to store", e)
                    print(serum_freeze_time, type(serum_freeze_time))
                    print(rec_time, type(rec_time))
                    print()
                issues.add(sample_id)
                add_to['Biospecimen_Receipt_to_Storage_Duration'].append('???')
            add_to['Centrifugation_Time'].append(20)
            try:
                add_to['RT_Serum_Clotting_Time'].append(
                    (datetime.datetime.combine(date.min, parser.parse(proc_time).time()) -
                    datetime.datetime.combine(date.min, parser.parse(coll_time).time())) / datetime.timedelta(hours=1))
            except Exception as e:
                if not (pd.isna(proc_time) or pd.isna(coll_time)):
                    print("clot", e)
                    print(proc_time, type(proc_time))
                    print(coll_time, type(coll_time))
                    print()
                issues.add(sample_id)
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
            add_to['Cohort'].append(cohort)
            add_to['Visit_Number'].append(visit)
            add_to['Biospecimen_ID'].append(cells_id)
            add_to['Biospecimen_Type'].append('PBMC')
            add_to['Biospecimen_Collection_Date_Duration_From_Index'].append(row['Days from Index'])
            add_to['Biospecimen_Processing_Batch_ID'].append('???') # decide how to handle
            add_to['Initial_Volume_of_Biospecimen'].append(cpt_vol)
            add_to['Biospecimen_Collection_Company_Clinic'].append('Icahn School of Medicine at Mount Sinai')
            add_to['Biospecimen_Collector_Initials'].append(coll_inits)
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
                add_to['Biospecimen_Collection_to_Receipt_Duration'].append(
                    (datetime.datetime.combine(date.min, parser.parse(rec_time).time()) -
                    datetime.datetime.combine(date.min, parser.parse(coll_time).time())) / datetime.timedelta(hours=1))
            except Exception as e:
                if not (pd.isna(rec_time) or pd.isna(coll_time)):
                    print("col to rec", e)
                    print(coll_time, type(coll_time))
                    print(rec_time, type(rec_time))
                    print()
                issues.add(sample_id)
                add_to['Biospecimen_Collection_to_Receipt_Duration'].append('???')
            try:
                add_to['Biospecimen_Receipt_to_Storage_Duration'].append(
                    (datetime.datetime.combine(date.min, parser.parse(cell_freeze_time).time()) -
                    datetime.datetime.combine(date.min, parser.parse(rec_time).time())) / datetime.timedelta(hours=1))
            except Exception as e:
                if not (pd.isna(rec_time) or pd.isna(cell_freeze_time)):
                    print("rec to store", e)
                    print(cell_freeze_time, type(cell_freeze_time))
                    print(rec_time, type(rec_time))
                    print()
                issues.add(sample_id)
                add_to['Biospecimen_Receipt_to_Storage_Duration'].append('???')
            add_to['Centrifugation_Time'].append(40)
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
                    issues.add(sample_id)
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

    writer = pd.ExcelWriter(util.d4 + '{}.xlsx'.format(output_fname))
    for sname, df2b in future_output.items():
        df = pd.DataFrame(df2b)
        df[df['Date'].apply(lambda val: first_date <= val <= last_date)].to_excel(writer, sheet_name=sname, index=False)
    writer.save()
    source.to_excel(util.script_output + 'SERONET_In_Window_Data_biospecimen_companion.xlsx', index=False)
    with open(util.script_output + 'trouble.csv', 'w+') as f:
        print("Sample ID", file=f)
        for sample in issues:
            print(sample, file=f)



if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Make Seronet monthly sample report.')
    argParser.add_argument('-u', '--update', action='store_true')
    argParser.add_argument('-f', '--filter', action='store_true')
    argParser.add_argument('-o', '--output_file', action='store', default='tmp')
    argParser.add_argument('-s', '--start', action='store', default='1/1/2021')
    argParser.add_argument('-e', '--end', action='store', default='12/31/2025')
    args = argParser.parse_args()
    if args.update:
        source = filter_windows(pull_from_source())
    else:
        if args.filter:
            source = filter_windows(pd.read_excel(util.unfiltered))
        else:
            source = pd.read_excel(util.filtered)
    make_ecrabs(source, args.start, args.end, args.output_file)
