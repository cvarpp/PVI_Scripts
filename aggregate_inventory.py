from matplotlib.pyplot import box
import pandas as pd
import numpy as np
import util
from datetime import date
import re
from copy import deepcopy
import os
import argparse
from helpers import clean_sample_id

freezer_map = pd.read_excel(util.freezer_map, sheet_name='Racks in Freezers', header=0)    

# Pull all file/folder locations from util.py
# Pull in sample data from dscf to check (query_dscf in helpers.py)

def pos_convert(idx):
    '''Converts a 0-80 index to a box position'''
    return str((idx // 9) + 1) + "/" + "ABCDEFGHI"[idx % 9]

def filter_box(box_name, uploaded, args, box_df=None):
    '''
    Returns True if the box is not in the uploaded list and has a valid number of samples
    '''
    if box_df is None:
        return False
    return (box_name not in uploaded) and ((args.min_count <= box_df.set_index('Name').loc[box_name, 'Tube Count'] <= 81) or (box_name in full_des) or (box_df.set_index('Name').loc[box_name, 'Done?'] == "Yes"))
if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Aggregate inventory of each sample type for FP upload')
    argParser.add_argument('-m', '--min_count', action='store', type=int, default=81)
    args = argParser.parse_args()
    inventory_boxes = pd.read_excel("~/Downloads/New Import Sheet.xlsx", sheet_name=None)
    freezer_map = pd.read_excel(util.freezer_map, sheet_name='Racks in Freezers', header=0)
    racks_in_freezers = freezer_map.dropna(subset='Rack ID', axis=0).copy().reset_index()
    freezers_positions = pd.read_excel(util.freezer_map, sheet_name='FP Shelf_Rack Names', header=0, index_col='Concat')
    sample_types = ['Plasma', 'Serum', 'Pellet', 'Saliva', 'PBMC', 'HT', '4.5 mL Tube', 'NPS', 'All']
    data = {'Name': [], 'Sample ID': [], 'Sample Type': [],'Freezer': [],'Level1': [],'Level2': [],'Level3': [],'Box': [],'Position': [], 'ALIQUOT': []}
    samples_data = {st: deepcopy(data) for st in sample_types if st != 'HT'}
    samples_data['Serum']['Heat treated?'] = []
    samples_data['Plasma']['Heat treated?'] = []
    samples_data['4.5 mL Tube']['Serum or Plasma?'] = []
    samples_data['All'] = deepcopy(data)
    box_counts = {}
    aliquot_counts = {}
    completion = []
    warnings = {}
    #Maybe Dictionary?
    boxes_lost_name = []
    boxes_lost_reason = []

    rack_shelf={}
    rack_rack={}
    rack_freezer={}
    rack_floor={}
    rack_concat={}

    shelf_errorn = 0
    pos_errorn = 0
    farm_errorn = 0
    freezer_errorn = 0

    for n, item in enumerate(racks_in_freezers['Rack ID']):
        
        if racks_in_freezers['Shelf'][n] != racks_in_freezers['Shelf'][n]:
            racks_in_freezers['Shelf'][n] = int(900+shelf_errorn)
            shelf_errorn+=1
        
        if racks_in_freezers['Position'][n] != racks_in_freezers['Position'][n]:
            racks_in_freezers['Position'][n] = int(900+pos_errorn)
            pos_errorn+=1

        if racks_in_freezers['Farm'][n] != racks_in_freezers['Farm'][n]:
            racks_in_freezers['Farm'][n] = "Unknown "+str(farm_errorn)
            shelf_errorn+=1

        if racks_in_freezers['Freezer'][n] != racks_in_freezers['Freezer'][n]:
            racks_in_freezers['Freezer'][n] = "Unknown "+str(freezer_errorn)
            pos_errorn+=1

        rack_shelf.update({item:racks_in_freezers['Shelf'][n]})
        rack_rack.update({item:racks_in_freezers['Position'][n]})
        rack_freezer.update({item:racks_in_freezers['Freezer'][n]})
        rack_floor.update({item:racks_in_freezers['Farm'][n]})
        print({item:str(racks_in_freezers['Farm'][n])+':'+str(racks_in_freezers['Freezer'][n])+':'+str(int(racks_in_freezers['Shelf'][n]))+':'+str(int(racks_in_freezers['Position'][n]))})
        rack_concat.update({item:str(racks_in_freezers['Farm'][n])+':'+str(racks_in_freezers['Freezer'][n])+':'+str(int(racks_in_freezers['Shelf'][n]))+':'+str(int(racks_in_freezers['Position'][n]))})
        print(rack_concat[item])
        

    #59 shelf errors
    #70 position errors
    # 1 freezer error
    # 1 farm error
    # 4/14/25

    for name, sheet in inventory_boxes.items():
        try:
            box_number = int(name.rstrip('LabFF_)( ').split()[-1])
        except:
            print(name, "has no box number")
            boxes_lost_name.append(name)
            boxes_lost_reason.append("Has No Box Number")
            continue

        try:
            sheet['Box Done?'][0] == 1.0
        except:
            print(name, ': No Check Box')
            boxes_lost_name.append(name)
            boxes_lost_reason.append("Has No Check Box")
            continue
        
        rack_number = sheet['Rack Number'][0]
        
        if rack_number != rack_number:
            
            #Add temp upload functionality?
            #Temp_marker = 1

            print(name, ": Rack number Not Filled in")
            boxes_lost_name.append(name)
            boxes_lost_reason.append("Rack Number Empty")
            continue

        if rack_number not in rack_shelf:
            print(name, ": Rack number Not in Inventory")
            boxes_lost_name.append(name)
            boxes_lost_reason.append("Rack Number Provided Not in Inventory")
            continue

        if "Sample ID" not in sheet.columns:
            print(name, ": No Sample ID column in sheet")
            boxes_lost_name.append(name)
            boxes_lost_reason.append("No Sample ID Column")
            continue

        if re.search('APOLLO RESEARCH \d+', name) or re.search('APOLLO NIH \d+', name):
            team = 'APOLLO'
        
        elif re.search('PSP', name):
            team = 'PSP'
        
        elif re.match("MIT|MARS|IRIS|TITAN|PRIORITY", name.upper()) is not None:
            team = 'PVI ' + name.split()[0]
        
        else:
            team = 'PVI'

        sample_type = 'N/A'

        if team == 'APOLLO':
            sample_type = 'Serum'
        else:
            for val in sample_types:
                if re.search(str(val).upper().split()[0], name.upper()):
                    sample_type = val
                    break

        if sample_type == 'N/A':
            print(name, "has a box number but no valid sample type")
            boxes_lost_name.append(name)
            boxes_lost_reason.append("No Valid Sample ID value")
            continue

        box_kinds = re.findall('RESEARCH|NIH|Lab|FF', name)
        if len(box_kinds) == 0:
            if sample_type in ['PBMC', 'HT', '4.5 mL Tube']:
                box_kinds.append('')
            else:
                print(name, " is neither lab nor ff. Not uploaded")
        
        box_sample_type = sample_type
        sheet = sheet.assign(sample_id=clean_sample_id).set_index('sample_id')
        for kind in box_kinds:
            if team == 'APOLLO':
                box_name = f"{team} {kind} {box_number}"
            elif box_sample_type in ['PBMC', 'HT', '4.5 mL Tube']:
                box_name = f"{team} {box_sample_type} {box_number}"
            else:
                box_name = f"{team} {box_sample_type} {kind} {box_number}"
            if box_name not in box_counts.keys():
                box_counts[box_name] = 0

            freezer_index = rack_concat[rack_number]

            FP_pos_logic = (freezers_positions.loc[freezer_index,'FP_Freezer'] == freezers_positions.loc[freezer_index,'FP_Freezer'] and
              freezers_positions.loc[freezer_index,'FP_Level1'] == freezers_positions.loc[freezer_index,'FP_Level1'] and
                freezers_positions.loc[freezer_index,'FP_Level2'] == freezers_positions.loc[freezer_index,'FP_Level2'] and
                    freezers_positions.loc[freezer_index,'FP_Level3'] == freezers_positions.loc[freezer_index,'FP_Level3'])
            
            Local_pos_logic = (freezers_positions.loc[freezer_index,'Farm'] == freezers_positions.loc[freezer_index,'Farm'] and
              freezers_positions.loc[freezer_index,'Freezer'] == freezers_positions.loc[freezer_index,'Freezer'] and
                freezers_positions.loc[freezer_index,'Shelf'] == freezers_positions.loc[freezer_index,'Shelf'] and
                    freezers_positions.loc[freezer_index,'Position'] == freezers_positions.loc[freezer_index,'Position'])

            if freezer_index in freezers_positions.index and FP_pos_logic == True:
                freezer = freezers_positions.loc[freezer_index,'FP_Freezer']
                level1 = freezers_positions.loc[freezer_index,'FP_Level1']
                level2 = freezers_positions.loc[freezer_index,'FP_Level2']
                level3 = freezers_positions.loc[freezer_index,'FP_Level3']

            elif freezer_index in freezers_positions.index and Local_pos_logic == True:
                print("Warning!: Official Freezer Pro positions not given")
                print("Default Document Positions used! consult with file prior to upload")
                freezer = freezers_positions.loc[freezer_index,'Farm']
                level1 = freezers_positions.loc[freezer_index,'Freezer']
                level2 = "shelf " + str(int(freezers_positions.loc[freezer_index,'Shelf']))
                level3 = "rack " + str(int(freezers_positions.loc[freezer_index,'Position']))
                warnings.update({name:"Official Freezer Pro positions not given default document positions used!"})
            
            #elif Temp_marker == 1:
                # freezer = 'Temporary Holding'
                # level1 = sample_type
                # level2 = 'Shelf1'
                # level3 = f'{sample_type} Temporary Holding'

            else:
                freezer = 'Annenberg 18'
                level1 = 'Freezer 1 (Eiffel Tower)'
                level2 = 'Shelf {}'.format(int(rack_shelf[rack_number]))
                level3 = 'Rack #{}'.format(int(rack_number))

                if box_sample_type == 'NPS':
                    freezer = 'Temporary PSP NPS'
                    level1 = 'freezer_nps'
                    if rack_number=='missing':
                        level2 = 'Shelf {}'.format(int(rack_shelf[rack_number]))
                    else:
                        pass
                    level3 = 'Rack #{}'.format(int(rack_number))
                elif box_sample_type == 'PBMC':
                    freezer = 'LN Tank #3'
                    level1 = 'PBMC SUPER TEMPORARY HOLDING'
                else:
                    level3 = 'Rack #{}'.format(rack_number)
            
            
            for idx, (sample_id, row) in enumerate(sheet.iterrows()):
                if re.search("[1-9A-Z][0-9]{4,5}", sample_id) is None:
                    continue
                sample_type = 'N/A'
                for val in sample_types:
                    try:
                        if re.search(str(val).upper().split()[0], str(row['Sample Type']).upper()):
                            sample_type = val
                            break
                    except:
                        print("Box Name: ",box_name)
                        exit(1)
                if sample_type == 'N/A':
                    print(row['Sample ID'], 'in box', box_name, 'has no sample type specified. (', row['Sample Type'], ')')
                    continue
                for sname in [sample_type, 'All']:
                    data = samples_data[sname]
                    data['Name'].append(sample_id)
                    data['Sample ID'].append(sample_id)
                    data['Sample Type'].append(row['Sample Type'].strip())
                    data['Freezer'].append(freezer)
                    data['Level1'].append(level1)
                    data['Level2'].append(level2)
                    data['Level3'].append(level3)
                    data['Box'].append(box_name)
                    data['Position'].append(pos_convert(idx))
                    data['ALIQUOT'].append(sample_id)
                    if sname in ['Serum', 'Plasma']:
                        if box_sample_type == 'HT':
                            data['Heat treated?'].append('Yes')
                        else:
                            data['Heat treated?'].append('No')
                    if sname == '4.5 mL Tube':
                        data['Serum or Plasma?'].append(row['Serum or Plasma?'])
                box_counts[box_name] += 1
                
                if (box_name not in completion) and (sheet['Box Done?'][0] == 1.0):
                        completion.append(box_name)
#%%
    uploaded = set()
    full_des = set(inventory_boxes['Full Boxes, DES']['Name'].to_numpy())
    box_data = {'Name': [], 'Tube Count': [], 'Done?': []}
    for (box_name, tube_count) in box_counts.items():
        box_data['Name'].append(box_name)
        box_data['Tube Count'].append(tube_count)
        if box_name in completion:
            box_data['Done?'].append("Yes")
        else:
            box_data['Done?'].append("No")

    #%%

    box_df = pd.DataFrame(box_data)
    lost_df = pd.DataFrame.from_dict({"Box Name":boxes_lost_name, "Reason Dropped":boxes_lost_reason})
    print(box_df)
    pd.DataFrame(samples_data['All']).to_excel(util.proc + 'inventory_in_progress.xlsx')
    uploading_boxes = box_df[box_df['Name'].apply(lambda val: filter_box(val, uploaded, args, box_df=box_df))]
    with pd.ExcelWriter(util.proc + 'aggregate_inventory_test_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))) as writer:
        for sample_type, data in samples_data.items():
            if sample_type != 'All':
                df = pd.DataFrame(data)
                df = df[df['Box'].apply(lambda val: filter_box(val, uploaded, args, box_df=box_df))]
                df.to_excel(writer, sheet_name='{}'.format(sample_type), index=False)
        box_df.to_excel(writer, sheet_name='Box Counts')
        uploading_boxes.to_excel(writer, sheet_name='Uploaded Box Counts')
        lost_df.to_excel(writer, sheet_name="Unprocessed boxes")
    for _, row in box_df[~box_df['Name'].apply(lambda val: filter_box(val, uploaded, args, box_df=box_df))].iterrows():
        print(row['Name'], 'with', row['Tube Count'], "tubes is partially inventoried and not marked as completed")
    if uploading_boxes.shape[0] > 0:
        print("Boxes to upload today:")
    for _, row in uploading_boxes.iterrows():
            print(row['Name'])
            uploaded.add(row['Name'])
    for name, item in warnings.items():
        print(name, " ", item)
