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
    return (box_name not in uploaded) and ((args.min_count <= box_df.set_index('Name').loc[box_name, 'Tube Count'] <= 81) or (box_name in full_des))

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Aggregate inventory of each sample type for FP upload')
    argParser.add_argument('-m', '--min_count', action='store', type=int, default=78)
    args = argParser.parse_args()
    inventory_boxes = pd.read_excel(util.inventory_input, sheet_name=None)
    sample_types = ['Plasma', 'Serum', 'Pellet', 'Saliva', 'PBMC', 'HT', '4.5 mL Tube', 'NPS', 'All']
    data = {'Name': [], 'Sample ID': [], 'Sample Type': [],'Freezer': [],'Level1': [],'Level2': [],'Level3': [],'Box': [],'Position': [], 'ALIQUOT': []}
    samples_data = {st: deepcopy(data) for st in sample_types if st != 'HT'}
    samples_data['Serum']['Heat treated?'] = []
    samples_data['Plasma']['Heat treated?'] = []
    samples_data['4.5 mL Tube']['Serum or Plasma?'] = []
    samples_data['All'] = deepcopy(data)
    box_counts = {}
    aliquot_counts = {}

    for name, sheet in inventory_boxes.items():
        try:
            box_number = int(name.rstrip('LabFF_ ').split()[-1])
        except:
            print(name, "has no box number")
            continue
        if "Sample ID" not in sheet.columns:
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
        elif team == 'PSP':
            sample_type = 'NPS'
        else:
            for val in sample_types:
                if re.search(str(val).upper().split()[0], name.upper()):
                    sample_type = val
                    break
        if sample_type == 'N/A':
            print(name, "has a box number but no valid sample type")
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
            if team in ['APOLLO', 'PSP']:
                box_name = f"{team} {kind} {box_number}"
            elif box_sample_type in ['PBMC', 'HT', '4.5 mL Tube']:
                box_name = f"{team} {box_sample_type} {box_number}"
            else:
                box_name = f"{team} {box_sample_type} {kind} {box_number}"
            if box_name not in box_counts.keys():
                box_counts[box_name] = 0
            freezer = 'Annenberg 18'
            level1 = 'Freezer 1 (Eiffel Tower)'
            level2 = 'Shelf 2'
            if team == 'APOLLO':
                level3 = 'APOLLO Rack'
            elif box_sample_type == 'NPS':
                freezer = 'Temporary PSP NPS'
                level1 = 'freezer_nps'
                level2 = 'shelf_nps'
                level3 = 'rack_nps'
            elif box_sample_type == 'Saliva' or box_sample_type == 'Pellet':
                level3 = '{} Lab/FF Rack'.format(box_sample_type)
            elif box_sample_type == 'PBMC':
                freezer = 'LN Tank #3'
                level1 = 'PBMC SUPER TEMPORARY HOLDING'
                level2 = ''
                level3 = ''
            else:
                level3 = '{} {} Rack'.format(box_sample_type, kind)
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
                        print(box_name)
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
    uploaded = set()
    full_des = set(inventory_boxes['Full Boxes, DES']['Name'].to_numpy())
    box_data = {'Name': [], 'Tube Count': []}
    for box_name, tube_count in box_counts.items():
        box_data['Name'].append(box_name)
        box_data['Tube Count'].append(tube_count)
    box_df = pd.DataFrame(box_data)
    pd.DataFrame(samples_data['All']).to_excel(util.proc + 'inventory_in_progress.xlsx')
    uploading_boxes = box_df[box_df['Name'].apply(lambda val: filter_box(val, uploaded, args, box_df=box_df))]
    with pd.ExcelWriter(util.proc + 'aggregate_inventory_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))) as writer:
        for sample_type, data in samples_data.items():
            if sample_type != 'All':
                df = pd.DataFrame(data)
                df = df[df['Box'].apply(lambda val: filter_box(val, uploaded, args, box_df=box_df))]
                df.to_excel(writer, sheet_name='{}'.format(sample_type), index=False)
        box_df.to_excel(writer, sheet_name='Box Counts')
        uploading_boxes.to_excel(writer, sheet_name='Uploaded Box Counts')
    for _, row in box_df[~box_df['Name'].apply(lambda val: filter_box(val, uploaded, args, box_df=box_df))].iterrows():
        print(row['Name'], 'with', row['Tube Count'], "tubes is partially inventoried and not being uploaded")
    if uploading_boxes.shape[0] > 0:
        print("Boxes to upload today:")
    for _, row in uploading_boxes.iterrows():
            print(row['Name'])
            uploaded.add(row['Name'])
