from matplotlib.pyplot import box
import pandas as pd
import numpy as np
from datetime import date
import re
from copy import deepcopy
import os
import pickle
import argparse
from helpers import clean_sample_id

# Pull all file/folder locations from util.py
# Pull in sample data from dscf to check (query_dscf in helpers.py)


def filter_box(box_name, box_df=None):
    if box_df is None:
        return False
    return (box_name not in uploaded) and ((args.min_count <= box_df.set_index('Name').loc[box_name, 'Tube Count'] <= 81) or (box_name in full_des))

def timp_check(name):
    mit = "MIT" in name.upper()
    mars = "MARS" in name.upper()
    iris = "IRIS" in name.upper()
    titan = "TITAN" in name.upper()
    prio = "PRIORITY" in name.upper()
    return mit or mars or iris or titan or prio

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Make Seronet monthly sample report.')
    argParser.add_argument('-m', '--min_count', action='store', type=int, default=78)
    args = argParser.parse_args()
    home = '~/The Mount Sinai Hospital/'
    processing = home + 'Simon Lab - Processing Team/'
    inventory_boxes = pd.read_excel(processing + 'ACTIVE_FreezerPro Import Sheet (use this!!).xlsx', sheet_name=None)
    sample_types = ['Plasma', 'Serum', 'Pellet', 'Saliva', 'PBMC', 'HT', '4.5 mL Tube', 'All']
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
            box_number = int(name.split()[-1])
        except:
            print(name, "has no box number")
            continue
        if "Sample ID" not in sheet.columns:
            continue
        if re.search('PSP', name):
            team = 'PSP'
        elif timp_check(name):
            team = 'PVI ' + name.split()[0]
        else:
            team = 'PVI'
        sample_type = 'N/A'
        for val in sample_types:
            if re.search(str(val).upper().split()[0], name.upper()):
                sample_type = val
                break
        if sample_type == 'N/A':
            print(name, "has a box number but no valid sample type")
            continue
        box_kinds = []
        if re.search("LAB", name.upper()):
            box_kinds.append("Lab")
        if re.search("FF", name.upper()):
            box_kinds.append("FF")
        if len(box_kinds) == 0:
            if sample_type in ['PBMC', 'HT', '4.5 mL Tube']:
                box_kinds.append('')
            else:
                print(name, " is neither lab nor ff")
        box_sample_type = sample_type
        sheet = sheet.assign(sample_id=clean_sample_id).set_index('sample_id')
        for kind in box_kinds:
            if box_sample_type in ['PBMC', 'HT', '4.5 mL Tube']:
                box_name = "{} {} {}".format(team, box_sample_type, box_number)
            else:
                box_name = "{} {} {} {}".format(team, box_sample_type, kind, box_number)
            if box_name not in box_counts.keys():
                box_counts[box_name] = 0
            freezer = 'Annenberg 18'
            level1 = 'Freezer 1 (Eiffel Tower)'
            level2 = 'Shelf 2'
            if box_sample_type == 'Saliva' or box_sample_type == 'Pellet':
                level3 = '{} Lab/FF Rack'.format(box_sample_type)
            elif box_sample_type == 'PBMC':
                freezer = 'LN Tank #3'
                level1 = 'PBMC SUPER TEMPORARY HOLDING'
                level2 = ''
                level3 = ''
            else:
                level3 = '{} {} Rack'.format(box_sample_type, kind)
            for i, (sample_id, row) in enumerate(sheet.iterrows()):
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
                # if (row['Sample ID'], sample_type) in aliquot_counts.keys():
                #     aliquot = aliquot_counts[(row['Sample ID'], sample_type)]
                #     aliquot_counts[(row['Sample ID'], sample_type)] += 1
                # else:
                #     aliquot = 1
                #     aliquot_counts[(row['Sample ID'], sample_type)] = 2
                data = samples_data[sample_type]
                data['Name'].append(row['Sample ID'])
                data['Sample ID'].append(row['Sample ID'])
                data['Sample Type'].append(row['Sample Type'].strip())
                data['Freezer'].append(freezer)
                data['Level1'].append(level1)
                data['Level2'].append(level2)
                data['Level3'].append(level3)
                data['Box'].append(box_name)
                data['Position'].append(i + 1)
                data['ALIQUOT'].append(row['Sample ID'])
                if sample_type in ['Serum', 'Plasma']:
                    if box_sample_type == 'HT':
                        data['Heat treated?'].append('Yes')
                    else:
                        data['Heat treated?'].append('No')
                if sample_type == '4.5 mL Tube':
                    data['Serum or Plasma?'].append(row['Serum or Plasma?'])
                data = samples_data['All']
                data['Name'].append(row['Sample ID'])
                data['Sample ID'].append(row['Sample ID'])
                data['Sample Type'].append(row['Sample Type'].strip())
                data['Freezer'].append(freezer)
                data['Level1'].append(level1)
                data['Level2'].append(level2)
                data['Level3'].append(level3)
                data['Box'].append(box_name)
                data['Position'].append(i + 1)
                data['ALIQUOT'].append(row['Sample ID'])
                box_counts[box_name] += 1
    # if os.path.exists(processing + 'script_data/uploaded_boxes.pkl'):
        # with open(processing + 'script_data/uploaded_boxes.pkl', 'rb') as f:
            # uploaded = pickle.load(f)
    # else:
        # uploaded = set()
    uploaded = set()
    full_des = set(inventory_boxes['Full Boxes, DES']['Name'].to_numpy())
    box_data = {'Name': [], 'Tube Count': []}
    for box_name, tube_count in box_counts.items():
        box_data['Name'].append(box_name)
        box_data['Tube Count'].append(tube_count)
    box_df = pd.DataFrame(box_data)
    pd.DataFrame(samples_data['All']).to_excel(processing + 'inventory_in_progress.xlsx')
    uploading_boxes = box_df[box_df['Name'].apply(lambda val: filter_box(val, box_df=box_df))]
    with pd.ExcelWriter(processing + 'aggregate_inventory_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))) as writer:
        for sample_type, data in samples_data.items():
            if sample_type != 'All':
                df = pd.DataFrame(data)
                df = df[df['Box'].apply(lambda val: filter_box(val, box_df=box_df))]
                df.to_excel(writer, sheet_name='{}'.format(sample_type), index=False)
        box_df.to_excel(writer, sheet_name='Box Counts')
        uploading_boxes.to_excel(writer, sheet_name='Uploaded Box Counts')
    for _, row in box_df[~box_df['Name'].apply(lambda val: filter_box(val, box_df=box_df))].iterrows():
        print(row['Name'], 'with', row['Tube Count'], "tubes is partially inventoried and not being uploaded")
    if uploading_boxes.shape[0] > 0:
        print("Boxes to upload today:")
    for _, row in uploading_boxes.iterrows():
            print(row['Name'])
            uploaded.add(row['Name'])
    # with open(processing + 'script_data/uploaded_boxes.pkl', 'wb+') as f:
        # pickle.dump(uploaded, f)
