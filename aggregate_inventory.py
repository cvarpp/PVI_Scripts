from matplotlib.pyplot import box
import pandas as pd
import numpy as np
from datetime import date
import re
from copy import deepcopy
import os
import pickle
import argparse

# Pull all file/folder locations from util.py
# Pull in sample data from dscf to check (query_dscf in helpers.py)


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
    sample_types = ['Plasma', 'Serum', 'Pellet', 'Saliva', 'PBMC', 'HT', '4.5', 'All']
    data = {'Name': [], 'Sample ID': [], 'Sample Type': [],'Freezer': [],'Level1': [],'Level2': [],'Level3': [],'Box': [],'Position': [], 'ALIQUOT': []}
    samples_data = {st: deepcopy(data) for st in sample_types if st != 'HT'}
    samples_data['Serum']['Heat treated?'] = []
    samples_data['Plasma']['Heat treated?'] = []
    samples_data['All'] = deepcopy(data)
    box_counts = {}
    aliquot_counts = {}
    for name, sheet in inventory_boxes.items():
        if not re.search('[0-9]{2,3}', name):
            print(name, "has no box number")
            continue
        box_number = int(re.search('[0-9]{2,3}', name)[0])
        if re.search('PSP', name):
            team = 'PSP'
        elif timp_check(name):
            team = 'PVI ' + name.split()[0]
        else:
            team = 'PVI'
        sample_type = 'N/A'
        for val in sample_types:
            if re.search(val.upper(), name.upper()):
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
            if sample_type == 'PBMC' or sample_type == 'HT' or sample_type == '4.5':
                box_kinds.append('')
            else:
                print(name, " is neither lab nor ff")
        box_sample_type = sample_type
        sheet['Sample ID'] = sheet['Sample ID'].astype(str).str.upper().str.strip()
        for kind in box_kinds:
            if box_sample_type == 'PBMC' or box_sample_type == 'HT':
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
            for i, row in sheet.iterrows():
                try:
                    sample_id = row['Sample ID']
                except:
                    print(name, "has no Sample ID column")
                    print("Fatal Error! Exiting...")
                    exit(1)
                if len(str(row['Sample ID'])) < 5:
                    continue
                sample_type = 'N/A'
                for val in sample_types:
                    try:
                        if re.search(val.upper(), str(row['Sample Type']).upper()):
                            sample_type = val
                            break
                    except:
                        print(box_name)
                        exit(1)
                if sample_type == 'N/A':
                    print(row['Sample ID'], 'in box', box_name, 'has no sample type specified. (', row['Sample Type'], ')')
                    continue
                if (row['Sample ID'], sample_type) in aliquot_counts.keys():
                    aliquot = aliquot_counts[(row['Sample ID'], sample_type)]
                    aliquot_counts[(row['Sample ID'], sample_type)] += 1
                else:
                    aliquot = 1
                    aliquot_counts[(row['Sample ID'], sample_type)] = 2
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
    writer = pd.ExcelWriter(processing + 'aggregate_inventory_{}.xlsx'.format(date.today().strftime("%m.%d.%y")))
    box_data = {'Name': [], 'Tube Count': []}
    if os.path.exists(processing + 'script_data/uploaded_boxes.pkl'):
        with open(processing + 'script_data/uploaded_boxes.pkl', 'rb') as f:
            uploaded = pickle.load(f)
    else:
        uploaded = set()
    full_des = set(inventory_boxes['Full Boxes, DES']['Name'].to_numpy())
    def filter_box(box_name):
        return (box_name not in uploaded) and ((args.min_count <= box_df.loc[box_name, 'Tube Count'] <= 81) or (box_name in full_des))
    for box_name, tube_count in box_counts.items():
        box_data['Name'].append(box_name)
        box_data['Tube Count'].append(tube_count)
        if tube_count < args.min_count or tube_count > 81:
            print(box_name, ',', tube_count)
    box_df = pd.DataFrame(box_data).set_index('Name')
    for sample_type, data in samples_data.items():
        df = pd.DataFrame(data)
        if sample_type != 'All':
            df = df[df['Box'].apply(filter_box)]
            df.to_excel(writer, sheet_name='{}'.format(sample_type), index=False)
        else:
            df.to_excel(processing + 'inventory_in_progress.xlsx')
    box_df.to_excel(writer, sheet_name='Box Counts')
    box_df[((box_df['Tube Count'].apply(lambda val: args.min_count <= val <= 81)) | box_df.apply(lambda row: row.name in full_des, axis=1)) & (box_df.apply(lambda row: row.name not in uploaded, axis=1))].to_excel(writer, sheet_name='Uploaded Box Counts')
    print("Boxes to upload today:")
    for box_name, tube_count in box_counts.items():
        if filter_box(box_name):
            if box_name not in uploaded:
                print(box_name)
                uploaded.add(box_name)
    with open(processing + 'script_data/uploaded_boxes.pkl', 'wb+') as f:
        pickle.dump(uploaded, f)
    writer.save()
