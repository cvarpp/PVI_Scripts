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

def filter_box(box_name, uploaded, box_df=None):
    '''
    Returns True if the box is not in the uploaded list and has a valid number of samples
    '''
    if box_df is None:
        return False
    return (box_name not in uploaded) and (box_df.set_index('Name').loc[box_name, 'Done?'] == "Yes")

class Box:
    def __init__(self, name, sheet, freezer_map, boxes_lost):
        self.valid = False
        self.name = name
        self.sheet = sheet
        try:
            self.box_number = int(name.rstrip('LabFF_)( ').split()[-1])
        except:
            boxes_lost[name] = "Has No Box Number"
            return
        try:
            self.box_done = sheet['Box Done?'][0] == 1.0
        except:
            boxes_lost[name] = "Has No Check Box"
            return
        try:
            self.rack_number = int(sheet['Rack Number'][0])
        except:
            boxes_lost[name] = "Rack Number Missing"
            return
        if self.rack_number not in freezer_map['Rack ID']:
            boxes_lost[name] = "Rack Number Provided Not in Inventory"
            return
        if "Sample ID" not in sheet.columns:
            boxes_lost[name] = "No Sample ID Column"
            return
        if re.search('APOLLO', name):
            self.team = 'APOLLO'
        elif re.search('PSP', name):
            self.team = 'PSP'
        elif re.match("MIT|MARS|IRIS|TITAN|PRIORITY", name.upper()) is not None:
            self.team = 'PVI ' + name.split()[0]
        else:
            self.team = 'PVI'
        self.sample_type = None
        if self.team == 'APOLLO':
            self.sample_type = 'Serum'
        else:
            for val in sample_types:
                if re.search(str(val).upper().split()[0], name.upper()):
                    self.sample_type = val
                    break
        if self.sample_type is None:
            boxes_lost[name] = "No Valid Sample ID value"
            return
        self.box_kinds = re.findall('RESEARCH|NIH|Lab|FF', name)
        if len(self.box_kinds) == 0:
            if sample_type in ['PBMC', 'HT', '4.5 mL Tube']:
                self.box_kinds.append('')
            else:
                boxes_lost[name] = "Neither lab nor FF"
                return
        if not self.box_done:
            boxes_lost[name] = 'Not marked complete'
        self.valid = True


if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Aggregate inventory of each sample type for FP upload')
    argParser.add_argument('-m', '--min_count', action='store', type=int, default=81)
    args = argParser.parse_args()
    inventory_boxes = pd.read_excel(util.inventory_input, sheet_name=None)
    freezer_map = pd.read_excel(util.freezer_map, sheet_name='Racks in Freezers', header=0)
    racks_in_freezers = freezer_map.dropna(subset='Rack ID', axis=0).copy().set_index('Rack ID')
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
    boxes_lost = {}
    typeless_samples = []

    for name, sheet in inventory_boxes.items():
        boxx = Box(name, sheet, freezer_map, boxes_lost)
        if not boxx.valid or not boxx.box_done:
            continue
        sheet = sheet.assign(sample_id=clean_sample_id).set_index('sample_id')
        for kind in boxx.box_kinds:
            if boxx.team == 'APOLLO':
                box_name = f"{boxx.team} {kind} {boxx.box_number}"
            elif boxx.sample_type in ['PBMC', 'HT', '4.5 mL Tube']:
                box_name = f"{boxx.team} {boxx.sample_type} {boxx.box_number}"
            else:
                box_name = f"{boxx.team} {boxx.sample_type} {kind} {boxx.box_number}"
            box_counts[box_name] = 0
            rack_info = racks_in_freezers.loc[boxx.rack_number, :]
            if boxx.sample_type == 'PBMC':
                freezer = 'LN Tank #3'
                level1 = 'PBMC SUPER TEMPORARY HOLDING'
                level2 = ''
                level3 = ''
            else:
                freezer = str(rack_info['Farm']) + " New"
                level1 = str(rack_info['Freezer'])
                level2 = "Shelf " + int(rack_info['Shelf'])
                level3 = "Rack " + int(boxx.rack_number)
            for idx, (sample_id, row) in enumerate(sheet.iterrows()):
                sample_id = str(sample_id).strip().upper()
                if re.search("[1-9A-Z][0-9]{4,5}", sample_id) is None:
                    continue
                sample_type = 'N/A'
                for val in sample_types:
                    if re.search(str(val).upper().split()[0], str(row['Sample Type']).upper()):
                        sample_type = val
                        break
                if sample_type == 'N/A':
                    typeless_samples.append((sample_id, box_name))
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
                        if boxx.sample_type == 'HT':
                            data['Heat treated?'].append('Yes')
                        else:
                            data['Heat treated?'].append('No')
                    if sname == '4.5 mL Tube':
                        data['Serum or Plasma?'].append(row['Serum or Plasma?'])
                box_counts[box_name] += 1
            completion.append(box_name)
    uploaded = set()
    box_data = {'Name': [], 'Tube Count': []}
    for (box_name, tube_count) in box_counts.items():
        box_data['Name'].append(box_name)
        box_data['Tube Count'].append(tube_count)
    box_df = pd.DataFrame(box_data)
    lost_df = pd.DataFrame.from_dict({"Box Name": boxes_lost.keys(), "Reason Dropped": boxes_lost.values()})
    typeless_samples.append((np.nan, np.nan))
    typeless_samples = np.array(typeless_samples)
    typeless_df = pd.DataFrame.from_dict({"Sample ID": typeless_samples[:, 0], "Box Name": typeless_samples[:, 1]})
    output_fname = 'aggregate_inventory_{}.xlsx'.format(date.today().strftime("%m.%d.%y"))
    with pd.ExcelWriter(util.proc + output_fname) as writer:
        for sample_type, data in samples_data.items():
            if sample_type != 'All':
                df = pd.DataFrame(data)
                df.to_excel(writer, sheet_name='{}'.format(sample_type), index=False)
        box_df.to_excel(writer, sheet_name='Uploaded Box Counts', index=False)
        lost_df.to_excel(writer, sheet_name="Unprocessed Boxes", index=False)
        typeless_df.to_excel(writer, sheet_name="Typeless Samples", index=False)
    if box_df.shape[0] > 0:
        print("Boxes to upload today:")
        for name in box_df['Name']:
            print(name)
    else:
        print(f"Nothing to upload, check Unprocessed Boxes tab of {output_fname}")
