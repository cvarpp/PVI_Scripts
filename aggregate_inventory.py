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

def pos_convert(idx):
    '''Converts a 0-80 index to a box position'''
    return str((idx // 9) + 1) + "/" + "ABCDEFGHI"[idx % 9]

class Box:
    def __init__(self, name, sheet, racks_in_freezers, boxes_lost):
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
        if self.rack_number not in racks_in_freezers.index:
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
            if self.sample_type in ['PBMC', 'HT', '4.5 mL Tube']:
                self.box_kinds.append('')
            else:
                boxes_lost[name] = "Neither lab nor FF"
                return
        if not self.box_done:
            boxes_lost[name] = 'Not marked complete'
        rack_info = racks_in_freezers.loc[self.rack_number, :]
        self.freezer = str(rack_info['Farm']) + " New"
        self.level1 = str(rack_info['Freezer'])
        self.level2 = str(rack_info['Level2'])
        if self.sample_type != 'PBMC':
            self.level3 = "Rack " + str(int(self.rack_number))
        else:
            self.level3 = ""
        self.valid = True


if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Aggregate inventory of each sample type for FP upload')
    argParser.add_argument('-m', '--min_count', action='store', type=int, default=81)
    args = argParser.parse_args()
    inventory_boxes = pd.read_excel(util.inventory_input, sheet_name=None)
    racks_in_freezers = pd.read_excel(util.freezer_map, sheet_name='Racks in Freezers').dropna(subset='Rack ID', axis=0).copy().set_index('Rack ID')
    sample_types = ['Plasma', 'Serum', 'Pellet', 'Saliva', 'PBMC', 'HT', '4.5 mL Tube', 'NPS', 'RNA', 'All']
    data = {'Name': [], 'Sample ID': [], 'Sample Type': [],'Freezer': [],'Level1': [],'Level2': [],'Level3': [],'Box': [],'Position': [], 'ALIQUOT': []}
    samples_data = {st: deepcopy(data) for st in sample_types if st != 'HT'}
    samples_data['Serum']['Heat treated?'] = []
    samples_data['Plasma']['Heat treated?'] = []
    samples_data['4.5 mL Tube']['Serum or Plasma?'] = []
    samples_data['All'] = deepcopy(data)
    box_counts = {}
    boxes_lost = {}
    typeless_samples = []

    for name, sheet in inventory_boxes.items():
        boxx = Box(name, sheet, racks_in_freezers, boxes_lost)
        if not boxx.valid or not boxx.box_done:
            continue
        sheet = sheet.assign(sample_id=clean_sample_id).set_index('sample_id')
        for kind in boxx.box_kinds:
            if "ATLAS" in name.upper() and boxx.sample_type in ['Serum', 'Plasma', 'Saliva']:
                kind = boxx.box_kinds[0]
                box_name = f"ATLAS {boxx.sample_type} {kind} {boxx.box_number}"
            elif boxx.team == 'APOLLO':
                box_name = f"{boxx.team} {kind} {boxx.box_number}"
            elif boxx.sample_type in ['PBMC', 'HT', '4.5 mL Tube']:
                box_name = f"{boxx.team} {boxx.sample_type} {boxx.box_number}"
            else:
                box_name = f"{boxx.team} {boxx.sample_type} {kind} {boxx.box_number}"
            box_counts[box_name] = 0
            for idx, (sample_id, row) in enumerate(sheet.iterrows()):
                sample_id = str(sample_id).strip().upper()
                if re.search("((PV)|[1-9A-Z])[0-9]{4,6}", sample_id) is None:
                    continue
                sample_type = 'N/A'
                for val in sample_types:
                    if re.search(str(val).upper().split()[0], str(row['Sample Type']).upper()):
                        sample_type = val
                        break
                if sample_type == 'N/A':
                    typeless_samples.append((sample_id, box_name))
                else:
                    for sname in [sample_type, 'All']:
                        data = samples_data[sname]
                        for col in ['Name', 'Sample ID', 'ALIQUOT']:
                            data[col].append(sample_id)
                        data['Sample Type'].append(sample_type)
                        data['Freezer'].append(boxx.freezer)
                        data['Level1'].append(boxx.level1)
                        data['Level2'].append(boxx.level2)
                        data['Level3'].append(boxx.level3)
                        data['Box'].append(box_name)
                        data['Position'].append(pos_convert(idx))
                        if sname in ['Serum', 'Plasma']:
                            if boxx.sample_type == 'HT':
                                data['Heat treated?'].append('Yes')
                            else:
                                data['Heat treated?'].append('No')
                        if sname == '4.5 mL Tube':
                            data['Serum or Plasma?'].append(row['Serum or Plasma?'])
                    box_counts[box_name] += 1
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
