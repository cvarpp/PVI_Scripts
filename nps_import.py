import pandas as pd
import os
from datetime import datetime
import argparse
import util
import re

def pos_convert(idx):
    '''Converts a 0-80 index to a box position'''
    return str((idx // 9) + 1) + "/" + "ABCDEFGHI"[idx % 9]

def transform_data(sheet_name, df):
    '''Transforms data from Excel sheet, removing empty rows and checking for duplicates'''
    df.columns = df.columns.str.strip().str.lower()

    if 'sample id' not in df.columns:
        print(f"Column 'Sample ID' not found in sheet '{sheet_name}'. Skipping...")
        return [], "", 0
    
    df = df.dropna(subset=['sample id'])

    if 'psp' in sheet_name.lower():
        box_group = 'PSP'
        if 'ff' in sheet_name.lower():
            box_type = 'FF'
        elif 'lab' in sheet_name.lower():
            box_type = 'Lab'
        else:
            print(f"Sheet '{sheet_name}' does not contain 'ff' or 'lab'. Skipping...")
            return [], "", 0
        match = re.search(r'\d+', sheet_name)
        if match:
            number = match.group()
            box_name = f"{box_group} NPS {box_type} {number}"
        else:
            print(f"No number found in sheet '{sheet_name}'. Skipping...")
            return [], "", 0
    elif 'cml' in sheet_name.lower():
        box_group = 'CML'
        if 'a' in sheet_name.lower():
            box_type = 'A'
        elif 'b' in sheet_name.lower():
            box_type = 'B'
        else:
            print(f"Sheet '{sheet_name}' does not contain 'A' or 'B'. Skipping...")
            return [], "", 0
        match = re.search(r'\d+', sheet_name)
        if match:
            number = match.group()
            box_name = f"{box_group} NPS {number}{box_type}"
        else:
            print(f"No number found in sheet '{sheet_name}'. Skipping...")
            return [], "", 0
    else:
        print(f"Sheet '{sheet_name}' does not contain 'PSP' or 'CML'. Skipping...")
        return [], "", 0

    if df['sample id'].duplicated().any():
        print(f"Duplicate 'Sample ID' found in box: {box_name}. Skipping...")
        return [], "", 0

    transformed_rows = []
    for idx, row in df.iterrows():
        transformed_row = {
            'Name': row['sample id'],
            'Sample ID': row['sample id'],
            'Sample Type': row.get('sample type', ''),
            'Freezer': 'Temporary ' + box_group + ' NPS',
            'Level1': 'freezer_nps'if box_group == 'PSP' else 'freezer_cml',
            'Level2': 'shelf_nps' if box_group == 'PSP' else 'shelf_cml',
            'Level3': 'rack_nps' if box_group == 'PSP' else 'rack_cml',
            'Box': box_name,
            'Position': pos_convert(idx),
            'ALIQUOT': row.get('aliquot', '')
        }
        transformed_rows.append(transformed_row)
    return transformed_rows, box_name, len(transformed_rows)

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='FP upload for PSP NPS and CML NPS box')
    args = argParser.parse_args()

    today_date = datetime.now().strftime("%Y.%m.%d")
    nps_data = pd.ExcelFile(os.path.join(util.proc, 'New Import Sheet.xlsx'))

    psp_data = []
    cml_data = []
    boxes_to_delete = []

    for sheet_name in nps_data.sheet_names:
        df = pd.read_excel(nps_data, sheet_name=sheet_name)
        transformed_data, box_name, tube_count = transform_data(sheet_name, df)
        if transformed_data:
            if 'PSP' in box_name:
                psp_data.extend(transformed_data)
            elif 'CML' in box_name:
                cml_data.extend(transformed_data)
            boxes_to_delete.append((box_name, tube_count))

    output_file_path = os.path.join(util.proc, f'psp_cml_nps_fp_upload_{today_date}.xlsx')
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        if psp_data:
            pd.DataFrame(psp_data).to_excel(writer, index=False, sheet_name='PSP NPS Upload')
        if cml_data:
            pd.DataFrame(cml_data).to_excel(writer, index=False, sheet_name='CML NPS Upload')
        if boxes_to_delete:
            pd.DataFrame(boxes_to_delete, columns=['Box Name', 'Tube Count']).to_excel(writer, index=False, sheet_name='Box to Delete')
    
    print(f"Output saved to {output_file_path}.")
    print("Boxes to upload:")
    if psp_data:
        print("PSP Boxes:")
        for box_name in pd.DataFrame(psp_data)['Box'].unique():
            print(box_name)
    if cml_data:
        print("CML Boxes:")
        for box_name in pd.DataFrame(cml_data)['Box'].unique():
            print(box_name)
            