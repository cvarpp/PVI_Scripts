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
    else:
        print(f"No number found in sheet '{sheet_name}'. Skipping...")
        return [], "", 0

    box_name = f"CML NPS {number} {box_type}"

    if df['sample id'].duplicated().any():
        print(f"Duplicate 'Sample ID' found in box: {box_name}. Skipping...")
        return [], "", 0

    transformed_rows = []
    for idx, row in df.iterrows():
        transformed_row = {
            'Name': row['sample id'],
            'Sample ID': row['sample id'],
            'Sample Type': row.get('sample type', ''),
            'Freezer': 'Temporary CML NPS',
            'Level1': 'freezer_cml',
            'Level2': 'shelf_cml',
            'Level3': 'rack_cml',
            'Box': box_name,
            'Position': pos_convert(idx),
            'ALIQUOT': row.get('aliquot', '')
        }
        transformed_rows.append(transformed_row)
    return transformed_rows, box_name, len(transformed_rows)

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='FP upload for CML NPS box')
    args = argParser.parse_args()

    today_date = datetime.now().strftime("%Y.%m.%d")
    nps_data = pd.ExcelFile(os.path.join(util.proc, 'New Import Sheet.xlsx'))

    all_data = []
    boxes_to_delete = []

    for sheet_name in nps_data.sheet_names:
        if 'cml' in sheet_name.lower():
            df = pd.read_excel(nps_data, sheet_name=sheet_name)
            transformed_data, box_name, tube_count = transform_data(sheet_name, df)
            if transformed_data:
                all_data.extend(transformed_data)
                boxes_to_delete.append((box_name, tube_count))

    output_df = pd.DataFrame(all_data)
    output_file_path = os.path.join(util.proc, f'cml_fp_upload_{today_date}.xlsx')
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        output_df.to_excel(writer, index=False, sheet_name='CML NPS Upload')
        boxes_to_delete_df = pd.DataFrame(boxes_to_delete, columns=['Box Name', 'Tube Count'])
        boxes_to_delete_df.to_excel(writer, index=False, sheet_name='Box to Delete')
    
    print(f"Output saved to {output_file_path}.")
    print("Boxes to upload:")
    for box_name in output_df['Box'].unique():
        print(box_name)
