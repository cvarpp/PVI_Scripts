import pandas as pd
import numpy as np
import argparse
import util
import os
from datetime import datetime

def convert_tube_position(position):
    '''Converts Tube Position from 'A01' to '1/A' format'''
    letter = position[0]
    number = int(position[1:])
    return f"{number}/{letter}"

def transform_sample_data(sheet, box_name, data):
    '''Transforms each row of the scanned sheet'''
    for idx, row in sheet.iterrows():
        tube_position = convert_tube_position(row['Tube Position'])
        sample_id = str(row['Sample ID'])
        data['Name'].append(sample_id)
        data['Sample ID'].append(sample_id)
        data['Sample Type'].append('Serum Micronic')
        data['Volume'].append('0.4')
        data['Freezer'].append('Annenberg 18')
        data['Level1'].append('Freezer 1 (Eiffel Tower)')
        data['Level2'].append('Shelf 4')
        data['Level3'].append('Rack 2')
        data['Box'].append(box_name)
        data['Position'].append(tube_position)
        data['ALIQUOT'].append(sample_id)
        data['Barcode'].append(row['Tube ID'])
        # data['Original Plate Barcode'].append(row['Rack ID'])
    return data

if __name__ == '__main__':
    micronics_folder = os.path.join(util.project_ws, 'CRP aliquoting/CRP Micronics Files/')
    argParser = argparse.ArgumentParser(description='Transform micronics data for FP upload')
    argParser.add_argument('-m', '--min_count', action='store', type=int, default=96, help='Minimum number of tubes for a plate to be considered inventory-ready')
    args = argParser.parse_args()

    today_date = datetime.now().strftime("%Y.%m.%d")

    data = {'Name': [], 'Sample ID': [], 'Sample Type': [], 'Volume': [], 'Freezer': [], 'Level1': [], 
            'Level2': [], 'Level3': [], 'Box': [], 'Position': [], 'ALIQUOT': [], 'Barcode': []}

    box_dfs = pd.read_excel(micronics_folder + 'CRP Micronics Import.xlsx', sheet_name=None)
    completed_boxes = box_dfs['Uploaded (Tabs to Delete)']['Plate Name'].unique()
    box_with_duplicate_barcodes = []

    for box_name, box_df in box_dfs.items():
        box_df = box_df.rename(columns={'Free Text': 'Sample ID'})
        if 'Serum' in box_name and box_name not in completed_boxes and 'Sample ID' in box_df.columns:
            try:
                assert (box_df.columns == np.array(['Tube Position', 'Tube ID', 'Rack ID', 'Date', 'Time', 'Sample ID', 'Status'])).all(), 'Column name issue'
                # assert 'NOREAD' not in list(box_df['Rack ID'].unique()), 'Rack barcode scan failed'
                assert (box_df.dropna(subset='Tube ID')['Status'].unique() == np.array(['Code OK'])).all(), 'Tube barcode scan failed'
                assert box_df['Sample ID'].count() == box_df['Tube ID'].count(), "Missing sample IDs or tube barcodes"
                
                if box_df['Tube ID'].duplicated().any():
                    box_with_duplicate_barcodes.append(box_name)
                    raise ValueError('Duplicate Tube ID barcodes found')
                
            except Exception as e:
                print(box_name, e)
                continue

            print(box_name, "is valid. Transforming...")
            transform_sample_data(box_df, box_name, data)

    output_df = pd.DataFrame(data)
    output_file = os.path.join(micronics_folder, f'micronics_fp_upload {today_date}.xlsx')
    output_df.to_excel(output_file, index=False)
    print(f"Output saved to {output_file}.")
    print("Boxes to upload:")
    for box_name in output_df['Box'].unique():
        print(box_name)