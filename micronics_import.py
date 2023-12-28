import pandas as pd
import argparse
import util
import os
from copy import deepcopy
from datetime import datetime

def convert_tube_position(position):
    '''Converts Tube Position from 'A01' to '1/A' format'''
    letter = position[0]
    number = int(position[1:])
    return f"{number}/{letter}"

def transform_sample_data(sheet, box_file_name, data):
    '''Transforms each row of the scanned sheet'''
    for idx, row in sheet.iterrows():
        tube_position = convert_tube_position(row['Tube Position'])
        box_name = box_file_name
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
        data['Original Plate Barcode'].append(row['Rack ID'])
    return data

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Transform micronics data for FP upload')
    argParser.add_argument('filename', nargs='+', type=str, help='Filename of the input Excel file')
    argParser.add_argument('-m', '--min_count', action='store', type=int, default=96)
    args = argParser.parse_args()

    today_date = datetime.now().strftime("%Y.%m.%d")

    data = {'Name': [], 'Sample ID': [], 'Sample Type': [], 'Volume': [], 'Freezer': [], 'Level1': [], 
            'Level2': [], 'Level3': [], 'Box': [], 'Position': [], 'ALIQUOT': [], 'Barcode': [], 
            'Original Plate Barcode': []}

    for filename in args.filename:
        input_file = os.path.join(util.project_ws, 'CRP aliquoting/CRP Micronics Files/', filename)
        file_name_parts = os.path.splitext(os.path.basename(input_file))[0].split()
        box_file_name = ' '.join(file_name_parts[:4])
        inventory_box = pd.read_excel(input_file, sheet_name=None)

        for name, sheet in inventory_box.items():
            transform_sample_data(sheet, box_file_name, data)
    
    output_file = os.path.join(util.project_ws, 'CRP aliquoting/CRP Micronics Files/', f'micronics_fp_upload {today_date}.xlsx')
    pd.DataFrame(data).to_excel(output_file, index=False)
    print(f"output is saved to {output_file}.")