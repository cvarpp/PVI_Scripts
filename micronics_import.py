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

def transform_sample_data(sheet, box_file_name):
    '''Transforms each row of the scanned sheet'''
    transformed_data = deepcopy(data)
    for idx, row in sheet.iterrows():
        tube_position = convert_tube_position(row['Tube Position'])
        box_name = box_file_name
        sample_id = str(row['Sample ID'])
        transformed_data['Name'].append(sample_id)
        transformed_data['Sample ID'].append(sample_id)
        transformed_data['Sample Type'].append('Serum Micronic')
        transformed_data['Volume'].append('0.4')
        transformed_data['Freezer'].append('Annenberg 18')
        transformed_data['Level1'].append('Freezer 1 (Eiffel Tower)')
        transformed_data['Level2'].append('Shelf 4')
        transformed_data['Level3'].append('Rack 2')
        transformed_data['Box'].append(box_name)
        transformed_data['Position'].append(tube_position)
        transformed_data['ALIQUOT'].append(sample_id)
        transformed_data['Barcode'].append(row['Tube ID'])
        transformed_data['Original Plate Barcode'].append(row['Rack ID'])
    return transformed_data

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
            transformed_sample_data = transform_sample_data(sheet, box_file_name)
            for key in data:
                data[key].extend(transformed_sample_data[key])
    
    output_file = os.path.join(util.project_ws, 'CRP aliquoting/CRP Micronics Files/', f'micronics_fp_upload {today_date}.xlsx')
    pd.DataFrame(data).to_excel(output_file, index=False)
    print(f"output is saved to {output_file}.")