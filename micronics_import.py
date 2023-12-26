import pandas as pd
import argparse
import util
from copy import deepcopy

def convert_tube_position(position):
    '''Converts Tube Position from 'A01' to 'A/1' format'''
    letter = position[0]
    number = int(position[1:])
    return f"{letter}/{number}"

def transform_sample_data(sheet):
    '''Transforms each row of the sample sheet and returns the transformed data'''
    transformed_data = deepcopy(data)
    for idx, row in sheet.iterrows():
        tube_position = convert_tube_position(row['Tube Position'])
        box_name = "PVI Micronic " + str(row['Rack ID'])
        sample_id = str(row['Sample ID'])
        transformed_data['Name'].append(sample_id)
        transformed_data['Sample ID'].append(sample_id)
        transformed_data['Sample Type'].append('Serum')
        transformed_data['Freezer'].append(f"Annaberg {int(tube_position.split('/')[1]) + 17}")
        transformed_data['Level1'].append(f"Level {int(tube_position.split('/')[1])}")
        transformed_data['Level2'].append(f"Shelf {int(tube_position.split('/')[1])}")
        transformed_data['Level3'].append(f"Rack {int(tube_position.split('/')[1])}")
        transformed_data['Box'].append(box_name)
        transformed_data['Position'].append(tube_position)
        transformed_data['ALIQUOT'].append(sample_id)
        transformed_data['Tube ID'].append(row['Tube ID'])
    return transformed_data

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Transform micronics data for FP upload')
    argParser.add_argument('-m', '--min_count', action='store', type=int, default=96)
    args = argParser.parse_args()

    input_file = util.proc + 'CRP Micronics Files/' + 'CRP Micronics Plate 8.xlsx'
    inventory_boxes = pd.read_excel(input_file, sheet_name=None)

    data = {'Name': [], 'Sample ID': [], 'Sample Type': [], 'Freezer': [], 'Level1': [], 
            'Level2': [], 'Level3': [], 'Box': [], 'Position': [], 'ALIQUOT': [], 'Tube ID': []}
    
    for name, sheet in inventory_boxes.items():
        transformed_sample_data = transform_sample_data(sheet)
        for key in data:
            data[key].extend(transformed_sample_data[key])

    output_file = util.proc + 'CRP Micronics Files/' + 'micronics_fp_upload TEST.xlsx'
    pd.DataFrame(data).to_excel(output_file, index=False)
    print(f"output is saved to {output_file}.")