import pandas as pd
import argparse
import util
import os

def next_sample_ids(mast_list, input_type):
    if input_type == 'STANDARD':
        num_samples = 6
    elif input_type == 'SERONET':
        num_samples = 18
    elif input_type == 'SERUM':
        num_samples = 24
    
    unassigned_samples = mast_list[mast_list['Kit Type'].isnull()].head(num_samples)
    return unassigned_samples['Sample ID'].tolist()

def box_range(plog, input_type):
    plog.dropna(subset=['Box Min', 'Box Max'], inplace=True)
    plog['Box Min'] = pd.to_numeric(plog['Box Min'], errors='coerce')
    plog['Box Max'] = pd.to_numeric(plog['Box Max'], errors='coerce')
    # plog['Box Min'] = plog['Box Min'].astype(int, errors='ignore')
    # plog['Box Max'] = plog['Box Max'].astype(int, errors='ignore')

    filtered_plog = plog[plog['Kit Type'] == input_type].sort_values(by='Box Max', ascending=False).iloc[0]
    recent_box_max = filtered_plog['Box Max']

    if input_type == 'STANDARD':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 6
    elif input_type == 'SERONET':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 6
    elif input_type == 'SERUM':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 4
    
    return box_start, box_end


if __name__ == '__main__':

    parser = argparse.ArgumentParser()
    parser.add_argument("print_type", choices=['STANDARD', 'SERONET', 'SERUM'], help="Choose print type")
    args = parser.parse_args()

    # Rename mast_list: Location - Kit Type, Study ID - Sample ID; plog: Study - Kit Type
    mast_list = (pd.read_excel(util.tracking + "Sample ID Master List.xlsx", sheet_name='Master Sheet', header=0)
                   .rename(columns={'Location': 'Kit Type', 'Study ID': 'Sample ID'}))
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=0)
              .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max', 'Study': 'Kit Type'})
              .drop(1))
    
    # Input
    print_type = args.print_type
    
    # mast_list: unassigned sample IDs
    unassigned_sample_ids = next_sample_ids(mast_list, print_type)

    # plog: box range
    box_start, box_end = box_range(plog, print_type)
    workbook_name = f"{print_type} {box_start}-{box_end}"
    
    # output path
    if print_type == 'STANDARD':
        output_path = util.tube_print/'Future Sheets'/'STANDARD'
    elif print_type == 'SERONET':
        output_path = util.tube_print/'Future Sheets'/'SERONET FULL'
    elif print_type == 'SERUM':
        output_path = util.tube_print/'Future Sheets'/'SERUM'
    
    # ## test
    # print("Printing Type:", print_type)
    # print("Unassigned Sample IDs:", unassigned_sample_ids)
    # print("Workbook Name:", workbook_name)
    # print("Output Path:", output_path)

