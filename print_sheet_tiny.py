import pandas as pd
import argparse
import util
import os


def get_box_range(print_type):
    # Rename Printing Log: Study - Kit Type
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=0)
            .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max', 'Study': 'Kit Type'})
            .drop(1))

    plog['Box Min'] = pd.to_numeric(plog['Box Min'], errors='coerce')
    plog['Box Max'] = pd.to_numeric(plog['Box Max'], errors='coerce')
    plog.dropna(subset=['Box Min', 'Box Max'], inplace=True)

    if print_type == 'SERONET':
        filtered_plog = plog[(plog['Kit Type'] == 'SERONET') & (plog['PBMCs'].isin(['No', 'no', 'NO']))]
    elif print_type == 'SERUM':
        filtered_plog = plog[(plog['Kit Type'] == 'SERUM')]
    elif print_type == 'STANDARD':
        filtered_plog = plog[(plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'].isin(['No', 'no', 'NO']))]
    elif print_type == 'SERONETPBMC':
        filtered_plog = plog[(plog['Kit Type'].isin(['SERONET', 'MIT (PBMCs)'])) & (plog['PBMCs'].isin(['Yes', 'yes', 'YES']))]
    elif print_type == 'STANDARDPBMC':
        filtered_plog = plog[(plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'].isin(['Yes', 'yes', 'YES']))]

    filtered_plog = filtered_plog.sort_values(by='Box Max', ascending=False).iloc[0]
    recent_box_max = filtered_plog['Box Max']

    if print_type == 'SERONET':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 6
    elif print_type == 'SERUM':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 4
    elif print_type == 'STANDARD':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 6
    elif print_type == 'SERONETPBMC':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 32
    elif print_type == 'STANDARDPBMC':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 32
    
    return int(box_start), int(box_end)


def get_sample_ids(sheet_name, box_start, box_end):
    print_planning_path = os.path.join(util.tube_print, 'Print Planning.xlsx')
    print_planning = pd.read_excel(print_planning_path, sheet_name=sheet_name)
    box_range = print_planning['Box ID'].apply(lambda bid: box_start <= bid <= box_end) # Use column name and loc instead of iloc
    sample_ids = print_planning.loc[box_range, 'Sample ID'].to_numpy()

    return sample_ids


def generate_workbook(assigned_sample_ids, box_start, box_end, sheet_name, template_folder, output_path, print_type):
    template_path = os.path.join(util.tube_print, 'Future Sheets', template_folder)
    template = pd.read_excel(template_path, sheet_name=None, header=None)

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in template.items():

            if print_type == 'STANDARD':
                if 'READ ME' in sheet_name:
                    box_numbers_range = range(box_start, box_end + 1)
                    sheet_data.loc[1:6, 7] = box_numbers_range
                    sheet_data.loc[7:24, 7] = assigned_sample_ids
                    sheet_data.loc[2:19, 12] = assigned_sample_ids
                    sheet_data.loc[2:19, 13] = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                if '1-Box Top round1' in sheet_name:
                    sheet_data.loc[1:6, 7] = assigned_sample_ids[:6]
                    sheet_data.loc[1:90, 1] = [sid for sid in assigned_sample_ids[:6] for _ in range(15)]
            
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

            # test use
            print(assigned_sample_ids)
            print(box_start, box_end)
            print(sheet_data.head())
            print(sheet_data.shape)
            print(sheet_data.index)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("print_type", choices=['SERONET', 'SERUM', 'STANDARD', 'SERONETPBMC', 'STANDARDPBMC'], help="Choose print type")
    args = parser.parse_args()

    # Input print_type (SERONET/SERUM/STANDARD/SERONETPBMC/STANDARDPBMC)
    print_type = args.print_type
    
    # Box range
    box_start, box_end = get_box_range(print_type)

    # Map print_type to workbook_name & template_file & output_path
    print_type_mapping = {
        'SERONET': ('Seronet Full', 'SERONET FULL/SERONET FULL Template.xlsx', 'Future Sheets/SERONET FULL'),
        'SERUM': ('Serum', 'SERUM/SERUM Template.xlsx', 'Future Sheets/SERUM'),
        'STANDARD': ('Standard', 'STANDARD/STANDARD Template.xlsx', 'Future Sheets/STANDARD'),
        'SERONETPBMC': ('Seronet Full', 'SERONET FULL/SERONET FULL PBMC Template.xlsx', 'Future Sheets/SERONET FULL'),
        'STANDARDPBMC': ('Standard', 'STANDARD/STANDARD PBMC Template.xlsx', 'Future Sheets/STANDARD')
    }
    sheet_name, template_file, output_folder = print_type_mapping[print_type]

    # Sample IDs within box range
    assigned_sample_ids = get_sample_ids(sheet_name, box_start, box_end)

    # Output
    workbook_name = f"{sheet_name.upper()} {'PBMC ' if 'PBMC' in print_type else ''}{box_start}-{box_end} test Tiny"
    template_path = os.path.join(util.tube_print, 'Future Sheets', template_file)
    output_path = os.path.join(util.tube_print, output_folder, f"{workbook_name}.xlsx")

    generate_workbook(assigned_sample_ids, box_start, box_end, workbook_name, template_path, output_path, print_type)

    print(f"'{workbook_name}' workbook generated in {output_folder}.")

