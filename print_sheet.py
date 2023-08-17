import pandas as pd
import argparse
import util
import os


def get_box_range(input_type):
    # Rename Printing Log: Study - Kit Type
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=0)
            .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max', 'Study': 'Kit Type'})
            .drop(1))

    plog['Box Min'] = pd.to_numeric(plog['Box Min'], errors='coerce')
    plog['Box Max'] = pd.to_numeric(plog['Box Max'], errors='coerce')
    plog.dropna(subset=['Box Min', 'Box Max'], inplace=True)

    if input_type == 'SERONET':
        filtered_plog = plog[(plog['Kit Type'] == 'SERONET') & (plog['PBMCs'] == 'No')]
    elif input_type == 'SERUM':
        filtered_plog = plog[(plog['Kit Type'] == 'SERUM')]
    elif input_type == 'STANDARD':
        filtered_plog = plog[(plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'] == 'No')]
    elif input_type == 'SERONETPBMC':
        filtered_plog = plog[(plog['Kit Type'].isin(['SERONET', 'MIT (PBMCs)'])) & (plog['PBMCs'] == 'Yes')]
    elif input_type == 'STANDARDPBMC':
        filtered_plog = plog[(plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'] == 'Yes')]

    filtered_plog = plog[(plog['Kit Type'] == input_type)].sort_values(by='Box Max', ascending=False).iloc[0]
    recent_box_max = filtered_plog['Box Max']

    if input_type == 'SERONET':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 6
    elif input_type == 'SERUM':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 4
    elif input_type == 'STANDARD':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 6
    elif input_type == 'SERONETPBMC':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 32
    elif input_type == 'STANDARDPBMC':
        box_start = recent_box_max + 1
        box_end = recent_box_max + 32
    
    return int(box_start), int(box_end)


def get_sample_ids(sheet_name, box_start, box_end):
    print_planning_path = os.path.join(util.tube_print, 'Print Planning.xlsx')
    print_planning = pd.read_excel(print_planning_path, sheet_name=sheet_name)
    box_range = print_planning['Box ID'].apply(lambda bid: box_start <= bid <= box_end) # Use column name and loc instead of iloc
    sample_ids = print_planning.loc[box_range, 'Sample ID'].to_numpy()

    return sample_ids


def seronet_workbook(assigned_sample_ids, box_start, box_end, workbook_name):
    template_path = os.path.join(util.tube_print, 'Future Sheets', 'SERONET FULL', 'SERONET FULL Template.xlsx')
    template = pd.read_excel(template_path, sheet_name=None)

    output_seronet = os.path.join(util.tube_print, 'Future Sheets', 'SERONET FULL', f"{workbook_name}.xlsx")

    with pd.ExcelWriter(output_seronet, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in template.items():
            if 'READ ME' in sheet_name:
                sheet_data.loc[2:19, 'M'] = assigned_sample_ids
                box_numbers = [box for _ in range(3) for box in range(box_start, box_end + 1)]
                sheet_data.loc[2:19, 'N'] = box_numbers
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=True)


def serum_workbook(assigned_sample_ids, box_start, box_end, workbook_name):
    template_path = os.path.join(util.tube_print, 'Future Sheets', 'SERUM', 'Serum Template.xlsx')
    template = pd.read_excel(template_path, sheet_name=None)

    output_serum = os.path.join(util.tube_print, 'Future Sheets', 'SERUM', f"{workbook_name}.xlsx")

    with pd.ExcelWriter(output_serum, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in template.items():
            if 'READ ME' in sheet_name:
                sheet_data.loc[1:24, 'N'] = assigned_sample_ids
                box_numbers = [box for _ in range(6) for box in range(box_start, box_end + 1)]
                sheet_data.loc[1:24, 'O'] = box_numbers
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)


def standard_workbook(assigned_sample_ids, box_start, box_end, workbook_name):
    template_path = os.path.join(util.tube_print, 'Future Sheets', 'STANDARD', 'STANDARD Template.xlsx')
    template = pd.read_excel(template_path, sheet_name=None)

    output_standard = os.path.join(util.tube_print, 'Future Sheets', 'STANDARD', f"{workbook_name}.xlsx")

    with pd.ExcelWriter(output_standard, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in template.items():
            if 'READ ME' in sheet_name:
                sheet_data.loc[2:19, 'M'] = assigned_sample_ids
                box_numbers = [box for _ in range(3) for box in range(box_start, box_end + 1)]
                sheet_data.loc[2:19, 'N'] = box_numbers
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)


def seronet_pbmc_workbook(assigned_sample_ids, box_start, box_end, workbook_name):
    template_path = os.path.join(util.tube_print, 'Future Sheets', 'SERONET FULL', 'SERONET FULL PBMC Template.xlsx')
    template = pd.read_excel(template_path, sheet_name=None)

    output_seronet_pbmc = os.path.join(util.tube_print, 'Future Sheets', 'SERONET PBMC', f"{workbook_name}.xlsx")

    with pd.ExcelWriter(output_seronet_pbmc, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in template.items():
            if 'READ ME' in sheet_name:
                sheet_data.loc[1:96, 'V'] = assigned_sample_ids
                box_numbers = [box for _ in range(3) for box in range(box_start, box_end + 1)]
                sheet_data.loc[1:96, 'W'] = box_numbers
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)


def standard_pbmc_workbook(assigned_sample_ids, box_start, box_end, workbook_name):
    template_path = os.path.join(util.tube_print, 'Future Sheets', 'STANDARD', 'STANDARD PBMC Template.xlsx')
    template = pd.read_excel(template_path, sheet_name=None)

    output_standard_pbmc = os.path.join(util.tube_print, 'Future Sheets', 'STANDARD PBMC', f"{workbook_name}.xlsx")

    with pd.ExcelWriter(output_standard_pbmc, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in template.items():
            if 'READ ME' in sheet_name:
                sheet_data.loc[1:96, 'V'] = assigned_sample_ids
                box_numbers = [box for _ in range(3) for box in range(box_start, box_end + 1)]
                sheet_data.loc[1:96, 'W'] = box_numbers
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("print_type", choices=['SERONET', 'SERUM', 'STANDARD', 'SERONETPBMC', 'STANDARDPBMC'], help="Choose print type")
    args = parser.parse_args()

    # Input: print_type (SERONET/SERUM/STANDARD/SERONETPBMC/STANDARDPBMC)
    print_type = args.print_type
    
    # Printing Log: box range
    box_start, box_end = get_box_range(print_type)

    # Map print_type to sheet_name & workbook
    print_type_mapping = {
        'SERONET': ('Seronet Full', seronet_workbook),
        'SERUM': ('Serum', serum_workbook),
        'STANDARD': ('Standard', standard_workbook),
        'SERONETPBMC': ('Seronet Full', seronet_pbmc_workbook),
        'STANDARDPBMC': ('Standard Standard', standard_pbmc_workbook)
    }
    sheet_name, workbook_generation = print_type_mapping[print_type]

    # Print Planning: sample IDs
    assigned_sample_ids = get_sample_ids(sheet_name, box_start, box_end)

    # Output
    if print_type in ['SERONET', 'SERUM', 'STANDARD']:
        workbook_name = f"{sheet_name.upper()} {box_start}-{box_end}"
    elif print_type in ['SERONETPBMC', 'STANDARDPBMC']:
        workbook_name = f"{print_type.replace('PBMC', ' PBMC')} {box_start}-{box_end}"
    
    workbook_generation(assigned_sample_ids, box_start, box_end, workbook_name)
    
    print("Output workbook is under PVI/ Print Shop/ Tube Printing/ Future Sheets.")

