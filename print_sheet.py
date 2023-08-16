import pandas as pd
import argparse
import util
import os


def get_box_range(plog, input_type):
    plog.dropna(subset=['Box Min', 'Box Max'], inplace=True)
    plog['Box Min'] = pd.to_numeric(plog['Box Min'], errors='coerce')
    plog['Box Max'] = pd.to_numeric(plog['Box Max'], errors='coerce')

    filtered_plog = plog[plog['Kit Type'] == input_type].sort_values(by='Box Max', ascending=False).iloc[0]
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
    
    return int(box_start), int(box_end)


def get_sample_ids(print_type, box_start, box_end):
    if print_type == 'SERONET':
        sheet_name = 'Seronet Full'
    elif print_type == 'SERUM':
        sheet_name = 'Serum'
    elif print_type == 'STANDARD':
        sheet_name = 'Standard'

    print_planning_path = os.path.join(util.tube_print, 'Print Planning.xlsx')
    print_planning = pd.read_excel(print_planning_path, sheet_name=sheet_name)
    box_range = (print_planning.iloc[:, 1] >= box_start) & (print_planning.iloc[:, 1] <= box_end)
    sample_ids = print_planning.loc[box_range, print_planning.columns[0]].reset_index(drop=True)

    return sample_ids


def get_output_file(print_type, workbook_name):
    if print_type == 'SERONET':
        return os.path.join(util.tube_print, 'Future Sheets', 'SERONET FULL', f"{workbook_name}.xlsx")
    elif print_type == 'SERUM':
        return os.path.join(util.tube_print, 'Future Sheets', 'SERUM', f"{workbook_name}.xlsx")
    elif print_type == 'STANDARD':
        return os.path.join(util.tube_print, 'Future Sheets', 'STANDARD', f"{workbook_name}.xlsx")


def seronet_workbook(assigned_sample_ids, box_start, box_end, workbook_name):
    template1_path = os.path.join(util.tube_print, 'Future Sheets', 'SERONET FULL', 'SERONET FULL Template.xlsx')
    template1 = pd.read_excel(template1_path, sheet_name=None)
    # template2_path = os.path.join(util.tube_print, 'Future Sheets', 'SERONET FULL', 'SERONET FULL PBMC Template.xlsx')
    # template2 = pd.read_excel(template2_path, sheet_name=None)

    output_file_seronet = os.path.join(util.tube_print, 'Future Sheets', 'SERONET FULL', f"{workbook_name}.xlsx")

    with pd.ExcelWriter(output_file_seronet, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in template1.items():
            if 'READ ME' in sheet_name:
                sheet_data.loc[2:19, 'M'] = assigned_sample_ids
                box_numbers = [box for _ in range(3) for box in range(box_start, box_end + 1)]
                sheet_data.loc[2:19, 'N'] = box_numbers
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=True)


def serum_workbook(writer, sample_ids, box_start, box_end):
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


def standard_workbook(writer, sample_ids, box_start, box_end):
    template1_path = os.path.join(util.tube_print, 'Future Sheets', 'STANDARD', 'STANDARD Template.xlsx')
    template1 = pd.read_excel(template1_path, sheet_name=None)
    # template2_path = os.path.join(util.tube_print, 'Future Sheets', 'STANDARD', 'STANDARD PBMC Template.xlsx')
    # template2 = pd.read_excel(template2_path, sheet_name=None)

    output_standard = os.path.join(util.tube_print, 'Future Sheets', 'STANDARD', f"{workbook_name}.xlsx")

    with pd.ExcelWriter(output_standard, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in template1.items():
            if 'READ ME' in sheet_name:
                sheet_data.loc[2:19, 'M'] = assigned_sample_ids
                box_numbers = [box for _ in range(3) for box in range(box_start, box_end + 1)]
                sheet_data.loc[2:19, 'N'] = box_numbers
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("print_type", choices=['SERONET', 'SERUM', 'STANDARD'], help="Choose print type")
    args = parser.parse_args()

    # Rename Printing Log: Study - Kit Type
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=0)
            .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max', 'Study': 'Kit Type'})
            .drop(1))

    # Input: print_type (SERONET/SERUM/STANDARD)
    print_type = args.print_type
    
    # Printing Log: box range
    box_start, box_end = get_box_range(plog, print_type)

    # Print Planning: sample IDs
    assigned_sample_ids = get_sample_ids(print_type, box_start, box_end)

    # Output
    if print_type == 'SERONET':
        workbook_name = f"{print_type} FULL {box_start}-{box_end}"
    else:
        workbook_name = f"{print_type} {box_start}-{box_end}"

    output_file = get_output_file(print_type, workbook_name)

    if print_type == 'SERONET':
        seronet_workbook(assigned_sample_ids, box_start, box_end, workbook_name)
    elif print_type == 'SERUM':
        serum_workbook(assigned_sample_ids, box_start, box_end, workbook_name)
    elif print_type == 'STANDARD':
        standard_workbook(assigned_sample_ids, box_start, box_end, workbook_name)
    
    print("Done! Output workbook is under PVI/ Print Shop/ Tube Printing/ Future Sheets.")

