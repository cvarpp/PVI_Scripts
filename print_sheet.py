import pandas as pd
import argparse
import util
import os


def get_box_range(plog, input_type):
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
    
    return int(box_start), int(box_end)


def get_sample_ids(print_type, box_start, box_end):
    if print_type == 'STANDARD':
        sheet_name = 'Standard'
    elif print_type == 'SERONET':
        sheet_name = 'Seronet Full'
    elif print_type == 'SERUM':
        sheet_name = 'Serum'

    print_planning_path = os.path.join(util.tube_print, 'Print Planning.xlsx')
    print_planning = pd.read_excel(print_planning_path, sheet_name=sheet_name)
    box_range = (print_planning.iloc[:, 1] >= box_start) & (print_planning.iloc[:, 1] <= box_end)
    sample_ids = print_planning.loc[box_range, print_planning.columns[0]].reset_index(drop=True)

    return sample_ids


def get_output_file(print_type, workbook_name):
    if print_type == 'STANDARD':
        return os.path.join(util.tube_print, 'Future Sheets', 'STANDARD', f"{workbook_name}.xlsx")
    elif print_type == 'SERONET':
        return os.path.join(util.tube_print, 'Future Sheets', 'SERONET FULL', f"{workbook_name}.xlsx")
    elif print_type == 'SERUM':
        return os.path.join(util.tube_print, 'Future Sheets', 'SERUM', f"{workbook_name}.xlsx")


def serum_worksheet(writer, sample_ids, box_start, box_end):
    # Sheet 1: 1 - Kits
    kits = {
        'Column A': range(1, len(sample_ids) * 2 + 1),
        'Column B': [sample_id for sample_id in sample_ids for _ in range(2)],
        'Column C': 'Serum',
        # 'Column M': range(1, len(sample_ids) + 1),
        # 'Column N': sample_ids,
        # 'Column O': [box for _ in range(6) for box in range(box_start, box_end + 1)]
    }
    kits_df = pd.DataFrame(kits)
    kits_df.to_excel(writer, sheet_name='1 - Kits', index=False, header=False)

    # Sheet 2: 2 - Tops
    tops = {
        'Column A': range(1, len(sample_ids) * 7 + 1),
        'Column B': [sample_id for sample_id in sample_ids for _ in range(7)],
        'Column D': 'Serum'
    }
    tops_df = pd.DataFrame(tops)
    tops_df.to_excel(writer, sheet_name='2 - Tops', index=False, header=False)

    # Sheet 3: 3 - Sides
    sides = {
        'Column A': range(1, len(sample_ids) * 7 + 1),
        'Column B': 'Serum',
        'Column C': [sample_id for sample_id in sample_ids for _ in range(7)],
        'Column D': 'SERUM'
    }
    sides_df = pd.DataFrame(sides)
    sides_df.to_excel(writer, sheet_name='3 - Sides', index=False, header=False)

    # Sheet 4: 4 - 4.5 mL Tops
    tops_4point5 = {
        'Column A': range(1, len(sample_ids) + 1),
        'Column C': sample_ids,
        'Column D': 'Serum'
    }
    tops_4point5_df = pd.DataFrame(tops_4point5)
    tops_4point5_df.to_excel(writer, sheet_name='4 - 4.5 mL Tops', index=False, header=False)

    # Sheet 5: 5 - 4.5 mL Sides
    sides_4point5 = {
        'Column A': range(1, len(sample_ids) + 1),
        'Column B': sample_ids,
        'Column C': 'Serum'
    }
    sides_4point5_df = pd.DataFrame(sides_4point5)
    sides_4point5_df.to_excel(writer, sheet_name='5 - 4.5 mL Sides', index=False, header=False)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("print_type", choices=['STANDARD', 'SERONET', 'SERUM'], help="Choose print type")
    args = parser.parse_args()

    # Rename Printing Log: Study - Kit Type
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=0)
            .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max', 'Study': 'Kit Type'})
            .drop(1))

    # Input: print_type & box_start & box_end (e.g.: SERUM 63 66)
    print_type = args.print_type
    
    # Printing Log: Find box range
    box_start, box_end = get_box_range(plog, print_type)
    workbook_name = f"{print_type} {box_start}-{box_end}"

    # Print Planning: Retrieve sample IDs
    assigned_sample_ids = get_sample_ids(print_type, box_start, box_end)

    # Output workbook
    output_file = get_output_file(print_type, workbook_name)

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        if print_type == 'SERUM':
            serum_worksheet(writer, assigned_sample_ids, box_start, box_end)
    
    print("Done! Output workbook is under PVI/ Print Shop/ Tube Printing/ Future Sheets.")

