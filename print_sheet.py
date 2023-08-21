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
        filtered_plog = plog[(plog['Kit Type'] == 'SERONET') & (plog['PBMCs'].isin(['No', 'no']))]
    elif print_type == 'SERUM':
        filtered_plog = plog[(plog['Kit Type'] == 'SERUM')]
    elif print_type == 'STANDARD':
        filtered_plog = plog[(plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'].isin(['No', 'no']))]
    elif print_type == 'SERONETPBMC':
        filtered_plog = plog[(plog['Kit Type'].isin(['SERONET', 'MIT (PBMCs)'])) & (plog['PBMCs'].isin(['Yes', 'yes']))]
    elif print_type == 'STANDARDPBMC':
        filtered_plog = plog[(plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'].isin(['Yes', 'yes']))]

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
    template = pd.read_excel(template_path, sheet_name=None)

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for sheet_name, sheet_data in template.items():
            if print_type == 'SERONET':
                if 'READ ME' in sheet_name:
                    sheet_data.loc[2:19, 'M'] = assigned_sample_ids
                    box_numbers_repeated = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                    sheet_data.loc[2:19, 'N'] = box_numbers_repeated
                elif '4.5 mL Tops' in sheet_name:
                    sheet_data.loc[0:17, 'C'] = assigned_sample_ids
                elif '4.5 mL Sides' in sheet_name:
                    sheet_data.loc[0:17, 'B'] = assigned_sample_ids_repeated
                elif '1-Box Top round1' in sheet_name:
                    assigned_sample_ids_1 = assigned_sample_ids[:6]
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_1
                    assigned_sample_ids_1_repeated = [sid for sid in assigned_sample_ids_1 for _ in range(15)]
                    sheet_data.loc[1:90, 'B'] = assigned_sample_ids_1_repeated
                elif '2-Box Side round1' in sheet_name:
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_1
                    sheet_data.loc[1:90, 'L'] = assigned_sample_ids_1_repeated
                    sheet_data['B'] = sheet_data['M'] + " " + sheet_data['L']
                elif '3-Box Top round2' in sheet_name:
                    assigned_sample_ids_2 = assigned_sample_ids[6:12]
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_2
                    assigned_sample_ids_2_repeated = [sid for sid in assigned_sample_ids_2 for _ in range(15)]
                    sheet_data.loc[1:90, 'B'] = assigned_sample_ids_2_repeated
                elif '4-Box Side round2' in sheet_name:
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_2
                    sheet_data.loc[1:90, 'L'] = assigned_sample_ids_2_repeated
                    sheet_data['B'] = sheet_data['M'] + " " + sheet_data['L']
                elif '5-Box Top round3' in sheet_name:
                    assigned_sample_ids_3 = assigned_sample_ids[12:18]
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_3
                    assigned_sample_ids_3_repeated = [sid for sid in assigned_sample_ids_3 for _ in range(15)]
                    sheet_data.loc[1:90, 'B'] = assigned_sample_ids_3_repeated
                elif '6-Box Side round3' in sheet_name:
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_3
                    sheet_data.loc[1:90, 'L'] = assigned_sample_ids_3_repeated
                    sheet_data['B'] = sheet_data['M'] + " " + sheet_data['L']
                elif '7-Kits' in sheet_name:
                    # assigned_sample_ids_repeated_total = []
                    # for i in range(0, len(assigned_sample_ids), 2):
                    #     assigned_sample_ids_repeated_total.extend([assigned_sample_ids[i], assigned_sample_ids[i + 1]] * 5)
                    assigned_sample_ids_repeated_total = [sid for i in range(0, len(assigned_sample_ids), 2) for sid in [assigned_sample_ids[i], assigned_sample_ids[i + 1]] * 5]
                    sheet_data.loc[0:89, 'B'] = assigned_sample_ids_repeated_total

            elif print_type == 'SERUM':
                if sheet_name == '1 - Kits':
                    box_numbers_repeated = [box for box in range(box_start, box_end + 1) for _ in range(6)]
                    assigned_sample_ids_repeated = [sid for sid in assigned_sample_ids for _ in range(2)]
                    sheet_data.loc[1:24, 'N'] = assigned_sample_ids
                    sheet_data.loc[1:24, 'O'] = box_numbers_repeated
                    sheet_data.loc[0:47, 'B'] = assigned_sample_ids_repeated
                elif sheet_name == '2 - Tops':
                    assigned_sample_ids_repeated = [sid for sid in assigned_sample_ids for _ in range(7)]
                    sheet_data.loc[0:167, 'B'] = assigned_sample_ids_repeated
                elif sheet_name == '3 - Sides':
                    sheet_data.loc[0:167, 'C'] = assigned_sample_ids_repeated
                elif sheet_name == '4 - 4.5 mL Tops':
                    sheet_data.loc[0:23, 'C'] = assigned_sample_ids
                elif sheet_name == '5 - 4.5 mL Sides':
                    sheet_data.loc[0:23, 'B'] = assigned_sample_ids
                    
            elif print_type == 'STANDARD':
                if 'READ ME' in sheet_name:
                    sheet_data.loc[2:19, 'M'] = assigned_sample_ids
                    box_numbers_repeated = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                    sheet_data.loc[2:19, 'N'] = box_numbers_repeated
                elif '4.5 mL Tops' in sheet_name:
                    sheet_data.loc[0:17, 'C'] = assigned_sample_ids
                elif '4.5 mL Sides' in sheet_name:
                    sheet_data.loc[0:17, 'B'] = assigned_sample_ids_repeated
                elif '1-Box Top round1' in sheet_name:
                    assigned_sample_ids_1 = assigned_sample_ids[:6]
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_1
                    assigned_sample_ids_1_repeated = [sid for sid in assigned_sample_ids_1 for _ in range(15)]
                    sheet_data.loc[1:90, 'B'] = assigned_sample_ids_1_repeated
                elif '2-Box Side round1' in sheet_name:
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_1
                    sheet_data.loc[1:90, 'L'] = assigned_sample_ids_1_repeated
                    sheet_data['B'] = sheet_data['M'] + " " + sheet_data['L']
                elif '3-Box Top round2' in sheet_name:
                    assigned_sample_ids_2 = assigned_sample_ids[6:12]
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_2
                    assigned_sample_ids_2_repeated = [sid for sid in assigned_sample_ids_2 for _ in range(15)]
                    sheet_data.loc[1:90, 'B'] = assigned_sample_ids_2_repeated
                elif '4-Box Side round2' in sheet_name:
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_2
                    sheet_data.loc[1:90, 'L'] = assigned_sample_ids_2_repeated
                    sheet_data['B'] = sheet_data['M'] + " " + sheet_data['L']
                elif '5-Box Top round3' in sheet_name:
                    assigned_sample_ids_3 = assigned_sample_ids[12:18]
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_3
                    assigned_sample_ids_3_repeated = [sid for sid in assigned_sample_ids_3 for _ in range(15)]
                    sheet_data.loc[1:90, 'B'] = assigned_sample_ids_3_repeated
                elif '6-Box Side round3' in sheet_name:
                    sheet_data.loc[1:6, 'H'] = assigned_sample_ids_3
                    sheet_data.loc[1:90, 'L'] = assigned_sample_ids_3_repeated
                    sheet_data['B'] = sheet_data['M'] + " " + sheet_data['L']
                elif '7-Kits' in sheet_name:
                    assigned_sample_ids_repeated_total = [sid for i in range(0, len(assigned_sample_ids), 2) for sid in [assigned_sample_ids[i], assigned_sample_ids[i + 1]] * 5]
                    sheet_data.loc[0:89, 'B'] = assigned_sample_ids_repeated_total
            
            elif print_type == 'SERONETPBMC':
                if 'Instructions' in sheet_name:
                    sheet_data.loc[1:96, 'V'] = assigned_sample_ids
                    box_numbers = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                    sheet_data.loc[1:96, 'W'] = box_numbers
                elif 'PBMC Tops' in sheet_name:
                    sheet_data.loc[0:95, 'B'] = assigned_sample_ids
                elif 'PBMC Sides' in sheet_name:
                    sheet_data.loc[0:95, 'B'] = assigned_sample_ids
            
            elif print_type == 'STANDARDPBMC':
                if 'Instructions' in sheet_name:
                    sheet_data.loc[1:96, 'V'] = assigned_sample_ids
                    box_numbers = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                    sheet_data.loc[1:96, 'W'] = box_numbers
                elif 'PBMC Tops' in sheet_name:
                    sheet_data.loc[0:95, 'B'] = assigned_sample_ids
                elif 'PBMC Sides' in sheet_name:
                    sheet_data.loc[0:95, 'B'] = assigned_sample_ids

            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)


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
    workbook_name = f"{sheet_name.upper()} {'PBMC ' if 'PBMC' in print_type else ''}{box_start}-{box_end} test"
    template_path = os.path.join(util.tube_print, 'Future Sheets', template_file)
    output_path = os.path.join(util.tube_print, output_folder, f"{workbook_name}.xlsx")

    generate_workbook(assigned_sample_ids, box_start, box_end, workbook_name, template_path, output_path, print_type)

    print(f"'{workbook_name}' workbook generated in {output_folder}.")

    # test use only
    print(box_start, box_end)
    print(assigned_sample_ids)
