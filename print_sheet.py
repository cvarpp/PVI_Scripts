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
        filtered_plog = plog[(plog['Kit Type'] == 'SERONET') & (plog['PBMCs'].astype(str).str.strip().str.lower() == 'no')]
    elif print_type == 'SERUM':
        filtered_plog = plog[(plog['Kit Type'] == 'SERUM')]
    elif print_type == 'STANDARD':
        filtered_plog = plog[(plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'].astype(str).str.strip().str.lower() == 'no')]
    elif print_type == 'SERONETPBMC':
        filtered_plog = plog[(plog['Kit Type'].isin(['SERONET', 'MIT (PBMCS)'])) & (plog['PBMCs'].astype(str).str.strip().str.lower() == 'yes')]
    elif print_type == 'STANDARDPBMC':
        filtered_plog = plog[(plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'].astype(str).str.strip().str.lower() == 'yes')]

    filtered_plog = filtered_plog.sort_values(by='Box Max', ascending=False).iloc[0]
    recent_box_max = filtered_plog['Box Max']

    box_range_mapping = {
    'SERONET': 6,
    'SERUM': 4,
    'STANDARD': 6,
    'SERONETPBMC': 32,
    'STANDARDPBMC': 32,
    }
    box_start = recent_box_max + 1
    box_end = recent_box_max + box_range_mapping[print_type]
    
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
            
            if print_type == 'SERONET':
                if 'READ ME' in sheet_name:
                    box_numbers_range = range(box_start, box_end + 1)
                    sheet_data.iloc[1:7, 7] = box_numbers_range
                    sheet_data.iloc[7:25, 7] = assigned_sample_ids
                    sheet_data.iloc[1:19, 12] = assigned_sample_ids
                    sheet_data.iloc[1:19, 13] = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                if '4.5 mL Tops' in sheet_name:
                    sheet_data.iloc[:18, 2] = assigned_sample_ids
                if '4.5 mL Sides' in sheet_name:
                    sheet_data.iloc[:18, 1] = assigned_sample_ids
                if '1-Box Top round1' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[:6]
                    sheet_data.iloc[:90, 1] = [sid for sid in assigned_sample_ids[:6] for _ in range(15)]
                if '2-Box Side round1' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[:6]
                    sheet_data.iloc[:90, 11] = [sid for sid in assigned_sample_ids[:6] for _ in range(15)]
                    sheet_data[1] = sheet_data[12] + " " + sheet_data[11]
                if '3-Box Top round2' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[6:12]
                    sheet_data.iloc[:90, 1] = [sid for sid in assigned_sample_ids[6:12] for _ in range(15)]
                if '4-Box Side round2' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[6:12]
                    sheet_data.iloc[:90, 11] = [sid for sid in assigned_sample_ids[6:12] for _ in range(15)]
                    sheet_data[1] = sheet_data[12] + " " + sheet_data[11]
                if '5-Box Top round3' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[12:18]
                    sheet_data.iloc[:90, 1] = [sid for sid in assigned_sample_ids[12:18] for _ in range(15)]
                if '6-Box Side round3' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[12:18]
                    sheet_data.iloc[:90, 11] = [sid for sid in assigned_sample_ids[12:18] for _ in range(15)]
                    sheet_data[1] = sheet_data[12] + " " + sheet_data[11]
                if '7-Kits' in sheet_name:
                    sheet_data.iloc[:90, 1] = [sid for i in range(0, len(assigned_sample_ids), 2) for sid in [assigned_sample_ids[i], assigned_sample_ids[i + 1]] * 5]

            if print_type == 'SERUM':
                if sheet_name == '1 - Kits':
                    sheet_data.iloc[1:25, 13] = assigned_sample_ids
                    sheet_data.iloc[1:25, 14] = [box for box in range(box_start, box_end + 1) for _ in range(6)]
                    sheet_data.iloc[0:48, 1] = [sid for sid in assigned_sample_ids for _ in range(2)]
                if sheet_name == '2 - Tops':
                    sheet_data.iloc[0:168, 1] = [sid for sid in assigned_sample_ids for _ in range(7)]
                if sheet_name == '3 - Sides':
                    sheet_data.iloc[0:168, 2] = [sid for sid in assigned_sample_ids for _ in range(7)]
                if sheet_name == '4 - 4.5 mL Tops':
                    sheet_data.iloc[0:24, 2] = assigned_sample_ids
                if sheet_name == '5 - 4.5 mL Sides':
                    sheet_data.iloc[0:24, 1] = assigned_sample_ids
                    
            if print_type == 'STANDARD':
                if 'READ ME' in sheet_name:
                    box_numbers_range = range(box_start, box_end + 1)
                    sheet_data.iloc[1:7, 7] = box_numbers_range
                    sheet_data.iloc[7:25, 7] = assigned_sample_ids
                    sheet_data.iloc[1:19, 12] = assigned_sample_ids
                    sheet_data.iloc[1:19, 13] = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                if '1-Box Top round1' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[:6]
                    sheet_data.iloc[:90, 1] = [sid for sid in assigned_sample_ids[:6] for _ in range(15)]
                if '2-Box Side round1' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[:6]
                    sheet_data.iloc[:90, 11] = [sid for sid in assigned_sample_ids[:6] for _ in range(15)]
                    sheet_data[1] = sheet_data[12] + " " + sheet_data[11]
                if '3-Box Top round2' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[6:12]
                    sheet_data.iloc[:90, 1] = [sid for sid in assigned_sample_ids[6:12] for _ in range(15)]
                if '4-Box Side round2' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[6:12]
                    sheet_data.iloc[:90, 11] = [sid for sid in assigned_sample_ids[6:12] for _ in range(15)]
                    sheet_data[1] = sheet_data[12] + " " + sheet_data[11]
                if '5-Box Top round3' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[12:]
                    sheet_data.iloc[:90, 1] = [sid for sid in assigned_sample_ids[12:] for _ in range(15)]
                if '6-Box Side round3' in sheet_name:
                    sheet_data.iloc[:6, 7] = assigned_sample_ids[12:]
                    sheet_data.iloc[:90, 11] = [sid for sid in assigned_sample_ids[12:] for _ in range(15)]
                    sheet_data[1] = sheet_data[12] + " " + sheet_data[11]
                if '7-Kits' in sheet_name:
                    sheet_data.iloc[:90, 1] = [sid for i in range(0, len(assigned_sample_ids), 2) for sid in [assigned_sample_ids[i], assigned_sample_ids[i + 1]] * 5]
            
            if print_type == 'SERONETPBMC':
                if 'Instructions' in sheet_name:
                    sheet_data.iloc[1:97, 21] = assigned_sample_ids
                    box_numbers = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                    sheet_data.iloc[1:97, 22] = box_numbers
                    sheet_data.iloc[4, 15] = "Box #" + str(box_start)
                    sheet_data.iloc[9, 15] = "Box #" + str(box_start + 5)
                    sheet_data.iloc[14, 15] = "Box #" + str(box_start + 10)
                    sheet_data.iloc[4, 18] = "Box #" + str(box_start + 15)
                    sheet_data.iloc[9, 18] = "Box #" + str(box_start + 20)
                    sheet_data.iloc[14, 18] = "Box #" + str(box_start + 25)
                    sheet_data.iloc[4:7, 16] = assigned_sample_ids[:3]
                    sheet_data.iloc[9:12, 16] = assigned_sample_ids[15:18]
                    sheet_data.iloc[14:17, 16] = assigned_sample_ids[30:33]
                    sheet_data.iloc[4:7, 19] = assigned_sample_ids[45:48]
                    sheet_data.iloc[9:12, 19] = assigned_sample_ids[60:63]
                    sheet_data.iloc[14:17, 19] = assigned_sample_ids[75:78]
                if 'PBMC Tops' in sheet_name:
                    sheet_data.iloc[:96, 1] = assigned_sample_ids
                if 'PBMC Sides' in sheet_name:
                    sheet_data.iloc[:96, 1] = assigned_sample_ids
            
            if print_type == 'STANDARDPBMC':
                if 'Instructions' in sheet_name:
                    sheet_data.iloc[1:97, 21] = assigned_sample_ids
                    box_numbers = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                    sheet_data.iloc[1:97, 22] = box_numbers
                    sheet_data.iloc[4, 15] = "Box #" + str(box_start)
                    sheet_data.iloc[9, 15] = "Box #" + str(box_start + 5)
                    sheet_data.iloc[14, 15] = "Box #" + str(box_start + 10)
                    sheet_data.iloc[4, 18] = "Box #" + str(box_start + 15)
                    sheet_data.iloc[9, 18] = "Box #" + str(box_start + 20)
                    sheet_data.iloc[14, 18] = "Box #" + str(box_start + 25)
                    sheet_data.iloc[4:7, 16] = assigned_sample_ids[:3]
                    sheet_data.iloc[9:12, 16] = assigned_sample_ids[15:18]
                    sheet_data.iloc[14:17, 16] = assigned_sample_ids[30:33]
                    sheet_data.iloc[4:7, 19] = assigned_sample_ids[45:48]
                    sheet_data.iloc[9:12, 19] = assigned_sample_ids[60:63]
                    sheet_data.iloc[14:17, 19] = assigned_sample_ids[75:78]
                if 'PBMC Tops' in sheet_name:
                    sheet_data.iloc[:288, 2] = [sid for sid in assigned_sample_ids for _ in range(3)]
                if 'PBMC Sides' in sheet_name:
                    sheet_data.iloc[:288, 2] = [sid for sid in assigned_sample_ids for _ in range(3)]

            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False, header=False)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-seronet_full', type=int, default=0, help='Number of SERONET FULL rounds')
    parser.add_argument('-serum', type=int, default=0, help='Number of SERUM rounds')
    parser.add_argument('-standard', type=int, default=0, help='Number of STANDARD rounds')
    parser.add_argument('-seronet_pbmc', type=int, default=0, help='Number of SERONET PBMC rounds')
    parser.add_argument('-standard_pbmc', type=int, default=0, help='Number of STANDARD PBMC rounds')
    args = parser.parse_args()

    round_counts = {
        'SERONET_FULL': args.seronet_full,
        'SERUM': args.serum,
        'STANDARD': args.standard,
        'SERONETPBMC': args.seronet_pbmc,
        'STANDARDPBMC': args.standard_pbmc,
    }

    for print_type, rounds in round_counts.items():
        if rounds > 0:
            # Box range
            box_start, box_end = get_box_range(print_type)

            # Map print_type to workbook_name & template_file & output_path
            print_type_mapping = {
                'SERONET_FULL': ('Seronet Full', 'SERONET FULL/SERONET FULL Template.xlsx', 'Future Sheets/SERONET FULL'),
                'SERUM': ('Serum', 'SERUM/SERUM Template.xlsx', 'Future Sheets/SERUM'),
                'STANDARD': ('Standard', 'STANDARD/STANDARD Template.xlsx', 'Future Sheets/STANDARD'),
                'SERONETPBMC': ('Seronet Full', 'SERONET FULL/SERONET FULL PBMC Template.xlsx', 'Future Sheets/SERONET FULL'),
                'STANDARDPBMC': ('Standard', 'STANDARD/STANDARD PBMC Template.xlsx', 'Future Sheets/STANDARD')
            }
            sheet_name, template_file, output_folder = print_type_mapping[print_type]

            # Sample IDs within box range
            assigned_sample_ids = get_sample_ids(sheet_name, box_start, box_end)

            # Output
            workbook_name = f"{sheet_name.upper()} {'PBMC ' if 'PBMC' in print_type else ''}{box_start}-{box_end} from scripts"
            template_path = os.path.join(util.tube_print, 'Future Sheets', template_file)
            output_path = os.path.join(util.tube_print, output_folder, f"{workbook_name}.xlsx")

            generate_workbook(assigned_sample_ids, box_start, box_end, workbook_name, template_path, output_path, print_type)

            print(f"'{workbook_name}' workbook generated in {output_folder}.")

