import pandas as pd
import numpy as np
import argparse
import util
import os
import warnings


def get_box_range(print_type, round_num):
    # Rename Printing Log: Study - Kit Type
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=0)
            .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max', 'Study': 'Kit Type'})
            .drop(1))
    plog[['Box Min', 'Box Max']] = plog[['Box Min', 'Box Max']].apply(pd.to_numeric, errors='coerce').dropna()

    filter = {
        'SERONET': (plog['Kit Type'] == 'SERONET') & (plog['PBMCs'].str.strip().str.lower() == 'no'),
        'SERUM': (plog['Kit Type'] == 'SERUM'),
        'STANDARD': (plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'].str.strip().str.lower() == 'no'),
        'SERONETPBMC': (plog['Kit Type'].isin(['SERONET', 'MIT (PBMCS)'])) & (plog['PBMCs'].str.strip().str.lower() == 'yes'),
        'STANDARDPBMC': (plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'].str.strip().str.lower() == 'yes'),
        'APOLLO':(plog['Kit Type'] == 'APOLLO') & (plog['PBMCs'].str.strip().str.lower() == 'no')
    }

    if print_type not in filter:
        print(f"{print_type} is not a valid option. Exiting...")
        exit(1)

    filtered_plog = plog[filter[print_type]].sort_values(by='Box Max', ascending=False).iloc[0]
    recent_box_max = filtered_plog['Box Max']

    box_range_mapping = {
        'SERONET': 6,
        'SERUM': 4,
        'STANDARD': 6,
        'SERONETPBMC': 32,
        'STANDARDPBMC': 32,
        'APOLLO': 5
    }
    box_start = recent_box_max + 1 + round_num * box_range_mapping[print_type]
    box_end = recent_box_max + box_range_mapping[print_type] + round_num * box_range_mapping[print_type]

    return int(box_start), int(box_end)


def get_sample_ids(sheet_name, box_start, box_end):
    print_planning_path = os.path.join(util.tube_print, 'Print Planning.xlsx')
    print_planning = pd.read_excel(print_planning_path, sheet_name=sheet_name)
    box_range = print_planning['Box ID'].between(box_start, box_end)
    sample_ids = print_planning.loc[box_range, 'Sample ID'].to_numpy()

    return sample_ids


def generate_workbook(assigned_sample_ids, box_start, box_end, sheet_name, template_folder, output_path, print_type):
    template_path = os.path.join(util.tube_print, 'Future Sheets', template_folder)
    template = pd.read_excel(template_path, sheet_name=None, header=None)

    with pd.ExcelWriter(output_path) as writer:
        for sheet_name, sheet_data in template.items():
            
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
            
            elif print_type in ('SERONET', 'STANDARD'):
                if 'READ ME' in sheet_name:
                    box_numbers_range = range(box_start, box_end + 1)
                    sheet_data.iloc[1:7, 7] = box_numbers_range
                    sheet_data.iloc[7:25, 7] = assigned_sample_ids
                    sheet_data.iloc[1:19, 12] = assigned_sample_ids
                    sheet_data.iloc[1:19, 13] = [box for box in range(box_start, box_end + 1) for _ in range(3)]
                
                for round_num in range(1, 4):
                    top_sheet = f'{round_num * 2 - 1}-Box Top round{round_num}'
                    side_sheet = f'{round_num * 2}-Box Side round{round_num}'
                    idx_start = (round_num - 1) * 6
                    idx_end = round_num * 6
                    
                    if top_sheet in sheet_name:
                        sheet_data.iloc[:6, 7] = assigned_sample_ids[idx_start:idx_end]
                        sheet_data.iloc[:90, 2] = [sid for sid in assigned_sample_ids[idx_start:idx_end] for _ in range(15)]
                    
                    if side_sheet in sheet_name:
                        sheet_data.iloc[:6, 7] = assigned_sample_ids[idx_start:idx_end]
                        sheet_data.iloc[:90, 11] = [sid for sid in assigned_sample_ids[idx_start:idx_end] for _ in range(15)]
                        sheet_data[2] = sheet_data[12] + " " + sheet_data[11]
                
                if '7-Kits' in sheet_name:
                    sheet_data.iloc[:90, 1] = [sid for i in range(0, len(assigned_sample_ids), 2) for sid in [assigned_sample_ids[i], assigned_sample_ids[i + 1]] * 5]

            elif print_type in ('SERONETPBMC', 'STANDARDPBMC'):
                if 'Instructions' in sheet_name:
                    sheet_data.iloc[1:97, 21] = assigned_sample_ids
                    sheet_data.iloc[1:97, 22] = [box for box in range(box_start, box_end + 1) for _ in range(3)]
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

                if print_type == 'SERONETPBMC':
                    if 'PBMC Tops' in sheet_name or '4.5mL Tops' in sheet_name:
                        sheet_data.iloc[:96, 2] = assigned_sample_ids
                    if 'PBMC Sides' in sheet_name or '4.5mL Sides' in sheet_name:
                        sheet_data.iloc[:96, 1] = assigned_sample_ids
        
                if print_type == 'STANDARDPBMC':
                    if 'PBMC Tops' in sheet_name or 'PBMC Sides' in sheet_name:
                        sheet_data.iloc[:288, 2] = [sid for sid in assigned_sample_ids for _ in range(3)]

            elif print_type == 'APOLLO':
                sids = assigned_sample_ids
                sid_kits = []
                for sid in sids:
                    sid_kits.extend([sid] * 5)
                counts = list(np.concatenate([2 * np.arange(5) + 1, 2 * np.arange(5) + 2])) * 9
                count_adders = np.arange(108, step=12)
                count_adders2 = []
                for count_adder in count_adders:
                    count_adders2.extend([count_adder] * 10)
                final_index = np.array(counts) + np.array(count_adders2)
                kit_df = pd.DataFrame({'Column A': final_index, 'Sample ID': sid_kits, 'Kit Type': 'APOLLO'}).sort_values(by='Column A')
                kit_df.to_excel(writer, sheet_name='1 - Kits', index=False, header=False)
                sid_aliquots = []
                for sid in sids:
                    sid_aliquots.extend([sid] * 16)
                for x in range(3):
                    round_num = x + 1
                    top_idx = round_num * 2
                    side_idx = top_idx + 1
                    local_sids = sid_aliquots[x * 16 * 6:(x + 1) * 16 * 6]
                    top_index = np.concatenate([np.arange(48) + 1, np.arange(48) + 1 + 78])
                    side_index = np.concatenate([np.arange(48) + 1, np.arange(48) + 1 + 50])
                    top_df = pd.DataFrame({'Column A': top_index, 'Blank': '', 'Sample ID': local_sids, 'Sample Type': 'Serum'})
                    top_df.to_excel(writer, sheet_name=f'{top_idx} - Tops {round_num}', index=False, header=False)
                    side_writers = ['Serum {}'.format(sid) for sid in local_sids]
                    side_df = pd.DataFrame({'Column A': side_index, 'Blank': '', 'Sample ID': side_writers, 'Kit Type': 'APOLLO'})
                    side_df.to_excel(writer, sheet_name=f'{side_idx} - Sides {round_num}', index=False, header=False)

            # sheet_data.to_excel(writer, sheet_name=sheet_name, index=False, header=False)


if __name__ == '__main__':
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    parser = argparse.ArgumentParser()
    parser.add_argument('-seronet_full', type=int, default=0, help='Number of SERONET FULL rounds')
    parser.add_argument('-serum', type=int, default=0, help='Number of SERUM rounds')
    parser.add_argument('-standard', type=int, default=0, help='Number of STANDARD rounds')
    parser.add_argument('-seronet_pbmc', type=int, default=0, help='Number of SERONET PBMC rounds')
    parser.add_argument('-standard_pbmc', type=int, default=0, help='Number of STANDARD PBMC rounds')
    parser.add_argument('-apollo', type=int, default=0, help='Number of APOLLO rounds')
    args = parser.parse_args()

    round_counts = {
        'SERONET': args.seronet_full,
        'SERUM': args.serum,
        'STANDARD': args.standard,
        'SERONETPBMC': args.seronet_pbmc,
        'STANDARDPBMC': args.standard_pbmc,
        'APOLLO': args.apollo,
    }

    for print_type, rounds in round_counts.items():
        for round_num in range(0, rounds):
            box_start, box_end = get_box_range(print_type, round_num)

            print_type_mapping = {
                'SERONET': ('Seronet Full', 'SERONET FULL/SERONET FULL Template.xlsx', 'Future Sheets/SERONET FULL'),
                'SERUM': ('Serum', 'SERUM/SERUM Template.xlsx', 'Future Sheets/SERUM'),
                'STANDARD': ('Standard', 'STANDARD/STANDARD Template.xlsx', 'Future Sheets/STANDARD'),
                'SERONETPBMC': ('Seronet Full', 'SERONET FULL/SERONET PBMC Template.xlsx', 'Future Sheets/SERONET FULL'),
                'STANDARDPBMC': ('Standard', 'STANDARD/STANDARD PBMC Template.xlsx', 'Future Sheets/STANDARD'),
                'APOLLO': ('APOLLO', 'APOLLO/APOLLO Template.xlsx', 'Future Sheets/APOLLO'),
            }
            sheet_name, template_file, output_folder = print_type_mapping[print_type]

            assigned_sample_ids = get_sample_ids(sheet_name, box_start, box_end)

            if not assigned_sample_ids.any():
                print(f"Need more assigned sample IDs for {print_type}. Skipping round {round_num + 1}.")
                continue
            
            # Output
            workbook_name = f"{sheet_name.upper()} {'PBMC ' if 'PBMC' in print_type else ''}{box_start}-{box_end} by scripts"
            template_path = os.path.join(util.tube_print, 'Future Sheets', template_file)
            output_path = os.path.join(util.tube_print, output_folder, f"{workbook_name}.xlsx")

            generate_workbook(assigned_sample_ids, box_start, box_end, workbook_name, template_path, output_path, print_type)
            print(f"'{workbook_name}' workbook generated in {output_folder}.")
            