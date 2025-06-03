import pandas as pd
import numpy as np
import argparse
import util
import os
import warnings

def parse_input():
    parser = argparse.ArgumentParser()
    parser.add_argument('--seronet_cam', type=int, default=0, help='Number of SERONET CAM rounds')
    parser.add_argument('--seronet_rtc', type=int, default=0, help='Number of SERONET RTC rounds')
    parser.add_argument('--serum', type=int, default=0, help='Number of SERUM rounds')
    parser.add_argument('--standard', type=int, default=0, help='Number of STANDARD rounds')
    parser.add_argument('--seronet_pbmc', type=int, default=0, help='Number of SERONET PBMC rounds')
    parser.add_argument('--standard_pbmc', type=int, default=0, help='Number of STANDARD PBMC rounds')
    parser.add_argument('--apollo', type=int, default=0, help='Number of APOLLO rounds')
    parser.add_argument('--nps', type=int, default=0, help='Number of NPS rounds')
    parser.add_argument('--cells', type=int, default=0, help='Number of CELLS rounds')
    parser.add_argument('--cells_pbmc', type=int, default=0, help='Number of CELLS PBMC rounds')
    args = parser.parse_args()

    session_counts = {
        'SERONET': args.seronet_cam + args.seronet_rtc,
        'SERONET_RTC': args.seronet_rtc,
        'SERUM': args.serum,
        'STANDARD': args.standard,
        'SERONETPBMC': args.seronet_pbmc,
        'STANDARDPBMC': args.standard_pbmc,
        'APOLLO': args.apollo,
        'NPS': args.nps,
        'CELLS': args.cells,
        'CELLSPBMC': args.cells_pbmc,
    }
    return session_counts

def last_printed_box(print_type):
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=0)
            .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max', 'Study': 'Kit Type'})
            .drop(1))
    plog['PBMCs'] = plog['PBMCs'].str.strip().str.lower()
    plog[['Box Min', 'Box Max']] = plog[['Box Min', 'Box Max']].apply(pd.to_numeric, errors='coerce').dropna()

    filter = {
        'SERONET': (plog['Kit Type'] == 'SERONET') & (plog['PBMCs'] == 'no'),
        'SERONET_RTC': (plog['Kit Type'] == 'SERONET') & (plog['PBMCs'] == 'no'),
        'SERUM': (plog['Kit Type'] == 'SERUM'),
        'STANDARD': (plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'] == 'no'),
        'SERONETPBMC': (plog['Kit Type'].isin(['SERONET', 'MIT (PBMCS)'])) & (plog['PBMCs'] == 'yes'),
        'STANDARDPBMC': (plog['Kit Type'] == 'STANDARD') & (plog['PBMCs'] == 'yes'),
        'APOLLO': (plog['Kit Type'] == 'APOLLO'),
        'NPS': (plog['Kit Type'] == 'NPS'),
        'CELLS': (plog['Kit Type'] == 'CELLS'),
        'CELLSPBMC': (plog['Kit Type'] == 'CELLS') & (plog['PBMCs'] == 'yes'),
    }
    if print_type not in filter:
        print(f"{print_type} is not a valid option. Exiting...")
        exit(1)
    recent_box_max = plog[filter[print_type]].sort_values(
        by='Box Max', ascending=False).iloc[0]['Box Max']

    return recent_box_max

class PrintSession():
    def __init__(self, kit_type, templates, session_num,):
        self.kit_type = kit_type
        self.kit_info = pd.read_excel(templates, sheet_name='Tube Counts', index_col='Kit Type').loc[kit_type, :]
        self.box_start = int(recent_box_max) + 1 + session_num * self.kit_info['Boxes per Print Session']
        self.box_end = self.box_start + self.kit_info['Boxes per Print Session'] - 1
        self.fiveml = self.kit_info['4.5 mL'] > 0
        self.pbmc = self.kit_info['PBMC'] > 0
        self.kits = self.kit_info['Collection Tubes'] > 0
        if kit_type in templates.sheet_names:
            self.instructions = pd.read_excel(templates, sheet_name=kit_type)
        else:
            self.instructions = None
        if self.kits or kit_type == 'NPS':
            self.bradys = pd.read_excel(templates, sheet_name=f"{kit_type.replace('_RTC', '')} Brady")
        else:
            self.bradys = None
        self.valid_sid_set = self.get_sample_ids()
        self.make_output_path()

    def get_sample_ids(self):
        print_planning_path = os.path.join(util.tube_print, 'Print Planning.xlsx')
        print_planning = pd.read_excel(print_planning_path, sheet_name=self.kit_info['Print Planning Sheet'])
        box_range = print_planning['Box ID'].between(self.box_start, self.box_end)
        self.sample_ids = print_planning.loc[box_range, 'Sample ID'].to_numpy()
        return self.sample_ids.size == self.kit_info['Boxes per Print Session'] * self.kit_info['IDs per Box']

    def make_output_path(self):
        output_prefix = "3CPTs " if self.kit_type == 'SERONET_RTC' else "2CPTs " if self.kit_type == 'SERONET' else ""
        workbook_name = f"{output_prefix}{print_type} {self.box_start}-{self.box_end}".strip()
        self.output_path = os.path.join(util.tube_print, 'Future Sheets',
                                        self.kit_info['Future Sheets Folder'], f"{workbook_name}.xlsx")

    def write_workbook(self):
        self.future_workbook = {}
        if self.instructions is not None:
            self.write_instructions()
        if self.bradys is not None:
            self.write_bradys()
        if self.kits:
            self.write_kits()
        for round_num in range(self.kit_info['Rounds per Print Session']):
            sid_count = self.kit_info['IDs per Round']
            sids = self.sample_ids[round_num * sid_count:(round_num + 1) * sid_count]
            if self.pbmc:
                self.pbmc_round(round_num, sids)
            else:
                self.twoml_round(round_num, sids)
        if self.fiveml:
            self.write_fivemls()
        with pd.ExcelWriter(self.output_path) as writer:
            for sheet_name, dataframe in self.future_workbook.items():
                include_header = 'Brady' in sheet_name
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False, header=include_header)

    def write_instructions(self):
        sid_replacer = {f'Sample {i + 1}': sid for i, sid in enumerate(self.sample_ids)}
        box_replacer = {f'Box {i + 1}': f'Box {self.box_start + i}' for i in range(self.kit_info['Boxes per Print Session'])}
        self.future_workbook['Instructions'] = self.instructions.replace(sid_replacer).replace(box_replacer)

    def write_bradys(self):
        sid_replacer = {f'Sample {i + 1}': sid for i, sid in enumerate(self.sample_ids)}
        box_replacer = {f'Box {i + 1}': self.box_start + i for i in range(self.kit_info['Boxes per Print Session'])}
        self.future_workbook['0 - Brady Labels'] = self.bradys.replace(sid_replacer).replace(box_replacer)

    def write_kits(self):
        kit_data = pd.DataFrame()
        sids = self.sample_ids
        tube_count = self.kit_info['Collection Tubes']
        aliquot_ids = np.repeat(sids, tube_count)
        aliquot_idxes = (np.repeat(range(len(sids) // 2), tube_count * 2) * 12 + 1 +
                         np.tile(range(tube_count), len(sids)) * 2 +
                         np.tile(np.repeat(range(2), tube_count), len(sids) // 2))
        kit_data['c1'] = aliquot_idxes
        kit_data['c2'] = aliquot_ids
        kit_data['c3'] = self.kit_type
        self.future_workbook['Collection Kits'] = kit_data.sort_values(by='c1')

    def pbmc_round(self, round_num, sids):
        top_sheet = '{} - Tops {}'.format(round_num * 2 + 1, round_num + 1)
        side_sheet = '{} - Sides {}'.format(round_num * 2 + 2, round_num + 1)
        aliquot_ids = np.repeat(sids, self.kit_info['PBMC'])
        top_count = self.kit_info['Tubes in Top Rack']
        top_racks = self.kit_info['Top Racks per Round']
        top_idxes = np.repeat([x for x in range(top_racks)], top_count) * 96 + 1
        top_idxes += np.array([x for x in range(top_count)] * top_racks)
        side_count = self.kit_info['Tubes in Side Rack']
        side_racks = self.kit_info['Side Racks per Round']
        side_idxes = np.repeat(range(side_racks), side_count) * 50 + 1
        side_idxes += np.array([x for x in range(side_count)] * side_racks)
        side_labels = ['PBMC' for _ in aliquot_ids]
        top_labels = side_labels.copy()
        side_c4 = np.array([self.kit_type[:-4]] * aliquot_ids.shape[0])
        top_data = pd.DataFrame({'c1': top_idxes[:aliquot_ids.shape[0]], 'c2': top_labels, 'c3': aliquot_ids})
        side_data = pd.DataFrame({'c1': side_idxes[:aliquot_ids.shape[0]], 'c2': aliquot_ids, 'c3': side_labels, 'c4': side_c4})
        if self.kit_type == 'SERONETPBMC':
            side_data['c5'] = np.array([['10^7 for NCI', '10^7 for NCI', ''] for _ in sids]).flatten()
            top_data['c4'] = np.array([['NCI', 'NCI', ''] for _ in sids]).flatten()
        else:
            top_data['c4'] = np.nan
            side_data['c5'] = np.nan
        self.future_workbook[top_sheet] = top_data
        self.future_workbook[side_sheet] = side_data

    def twoml_round(self, round_num, sids):
        top_sheet = '{} - Tops {}'.format(round_num * 2 + 1, round_num + 1)
        side_sheet = '{} - Sides {}'.format(round_num * 2 + 2, round_num + 1)
        side_labels = np.array(['Plasma'] * self.kit_info['Plasma'] +
                       ['Serum'] * self.kit_info['Serum'] +
                       ['Saliva'] * self.kit_info['Saliva'] +
                       ['NPS'] * self.kit_info['NPS'])
        top_labels = side_labels.copy()
        top_labels[top_labels == 'Saliva'] = 'Slva'
        top_labels[top_labels == 'Plasma'] = 'Plsma'
        aliquot_ids = np.repeat(sids, side_labels.shape[0])
        top_count = self.kit_info['Tubes in Top Rack']
        top_racks = self.kit_info['Top Racks per Round']
        top_idxes = np.repeat(range(top_racks), top_count) * 78 + 1
        top_idxes += np.array([x for x in range(top_count)] * top_racks)
        side_count = self.kit_info['Tubes in Side Rack']
        side_racks = self.kit_info['Side Racks per Round']
        side_idxes = np.repeat(range(side_racks), side_count) * 50 + 1
        side_idxes += np.array([x for x in range(side_count)] * side_racks)
        top_data = pd.DataFrame({'c1': top_idxes[:aliquot_ids.shape[0]]})
        side_data = pd.DataFrame({'c1': side_idxes[:aliquot_ids.shape[0]], 'c2': np.nan})
        if self.kit_type == 'NPS':
            top_data['c2'] = [str(idx)[:4] for idx in aliquot_ids]
            top_data['c3'] = [str(idx)[4:] for idx in aliquot_ids]
            side_data['c3'] = aliquot_ids
        else:
            top_data['c2'] = np.nan
            top_data['c3'] = aliquot_ids
            side_data['c3'] = np.tile(side_labels, len(sids))
            side_data['c3'] += ' '
            side_data['c3'] += aliquot_ids
        top_data['c4'] = np.tile(top_labels, len(sids))
        side_data['c4'] = self.kit_type
        self.future_workbook[top_sheet] = top_data
        self.future_workbook[side_sheet] = side_data

    def write_fivemls(self):
        top_data = pd.DataFrame()
        side_data = pd.DataFrame()
        aliquot_ids = self.sample_ids
        aliquot_idxes = np.arange(len(aliquot_ids)) + 1
        top_data['c1'] = aliquot_idxes
        top_data['c2'] = 'SNET' # only SeroNet has 4.5mL tubes
        top_data['c3'] = aliquot_ids
        side_data['c1'] = aliquot_idxes
        side_data['c2'] = aliquot_ids
        side_data['c3'] = 'SERONET Serum'
        self.future_workbook['4.5mL Tops'] = top_data
        self.future_workbook['4.5mL Sides'] = side_data

if __name__ == '__main__':
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    session_counts = parse_input()

    templates = pd.ExcelFile(util.tube_print + 'Future Sheets' + os.sep + 'Central Template Sheet.xlsx')
    for print_type, sessions in session_counts.items():
        if print_type == 'SERONET_RTC':
            continue
        recent_box_max = last_printed_box(print_type)
        for session_num in range(sessions):
            kit_type = print_type
            if print_type == 'SERONET' and session_num < session_counts['SERONET_RTC']:
                kit_type = 'SERONET_RTC'
            print_session = PrintSession(kit_type, templates, session_num)
            if not print_session.valid_sid_set:
                print(f"""Need more assigned sample IDs for {print_type}.
                      Could not generate {print_session.box_start} - {print_session.box_end}.""")
                break

            print_session.write_workbook()
            print(f"'{print_session.output_path.split(os.sep)[-1]}' workbook generated in {print_session.kit_info['Future Sheets Folder']}.")
