import pandas as pd
import numpy as np
import argparse
import util
from helpers import query_intake, query_dscf, clean_sample_id
from cam_convert import transform_cam
import datetime
import os
import requests
from time import sleep

def query_import_sheet(df_valid):
    inventory_boxes = pd.read_excel(util.inventory_input, sheet_name=None)
    boxes = []
    for sname, sheet in inventory_boxes.items():
        if 'Sample ID' not in sheet.columns or 'PSP' in sname or 'cell' in sname.lower() or 'NPS' in sname:
            continue
        if sname in ['TEMPLATE', 'Box Converter']:
            continue
        if 'Sample Type' not in sheet.columns:
            print(sname, "missing sample type column")
            continue
        boxes.append(sheet.reset_index().rename(
                columns={'index': 'Position'}
            ).dropna(subset=['Sample ID']).assign(
            sample_id=clean_sample_id,
            Box=sname,
            sample_type=lambda df: df['Sample Type'].str.strip().str.title().str.rstrip('s')
            ).loc[:, ['sample_id', 'Box', 'Position', 'sample_type']]
        )
    inventory = pd.concat(boxes, axis=0).join(df_valid.set_index('sample_id'), on='sample_id', rsuffix='_not_inventory')
    return inventory

def write_missing_ids(df_valid, invalid_ids, recency_date):
    in_intake = df_valid.dropna(subset='participant_id').index
    missing_intake = df_valid.loc[[idx for idx in df_valid.index if idx not in in_intake], :]
    in_dscf = df_valid.dropna(subset='Date Processing Started').index
    missing_dscf = df_valid.loc[[idx for idx in df_valid.index if idx not in in_dscf], :]
    in_cam = df_valid.dropna(subset='Date').index
    missing_cam = df_valid.loc[[idx for idx in df_valid.index if idx not in in_cam], :]
    fname_missing_ids = util.proc + 'Troubleshooting/Cross-Sheet Issues.xlsx'
    with pd.ExcelWriter(fname_missing_ids) as writer:
        invalid_ids = invalid_ids.loc[:, ['sample_id', 'Date Collected', 'Date Processing Started', 'Date', 'proc_inits']]
        recent_invalid = invalid_ids[(invalid_ids['Date Collected'] > recency_date) | (invalid_ids['Date Processing Started'] > recency_date)]
        recent_invalid.to_excel(writer, sheet_name='Invalid IDs (Recent)', index=False, freeze_panes=(1,1))
        missing_intake = missing_intake.loc[:, ['sample_id', 'Date Collected', 'Date Processing Started', 'Date', 'proc_inits']]
        recent_missing_intake = missing_intake[missing_intake['Date Processing Started'] > recency_date]
        recent_missing_intake.to_excel(writer, sheet_name='Missing from Intake (Recent)', index=False, freeze_panes=(1,1))
        missing_dscf = missing_dscf.loc[:, ['sample_id', 'Date Collected', 'Study', 'Visit Type / Samples Needed', 'participant_id']]
        recent_missing_dscf = missing_dscf[(missing_dscf['Date Collected'] > recency_date)]
        recent_missing_dscf.to_excel(writer, sheet_name='Missing from DSCF (Recent)', index=False, freeze_panes=(1,1))
        missing_cam = missing_cam.loc[:, ['sample_id', 'Date Collected', 'Study', 'Visit Type / Samples Needed', 'participant_id', 'proc_inits']]
        recent_missing_cam = missing_cam[(missing_cam['Date Collected'] > recency_date)]
        recent_missing_cam.to_excel(writer, sheet_name='Missing from CAM (Recent)', index=False, freeze_panes=(1,1))
        invalid_ids.to_excel(writer, sheet_name='Invalid IDs', index=False)
        missing_intake.to_excel(writer, sheet_name='Missing from Intake', index=False)
        missing_dscf.to_excel(writer, sheet_name='Missing from DSCF', index=False)
        missing_cam.to_excel(writer, sheet_name='Missing from CAM', index=False)
    print("Report written to {}".format(fname_missing_ids))
    print("DSCF: {}, Intake: {}, CAM: {}, Invalid: {}".format(recent_missing_dscf.shape[0], recent_missing_intake.shape[0], recent_missing_cam.shape[0], recent_invalid.shape[0]))

def query_fp(recent_valid, inventory_counts):
    sample_types = ['Plasma', 'Serum', 'Saliva', 'Pellet', 'PBMC', '4.5 mL Tube']
    fp_user = os.environ[util.fp_user]
    fp_pass = os.environ[util.fp_pass]
    token_response = requests.post(f'{util.fp_url}/auth/login', json={'username': fp_user, 'password': fp_pass})
    if token_response.status_code != 200:
        print("Failed authentication. Fatal error, exiting...")
        exit(1)
    token = token_response.json()['data']['attributes']['token']
    headers = {'Authorization': f'token {token}'}
    vial_type_suffix = 'include=sample_type&fields[sample]=name,vials&fields[sample_type]=name'
    fp_data = {'Sample ID': []}
    for stype in sample_types:
        fp_data[stype] = []
    print("Querying FP for {} samples... Estimated time {:.1f} minutes".format(recent_valid.shape[0], recent_valid.shape[0] * 2.2 / 60))
    for sid in recent_valid['sample_id'].to_numpy():
        sleep(2.2)
        sid_response = requests.get(f'{util.fp_url}/samples?filter[name_eq]={sid}&{vial_type_suffix}', headers=headers)
        if sid_response.status_code != 200:
            print("Failed to query sample ID {}.".format(sid))
            for col in fp_data.keys():
                if col == 'Sample ID':
                    fp_data[col].append(sid)
                else:
                    fp_data[col].append(0)
        elif len(sid_response.json()['data']) == 0:
            for col in fp_data.keys():
                if col == 'Sample ID':
                    fp_data[col].append(sid)
                else:
                    fp_data[col].append(0)
        else:
            sid_json = sid_response.json()
            sid_data = {}
            for sample, sample_type in zip(sid_json['data'], sid_json['included']):
                stype = sample_type['attributes']['name']
                if stype not in sid_data.keys():
                    sid_data[stype] = 0
                sid_data[stype] += len(sample['relationships']['vials']['data'])
            for col in fp_data.keys():
                if col == 'Sample ID':
                    fp_data[col].append(sid)
                elif col in sid_data.keys():
                    fp_data[col].append(sid_data[col])
                else:
                    fp_data[col].append(0)
    fp_df = pd.DataFrame(fp_data)
    proc_info = recent_valid.set_index('sample_id').loc[:, ['Date Collected', '# of aliquots frozen', 'viability', '# cells per aliquot', 'cpt_vol', 'sst_vol',
                                        'Total volume of plasma (mL)', 'Total volume of serum (mL)', 'Saliva Volume (mL)',
                                        '4.5 mL Tube Needed', '4.5 mL Aliquot?', 'proc_inits']]
    all_inv = fp_df.join(inventory_counts, on='Sample ID', lsuffix='_fp', rsuffix='_import').fillna(0)
    for stype in sample_types:
        all_inv[stype] = all_inv[stype + '_fp'] + all_inv[stype + '_import']
    all_inv['Import Sheet Aliquots'] = all_inv.loc[:, [stype + '_import' for stype in sample_types]].sum(axis=1)
    all_inv['Still in Import Sheet'] = (all_inv['Import Sheet Aliquots'] > 0).apply(lambda val: "Yes" if val else "No")
    output = proc_info.join(all_inv.set_index('Sample ID')).loc[:, ['Date Collected', 'Still in Import Sheet', 'cpt_vol', 'Total volume of plasma (mL)',
        'Plasma', '# of aliquots frozen', '# cells per aliquot', 'PBMC', 'sst_vol', 'Total volume of serum (mL)', 'Serum', '4.5 mL Tube Needed',
        '4.5 mL Aliquot?', '4.5 mL Tube', 'Saliva Volume (mL)', 'Saliva', 'Import Sheet Aliquots', 'viability', 'proc_inits',
        'Pellet'] + [stype + '_fp' for stype in sample_types] + [stype + '_import' for stype in sample_types]]
    return output

if __name__ == '__main__':
    argparser = argparse.ArgumentParser()
    argparser.add_argument('-d', '--debug', action='store_true')
    argparser.add_argument('-c', '--use_cache', action='store_true')
    argparser.add_argument('-r', '--recent_cutoff', type=int, default=120, help='Number of days before today to consider recent')
    argparser.add_argument('-fp', '--freezerpro', action='store_true')
    args = argparser.parse_args()

    intake = query_intake(include_research=True, use_cache=args.use_cache)
    dscf = query_dscf(use_cache=args.use_cache).rename(columns={'Sample ID': 'Processing Sample ID'})
    if args.use_cache:
        cam_full = pd.read_excel(util.cam_long)
    else:
        cam_full = transform_cam(debug=args.debug)
    cam_full = cam_full.dropna(subset=['Participant ID', 'sample_id']).drop_duplicates(subset='sample_id').set_index('sample_id').loc[:, ['Date', 'Participant ID', 'Time', 'Time Collected', 'Phlebotomist', 'Visit Coordinator']]
    df = intake.join(dscf, how='outer').join(cam_full, how='outer', rsuffix='_cam').reset_index().sort_values(by=['Date Collected', 'Date Processing Started', 'sample_id'])
    valid_ids = set(pd.read_excel(util.tracking + 'Sample ID Master List.xlsx', sheet_name='Master Sheet').dropna(subset='Location')['Sample ID'].astype(str).unique())
    invalid_ids = df[~df['sample_id'].isin(valid_ids)]
    df_valid = df[df['sample_id'].isin(valid_ids)]
    recency_date = (datetime.datetime.today() - datetime.timedelta(days=args.recent_cutoff)).date()
    if not args.debug:
        write_missing_ids(df_valid, invalid_ids, recency_date)
    inventory = query_import_sheet(df_valid)
    inventory_counts = inventory.loc[:, ['sample_id', 'sample_type', 'Position']].groupby(
        ['sample_id', 'sample_type']
        ).count().unstack().droplevel(0, axis='columns').fillna(0).rename(
            columns={'4.5 Ml Tube': '4.5 mL Tube', 'Pbmc': 'PBMC'}
        ).astype(int)
    inventory_missing_intake = inventory[inventory['Date Collected'].isna()]
    recent_valid = df_valid[df_valid['Date Collected'] > recency_date].copy()
    fname_inventory = util.proc + 'Troubleshooting/Inventory Check.xlsx'
    if not args.debug:
        with pd.ExcelWriter(fname_inventory) as writer:
            inventory_missing_intake.loc[:, ['sample_id', 'Box', 'Position', 'sample_type']].to_excel(writer, sheet_name='Missing from Intake', index=False, freeze_panes=(1, 1))
            inventory_counts.to_excel(writer, sheet_name='New Import Sheet Inventory', freeze_panes=(1, 1))
            if args.freezerpro:
                inv_check = query_fp(recent_valid, inventory_counts)
                inv_check.to_excel(writer, sheet_name='Inventory vs Expected', freeze_panes=(1, 1))
        print("Inventory typo report written to {}".format(fname_inventory))
    print("{} likely inventory typos".format(inventory_missing_intake.shape[0]))

