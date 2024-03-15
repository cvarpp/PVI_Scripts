import numpy as np
import pandas as pd
import util
import os
import requests
from time import sleep
import argparse

future_vial_info = {'ID': ['25', '26', '27', '28', '29', '32'], 'Name': ['Serum','Plasma','Pellet','Saliva','PBMC','4.5 mL Tube']}
vial_info = pd.DataFrame(future_vial_info).set_index('ID')

if __name__ == '__main__':
    argparser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    argparser.add_argument('-d', '--debug', action='store_true')
    argparser.add_argument('-i', '--input_fname', action='store', required=True)
    argparser.add_argument('-o', '--output_fname', action='store', required=True)
    args = argparser.parse_args()
    soi = pd.read_excel(args.input_fname)
    if 'Sample ID' in soi.columns:
        soi['sample_id'] = soi['Sample ID'].astype(str).str.strip().str.upper()
    sample_types = ['Plasma', 'Serum', 'Saliva', 'Pellet', 'PBMC', '4.5 mL Tube']
    fp_user = os.environ[util.fp_user]
    fp_pass = os.environ[util.fp_pass]
    token_response = requests.post(f'{util.fp_url}/auth/login', json={'username': fp_user, 'password': fp_pass})
    if token_response.status_code != 200:
        print("Failed authentication. Fatal error, exiting...")
        exit(1)
    token = token_response.json()['data']['attributes']['token']
    headers = {'Authorization': f'token {token}'}
    vial_type_suffix = 'include=vials&fields[sample]=name,vials,sample_type&fields[vial]=position_display,barcode_tag,box'
    cols = ['Sample ID', 'Sample Type', 'Vial ID']
    vial_cols = ['Vial ID', 'Barcode', 'Position', 'Box ID']
    future_sids = {col: [] for col in cols}
    future_vials = {col: [] for col in vial_cols}
    print("Querying FP for {} samples... Estimated time {:.1f} minutes".format(soi.shape[0], soi.shape[0] * 0.3 / 60))
    for sid in soi['sample_id'].to_numpy()[:10]:
        sleep(0.3)
        sid_response = requests.get(f'{util.fp_url}/samples?filter[name_eq]={sid}&{vial_type_suffix}', headers=headers)
        if sid_response.status_code != 200:
            print("Failed to query sample ID {}.".format(sid))
        else:
            sid_json = sid_response.json()
            for sample_data in sid_json['data']:
                assert sample_data['attributes']['name'] == sid
                stype = sample_data['relationships']['sample_type']['data']['id']
                if stype not in vial_info.index:
                    print(stype, "not in known sample types.")
                    continue
                for vial in sample_data['relationships']['vials']['data']:
                    future_sids['Sample ID'].append(sid)
                    future_sids['Sample Type'].append(vial_info.loc[stype, 'Name'])
                    future_sids['Vial ID'].append(vial['id'])
            for vial_data in sid_json['included']:
                if vial_data['type'] == 'sample_type':
                    continue
                future_vials['Vial ID'].append(vial_data['id'])
                future_vials['Barcode'].append(vial_data['attributes']['barcode_tag'])
                future_vials['Position'].append(vial_data['attributes']['position_display'])
                future_vials['Box ID'].append(vial_data['relationships']['box']['data']['id'])
    samples = pd.DataFrame(future_sids)
    vials = pd.DataFrame(future_vials).set_index('Vial ID')
    box_cols = ['ID', 'Box Name', 'Box Barcode', 'Container']
    future_boxes = {col: [] for col in box_cols}
    print("Querying FP for {} boxes... Estimated time {:.1f} minutes".format(vials['Box ID'].unique().shape[0], vials['Box ID'].unique().shape[0] * 0.3 / 60))
    for box_id in vials['Box ID'].unique():
        sleep(0.3)
        box_response = requests.get(f'{util.fp_url}/boxes/{box_id}?include=container&fields[box]=name,barcode_tag&fields[container]=name,barcode_tag', headers=headers)
        if box_response.status_code != 200:
            print("Failed to query box {}.".format(box_id))
            print(box_response)
        else:
            box_json = box_response.json()
            future_boxes['ID'].append(box_id)
            future_boxes['Box Name'].append(box_json['data']['attributes']['name'])
            future_boxes['Box Barcode'].append(box_json['data']['attributes']['barcode_tag'])
            future_boxes['Container'].append(box_json['included'][0]['attributes']['name'])
    boxes = pd.DataFrame(future_boxes).set_index('ID')
    output = samples.join(vials, on='Vial ID').join(boxes, on='Box ID').sort_values(by='Box Name', key=lambda s: s.astype(str).str.contains('FF'), kind='stable')
    big_boxes = output[~output['Position'].astype(str).str.contains('/')]
    reg_boxes = output[output['Position'].astype(str).str.contains('/')]
    typecounts = reg_boxes.loc[:, ['Sample ID', 'Sample Type', 'Vial ID']].groupby(['Sample ID', 'Sample Type']
                    ).count().unstack(level=-1).droplevel(0, axis=1)

    if not args.debug:
        with pd.ExcelWriter(args.output_fname) as writer:
            typecounts.fillna(0).to_excel(writer, sheet_name='Summary')
            for stype in sample_types:
                filtered = reg_boxes[reg_boxes['Sample Type'] == stype]
                filtered.to_excel(writer, sheet_name=stype, index=False)
            big_boxes.to_excel(writer, sheet_name='Collaborator or Discarded', index=False)
