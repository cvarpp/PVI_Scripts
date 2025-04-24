import numpy as np
import pandas as pd
import os
import requests
import argparse

"""
CLI script that locates samples in FreezerPro
Input: An excel file with a column of sample IDs with header "Sample ID"
Output: An excel file with data about the number and type of samples for each ID
Args: -i for input, -o for output, -d for debug (optional)
Example: fp_locs.py -i inputs.xlsx -o outputs.xlsx
Works best with environment variables named FP_USER and FP_PASS containing nuclear launch codes

To Do: Add detailed information about the location of each sample to the output file
NOTES:
-Refused to work until I installed pip-system-certs through pip despite already having it
"""

# Establish a table for convering FP numerical sample types to human readable form
future_vial_info = {'ID': ['25', '26', '27', '28', '29', '32'], 'Name': ['Serum','Plasma','Pellet','Saliva','PBMC','4.5 mL Tube']}
vial_info = pd.DataFrame(future_vial_info).set_index('ID')
fp_url = "https://mssm-simonlab.FreezerPro.com/api/v2"

# If script is being run directly...
if __name__ == '__main__':

    # Parse CLI arguments and store in 'args'
    argparser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    argparser.add_argument('-d', '--debug', action='store_true')
    argparser.add_argument('-i', '--input_fname', action='store', required=True)
    argparser.add_argument('-o', '--output_fname', action='store', required=True)
    args = argparser.parse_args()

    # Read input excel file into a pandas dataframe and store cleaned inputs in 'soi'
    if args.input_fname[-3:] == 'csv':
        soi = pd.read_csv(args.input_fname)['Sample ID'].astype(str).str.strip().str.upper()
    elif args.input_fname[-4:].isin(['xlsx', 'xlsm', '.xls']):
        soi = pd.read_excel(args.input_fname)['Sample ID'].astype(str).str.strip().str.upper()

    # Authenticate using username and password, then store token
    username = os.environ.get("FP_USER")
    password = os.environ.get("FP_PASS")
    if not username or not password:
        print("FP_USER and FP_PASS env variables not found. Please update then restart your shell.")
        manual_response = ""
        username = ""
        password = ""
        while manual_response == "":
            manual_response = input("Input username/password manually? Y/N: ").lower()
            if(manual_response in ["y", "yes", "sure, why not", "ok", "fine", "yes please"]):
                username = input("Input username: ")
                password = input ("Input password: ")
            elif(manual_response == 'n'):
                continue
            else:
                manual_response = ""
    token_response = requests.post(f'{fp_url}/auth/login', json={'username': username, 'password': password})
    if token_response.status_code != 200:
        print("Failed authentication. Fatal Error. Exiting...")
        exit(1)
    token = token_response.json()['data']['attributes']['token']

    # Save token to json header for requesting sample data
    headers = {'Authorization': f'token {token}'}
    # List appropriate vial types for requesting sample data
    vial_type_suffix = 'include=vials&fields[sample]=name,vials,sample_type&fields[vial]=position_display,barcode_tag,box'
    # Make empty output lists
    cols = ['Sample ID', 'Sample Type', 'Vial ID']
    vial_cols = ['Vial ID', 'Barcode', 'Position', 'Box ID']
    future_sids = {col: [] for col in cols}
    future_vials = {col: [] for col in vial_cols}
    # Estimate querying time
    print("Querying FP for {} samples... Estimated time {:.0f}M {:.0f}S".format(soi.shape[0],
                                                                                  soi.shape[0] * 0.15 // 60,
                                                                                  soi.shape[0] * .15 % 60))

    # Fetch data for each sample ID
    for sid in soi.to_numpy():
        sid_response = requests.get(f'{fp_url}/samples?filter[name_eq]={sid}&{vial_type_suffix}', headers=headers)
        if sid_response.status_code != 200:
            print("Failed to query sample ID {}.".format(sid))
        else:
            # Parse and store the response for each successfully queried sample ID
            sid_json = sid_response.json()
            for sample_data in sid_json['data']:
                assert sample_data['attributes']['name'] == sid, "Sample ID returned differed from the query."
                stype = sample_data['relationships']['sample_type']['data']['id']
                if stype not in vial_info.index:
                    print(stype, "not in known sample types.")
                    continue
                for vial in sample_data['relationships']['vials']['data']:
                    future_sids['Sample ID'].append(sid) # Log the ID
                    future_sids['Sample Type'].append(vial_info.loc[stype, 'Name']) # Lookup and log the sample type
                    future_sids['Vial ID'].append(vial['id'])
            try:
                for vial_data in sid_json['included']:
                    if vial_data['type'] == 'sample_type': # Type is almost always vial
                        continue
                    future_vials['Vial ID'].append(vial_data['id'])
                    future_vials['Barcode'].append(vial_data['attributes']['barcode_tag'])
                    future_vials['Position'].append(vial_data['attributes']['position_display'])
                    future_vials['Box ID'].append(vial_data['relationships']['box']['data']['id'])
            except KeyError:
                if args.debug:
                    print(f"Sample {sid} failed to return box information.")

    
    # Convert all of the lists of dicts we made into pandas dataframes
    samples = pd.DataFrame(future_sids)
    vials = pd.DataFrame(future_vials).set_index('Vial ID')
    box_cols = ['ID', 'Box Name', 'Box Barcode', 'Container']
    future_boxes = {col: [] for col in box_cols}

    # A similar process is done for each unique box ID found above
    print("Querying FP for {} boxes... Estimated time {:.0f}M {:.0f}S".format(vials['Box ID'].unique().shape[0], 
                                                                         vials['Box ID'].unique().shape[0] * 0.05 // 60,
                                                                         vials['Box ID'].unique().shape[0] * 0.05 % 60)) # This is faster than fetching vials
    for box_id in vials['Box ID'].unique():
        box_response = requests.get(f'{fp_url}/boxes/{box_id}?include=container&fields[box]=name,barcode_tag&fields[container]=name,barcode_tag', headers=headers)
        if box_response.status_code != 200:
            print("Failed to query box {}.".format(box_id))
            print(box_response)
        else:
            box_json = box_response.json()
            future_boxes['ID'].append(box_id)
            future_boxes['Box Name'].append(box_json['data']['attributes']['name'])
            future_boxes['Box Barcode'].append(box_json['data']['attributes']['barcode_tag'])
            future_boxes['Container'].append(box_json['included'][0]['attributes']['name'])

    # And again convert our lists of dicts to dataframes
    boxes = pd.DataFrame(future_boxes).set_index('ID')
    output = samples.join(vials, on='Vial ID').join(boxes, on='Box ID').sort_values(by='Box Name', key=lambda s: s.astype(str).str.contains('FF'), kind='stable')
    big_boxes = output[~output['Position'].astype(str).str.contains('/')]
    reg_boxes = output[output['Position'].astype(str).str.contains('/')]
    typecounts = reg_boxes.loc[:, ['Sample ID', 'Sample Type', 'Vial ID']].groupby(['Sample ID', 'Sample Type']
                    ).count().unstack(level=-1).droplevel(0, axis=1)

    # Write data to output excel file
    if not args.debug:
        with pd.ExcelWriter(args.output_fname) as writer:
            typecounts.fillna(0).to_excel(writer, sheet_name='Summary')
            for stype in vial_info['Name'].unique():
                filtered = reg_boxes[reg_boxes['Sample Type'] == stype]
                if filtered.shape[0] > 0:
                    filtered.to_excel(writer, sheet_name=stype, index=False)
            if big_boxes.shape[0] > 0:
                big_boxes.to_excel(writer, sheet_name='Collaborator or Discarded', index=False)


