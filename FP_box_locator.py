import numpy as np
import pandas as pd
import requests
import csv
import time
import os
import argparse
import getpass



"""
CLI Script to return box locations from FreezerPro
Requires an input csv with columns named UID, Box, and Barcode containing box IDs, box names. and box barcode numbers.
"""

fp_url = "https://mssm-simonlab.FreezerPro.com/api/v2"

# If script is being run directly...
if __name__ == '__main__':
    # Handle arguments
    argparser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    argparser.add_argument('-d', '--debug', action='store_true')
    argparser.add_argument('-i', '--input_fname', action='store', required=True)
    argparser.add_argument('-o', '--output_fname', action='store', default="9x9_box_locations.csv")
    argparser.add_argument('-t', '--sleep_time', action='store', default=.04)
    args = argparser.parse_args()

    if args.output_fname[-4:] != ".csv":
        print("Error: Output file must be .csv\nExiting...")
        exit(1)

    # Parse input file
    box_inputs = pd.read_csv(args.input_fname, header=0)
    box_ids = box_inputs['UID'].to_list()
    box_names = box_inputs['Box'].to_list()
    barcodes = box_inputs['Barcode'].to_list()


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
                password = getpass.getpass("Input password: ")
            elif(manual_response == 'n'):
                continue
            else:
                manual_response = ""
        
    token_response = requests.post(f'{fp_url}/auth/login', json={'username': username, 'password': password})
    if token_response.status_code != 200:
        print("Failed authentication. Fatal Error. Exiting...")
        exit(1)
    token = token_response.json()['data']['attributes']['token']
    headers = {'Authorization': f'token {token}'}

    # Create output list
    all_parent_names = [["Barcode", "Box Name", "Location"]]

    # Iterate through input IDs and find all parents by repeatedly looking up the immediate parent. Store names.
    iterator = 0
    # For each box ID
    how_many = len(box_ids)
    duration = int(args.sleep_time * how_many * 9) # Determined experimentally
    time_units = "seconds" if duration < 60 else "minutes"
    duration = duration if duration < 60 else int(duration / 60)
    print(f"Searching for {how_many} boxes. This will take at least {duration} {time_units}.")
    for box_id in box_ids:
        time.sleep(args.sleep_time)
        box_response = requests.get(f'{fp_url}/boxes/{box_id}?include=container&fields[box]=name,barcode_tag&fields[container]=name,barcode_tag', headers=headers)
        parent_names = []
        parent_ids = []
        parent_names.append(barcodes[iterator])
        parent_names.append(box_names[iterator])
        iterator += 1
        if box_response.status_code != 200:
            print("Failed to query box {}.".format(box_id))
        else:
            box_json = box_response.json()
            current_id = int(box_json['included'][0]['id'])
            current_name = box_json['included'][0]['attributes']['name']
            parent_ids.append(current_id)
            success = True
            while success:
                try:
                    container_json = requests.get(f'{fp_url}/subdivisions/{current_id}', headers=headers).json()
                    current_id = int(container_json['data']['relationships']['parent']['data']['id'])
                    current_name = container_json['data']['attributes']['name']
                    parent_names.append(current_name)
                    parent_ids.append(current_id)
                except KeyError:
                    try:
                        container_json = requests.get(f'{fp_url}/freezers/{current_id}', headers=headers).json()
                        current_name = container_json['data']['attributes']['name']
                        parent_names.append(current_name)
                        parent_ids.append(current_id)
                    finally:
                        success = False
        all_parent_names.append(parent_names)

        # Save to file every so often in case the program crashes, you get Fail2Banned, etc

        if iterator % 200 == 0:
            with open(args.output_fname, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerows(all_parent_names)

# Write output to csv

with open(args.output_fname, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerows(all_parent_names)