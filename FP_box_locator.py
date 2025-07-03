import numpy as np
import pandas as pd
import requests
import time
import os
import argparse
import getpass
import util

import json # for debugging

"""
CLI Script to return box locations from FreezerPro
Requires an input csv with columns named UID, Box, and Barcode containing box IDs, box names. and box barcode numbers.
"""

# If script is being run directly...
if __name__ == '__main__':
    # Handle arguments
    argparser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    argparser.add_argument('-i', '--input_fname', action='store', required=True)
    argparser.add_argument('-o', '--output_fname', action='store', default="9x9_box_locations.csv")
    argparser.add_argument('-t', '--sleep_time', action='store', default=.04)
    args = argparser.parse_args()

    assert 'xls' in args.output_fname[-4:], "Output filename should end in xls[xm]?"

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
        
    token_response = requests.post(f'{util.fp_url}auth/login', json={'username': username, 'password': password})
    if token_response.status_code != 200:
        print("Failed authentication. Fatal Error. Exiting...")
        exit(1)
    token = token_response.json()['data']['attributes']['token']
    headers = {'Authorization': f'token {token}'}

    # For each box ID
    how_many = len(box_ids)
    duration = int(round(how_many / 6, 0)) # Determined experimentally
    time_units = "seconds" if duration < 60 else "minutes"
    duration = duration if duration < 60 else int(duration / 60)
    print(f"Searching for {how_many} boxes. This will take at least {duration} {time_units}.")
    last_call = [time.time()]
    def make_request(s, headers=headers, last_call=last_call):
        diff = time.time() - last_call[0]
        if diff < 0.05:
            time.sleep(0.05 - diff + 0.005)
        last_call[0] = time.time()
        response = requests.get(s, headers=headers)
        return response
    pre_box_queries = time.time()
    box_parents = {'Box ID': [], 'Parent ID': []}
    freezers = {}
    levels = []
    for _ in range(5):
        levels.append({})
    for i, box_id in enumerate(box_ids):
        box_response = make_request(f'{util.fp_url}boxes/{box_id}?include=parents')
        if box_response.status_code != 200:
            print("Failed to query box {}.".format(box_id))
            time.sleep(0.05)
            freezers[box_id] = np.nan
            for i in range(5):
                levels[i][box_id] = np.nan
        else:
            box_json = box_response.json()
            parents = [parent["attributes"]["name"] for parent in box_json["included"]]
            if len(parents) == 0:
                print("Panicking!")
                print(box_json)
                exit(1)
            freezers[box_id] = parents[0]
            for i in range(5):
                if i < len(parents) - 1:
                    levels[i][box_id] = parents[i + 1]
                else:
                    levels[i][box_id] = np.nan
    box_inputs['Freezer'] = box_inputs["UID"].apply(lambda uid: freezers[uid])
    for i in range(5):
        box_inputs[f'Level{i+1}'] = box_inputs["UID"].apply(lambda uid: levels[i][uid])
    print('{:.1f}'.format(time.time() - pre_box_queries), "seconds for", how_many, "boxes")
    box_inputs.to_excel(args.output_fname, index=False, freeze_panes=(1,2))
