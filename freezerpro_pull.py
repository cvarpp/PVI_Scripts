import requests
import os
import pandas as pd
import urllib3

if __name__ == '__main__':
    # urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    fp_url = "https://mssm-simonlab.freezerpro.com/api"
    freezers = (requests.post(fp_url, params={'method': 'freezers'},
                              auth=('api', os.environ["FP_PASS"]))
                        .json()["Freezers"])
    freezer_info = {k: [] for k in freezers[0].keys()}
    demo_subdiv = (requests.post(fp_url, params={'method': 'subdivisions', 'id': freezers[0]['id']},
                                auth=('api', os.environ["FP_PASS"]))
                           .json()['Subdivisions'][0])
    subdiv_info = {k: [] for k in demo_subdiv.keys()}
    subdiv_info['freezer_id'] = []
    subdiv_info['freezer_name'] = []
    subdiv_info['freezer_rfid'] = []
    for freezer in freezers:
        print(freezer['name'])
        for k in freezer.keys():
            freezer_info[k].append(freezer[k])
        subdiv1s = (requests.post(fp_url, params={'method': 'subdivisions', 'id': freezer['id']},
                                auth=('api', os.environ["FP_PASS"]))
                           .json()['Subdivisions'])
        for subdiv1 in subdiv1s:
            for k in subdiv1.keys():
                subdiv_info[k].append(subdiv1[k])
            subdiv_info['freezer_id'].append(freezer['id'])
            subdiv_info['freezer_name'].append(freezer['name'])
            subdiv_info['freezer_rfid'].append(freezer['rfid_tag'])
            subdiv2s = (requests.post(fp_url, params={'method': 'subdivisions', 'id': subdiv1['id']},
                                    auth=('api', os.environ["FP_PASS"]))
                                .json()['Subdivisions'])
            for subdiv2 in subdiv2s:
                for k in subdiv2.keys():
                    subdiv_info[k].append(subdiv2[k])
                subdiv_info['freezer_id'].append(subdiv1['id'])
                subdiv_info['freezer_name'].append(subdiv1['name'])
                subdiv_info['freezer_rfid'].append(subdiv1['rfid_tag'])
                subdiv3s = (requests.post(fp_url, params={'method': 'subdivisions', 'id': subdiv2['id']},
                                        auth=('api', os.environ["FP_PASS"]))
                                    .json()['Subdivisions'])
                for subdiv3 in subdiv3s:
                    for k in subdiv3.keys():
                        subdiv_info[k].append(subdiv3[k])
                    subdiv_info['freezer_id'].append(subdiv2['id'])
                    subdiv_info['freezer_name'].append(subdiv2['name'])
                    subdiv_info['freezer_rfid'].append(subdiv2['rfid_tag'])
    boxes = (requests.post(fp_url, params={'method': 'boxes', 'limit': '5000'},
                              auth=('api', os.environ["FP_PASS"]))
                        .json()["Boxes"])
    box_info = {k: [] for k in boxes[0].keys()}
    for box in boxes:
        for k in box.keys():
            box_info[k].append(box[k])
    writer = pd.ExcelWriter('data/freezerpro_pull.xlsx') # Name of 
    freezer_df = pd.DataFrame(freezer_info)
    freezer_df.to_excel(writer, sheet_name='Freezers', index=False)
    subdiv_df = pd.DataFrame(subdiv_info)
    subdiv_df.to_excel(writer, sheet_name='Subdivisions', index=False)
    box_df = pd.DataFrame(box_info)
    box_df.to_excel(writer, sheet_name='Boxes', index=False)
    writer.save()
    writer.close()