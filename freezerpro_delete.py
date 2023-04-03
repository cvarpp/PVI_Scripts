import requests
import os
import pandas as pd
import urllib3

if __name__ == '__main__':
    exit(0)
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    fp_url = "https://mssm-simonlab.freezerpro.com/api"
    freezers = (requests.post(fp_url, params={'method': 'freezers'},
                              auth=('api', os.environ["FP_PASS"]), verify=False)
                        .json()["Freezers"])
    freezer_info = {k: [] for k in freezers[0].keys()}
    demo_subdiv = (requests.post(fp_url, params={'method': 'subdivisions', 'id': freezers[0]['id']},
                                auth=('api', os.environ["FP_PASS"]), verify=False)
                           .json()['Subdivisions'][0])