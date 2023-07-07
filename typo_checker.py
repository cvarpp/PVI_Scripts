import pandas as pd
import numpy as np
import argparse
import util
from helpers import query_intake, query_dscf
from cam_convert import transform_cam
import datetime

if __name__ == '__main__':
    argparser = argparse.ArgumentParser()
    argparser.add_argument('-d', '--debug', action='store_true')
    argparser.add_argument('-c', '--use_cache', action='store_true')
    argparser.add_argument('-r', '--recent_cutoff', type=int, default=120, help='Number of days before today to consider recent')
    args = argparser.parse_args()

    intake = query_intake(include_research=True, use_cache=args.use_cache)
    dscf = query_dscf(use_cache=args.use_cache)
    cam_full = transform_cam(debug=True).dropna(subset=['Participant ID', 'sample_id']).drop_duplicates(subset='sample_id').set_index('sample_id')
    df = intake.join(dscf, how='outer').join(cam_full, how='outer', rsuffix='_cam').reset_index().sort_values(by=['Date Collected', 'Date Processing Started', 'sample_id'])
    valid_ids = set(pd.read_excel(util.tracking + 'Sample ID Master List.xlsx', sheet_name='Master Sheet').dropna(subset='Location')['Sample ID'].astype(str).unique())
    invalid_ids = df[~df['sample_id'].isin(valid_ids)]
    df_valid = df[df['sample_id'].isin(valid_ids)]
    in_intake = df_valid.dropna(subset='participant_id').index
    missing_intake = df_valid.loc[[idx for idx in df_valid.index if idx not in in_intake], :]
    in_dscf = df_valid.dropna(subset='Date Processing Started').index
    missing_dscf = df_valid.loc[[idx for idx in df_valid.index if idx not in in_dscf], :]
    in_cam = df_valid.dropna(subset='Date').index
    missing_cam = df_valid.loc[[idx for idx in df_valid.index if idx not in in_cam], :]
    recency_date = (datetime.datetime.today() - datetime.timedelta(days=args.recent_cutoff)).date()
    if not args.debug:
        with pd.ExcelWriter(util.clin_ops + 'Daily Troubleshooting.xlsx') as writer:
            recent_invalid = invalid_ids[(invalid_ids['Date Collected'] > recency_date) | (invalid_ids['Date Processing Started'] > recency_date)]
            recent_invalid.to_excel(writer, sheet_name='Invalid IDs (Recent)', index=False)
            recent_missing_intake = missing_intake[missing_intake['Date Collected'] > recency_date]
            recent_missing_intake.to_excel(writer, sheet_name='Missing from Intake (Recent)', index=False)
            recent_missing_dscf = missing_dscf[(missing_dscf['Date Collected'] > recency_date)]
            recent_missing_dscf.to_excel(writer, sheet_name='Missing from DSCF (Recent)', index=False)
            recent_missing_cam = missing_cam[(missing_cam['Date Collected'] > recency_date)]
            recent_missing_cam.to_excel(writer, sheet_name='Missing from CAM (Recent)', index=False)
            invalid_ids.to_excel(writer, sheet_name='Invalid IDs', index=False)
            missing_intake.to_excel(writer, sheet_name='Missing from Intake', index=False)
            missing_dscf.to_excel(writer, sheet_name='Missing from DSCF', index=False)
            missing_cam.to_excel(writer, sheet_name='Missing from CAM', index=False)
    print("DSCF: {}, Intake: {}, CAM: {}, Invalid: {}".format(missing_dscf.shape[0], missing_intake.shape[0], missing_cam.shape[0], invalid_ids.shape[0]))
