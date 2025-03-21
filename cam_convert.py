import pandas as pd
import util
import datetime
import argparse
from helpers import clean_sample_id, coerce_date
import warnings


def pull_archive(output_fname=util.clin_ops + 'CAM Archive/Archive Long.xlsx'):
    '''
    Pulls all CAM Archive data into a single dataframe
    '''
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*", module='openpyxl')
        cam_archive = pd.read_excel(util.clin_ops + 'CAM Archive/CAM Archive.xlsx', sheet_name=None, header=None)

    sheet_df_arch = []
    shared_header_1 = ['Date', 'Time', 'Time collected', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Visit Coordinator', 'Internal Notes']
    shared_header_2 = ['Date', 'Time', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Time Collected', 'Phlebotomist', 'Visit Coordinator', 'Internal Notes']

    # shared_header = ['Date', 'Time', 'Time collected', 'Patient Name', 'Study',
    #     'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
    #     'Sample ID', 'Visit Coordinator', 'Internal Notes']
    for _sname, df_week in cam_archive.items():
        if df_week.shape[0] < 20:
            continue
        elif df_week.shape[0] < 200:
            df_week.columns = df_week.iloc[6, :]
            date_dfs = [df_week.iloc[:, col_num:col_num + 11] for col_num, col in enumerate(df_week.columns) if type(col) == str and 'date' in col.lower()]
            for df in date_dfs:
                df.columns = shared_header_1
            sheet_df_arch.extend(date_dfs)
        else:
            df_week.columns = df_week.iloc[14, :]
            date_df = df_week.loc[:, shared_header_2].copy()
            sheet_df_arch.append(date_df)
        # df_week.columns = df_week.iloc[6, :]
        # date_dfs = [df_week.iloc[:, col_num:col_num + 11].copy() for col_num, col in enumerate(df_week.columns) if type(col) == str and 'date' in col.strip().lower()]
        # for df in date_dfs:
        #     df.columns = shared_header
        #     df = df[df['Date'] != 'Date']
        #     sheet_df_arch.append(df.dropna(subset=['Sample ID']))
        #     continue
    idx_cols = ['Date', 'Time', 'Participant ID', 'Sample ID']
    cam_arch = pd.concat(sheet_df_arch).drop_duplicates(subset=idx_cols)
    cam_arch.to_excel(output_fname, index=False)
    return cam_arch

def transform_cam(update=False, debug=False):

    long_archive = util.clin_ops + 'CAM Archive/Archive Long.xlsx'

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*", module='openpyxl')
        if update:
            cam_arch = pull_archive(output_fname=long_archive)
        else:
            cam_arch = pd.read_excel(long_archive)
        cam_active = pd.read_excel(util.clin_ops + 'CAM Clinic Schedule.xlsx', sheet_name=None, header=None)
    shared_header_1 = ['Date', 'Time', 'Time collected', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Visit Coordinator', 'Internal Notes']
    shared_header_2 = ['Date', 'Time', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Time Collected', 'Phlebotomist', 'Visit Coordinator', 'Internal Notes']

    sheet_dfs = []
    for _sname, df_week in cam_active.items():
        if df_week.shape[0] < 20:
            continue
        elif df_week.shape[0] < 200:
            df_week.columns = df_week.iloc[6, :]
            date_dfs = [df_week.iloc[:, col_num:col_num + 11] for col_num, col in enumerate(df_week.columns) if type(col) == str and 'date' in col.lower()]
            for df in date_dfs:
                df.columns = shared_header_1
            sheet_dfs.extend(date_dfs)
        else:
            df_week.columns = df_week.iloc[14, :]
            date_df = df_week.loc[:, shared_header_2].copy()
            sheet_dfs.append(date_df)

    exclude = set(['Participant ID', 'Lunch Break', 'BLOCK'])
    cam_act = pd.concat(sheet_dfs).dropna(subset=['Participant ID'])
    idx_cols = ['Date', 'Time', 'Participant ID', 'Sample ID']
    cam_both = (pd.concat([cam_act, cam_arch]).drop_duplicates(subset=idx_cols)
                    .query('`Participant ID` not in @exclude')
                    .assign(Date = lambda df: coerce_date(df['Date']).dt.date,
                            idx=lambda df: df.loc[:, idx_cols].astype(str).sum(axis=1),
                            sample_id=clean_sample_id)
                    .sort_values(by='Date', ascending=False)
                    .set_index('idx'))

    output_fname = util.cam_long
    if not debug:
        cam_both.to_excel(output_fname, index=False)
        print("Long-Form CAM Schedule written to", output_fname)
    return cam_both

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Convert CAM clinic schedule from wide multi-sheet to long single-sheet format')
    argParser.add_argument('-u', '--update', action='store_true')
    argParser.add_argument('-d', '--debug', action='store_true')
    args = argParser.parse_args()
    cam_both = transform_cam(args.update, args.debug)
