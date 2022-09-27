# -*- coding: utf-8 -*-
"""
Created on Thu May 26 11:11:43 2022

@author: bmona
"""
#%%
import pandas as pd
import util
from datetime import date

if __name__ == '__main__':
    cam_archive = pd.read_excel(util.clin_ops + 'CAM Archive/CAM Archive.xlsx', sheet_name=None, header=None)

    sheet_df_arch = []

    shared_header = ['Date', 'Time', 'Coordinator Initials', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Processing location', 'Internal Notes']
    for sname, df_week in cam_archive.items():
        if df_week.shape[0] < 20:
            continue
        df_week.columns = df_week.iloc[6, :]
        date_dfs = [df_week.iloc[:, col_num:col_num + 11].copy() for col_num, col in enumerate(df_week.columns) if type(col) == str and 'date' in col.strip().lower()]
        for df in date_dfs:
            df.columns = shared_header
            df = df[df['Date'] != 'Date']
            sheet_df_arch.append(df.dropna(subset=['Sample ID']))
    cam_arch = pd.concat(sheet_df_arch).drop_duplicates()

    cam_active = pd.read_excel(util.clin_ops + 'CAM Clinic Schedule.xlsx', sheet_name=None, header=None)

    shared_header_1 = ['Date', 'Time', 'Coordinator Initials', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Processing location', 'Internal Notes']
    shared_header_2 = ['Date', 'Time', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Time collected', 'Phlebotomist', 'Coordinator Initials', 'Internal Notes']

    sheet_df1 = []
    sheet_df2 =[]

    for sname, df_week in cam_active.items():
        if df_week.shape[0] < 20:
            continue
        elif df_week.shape[0] < 200:
            df_week.columns = df_week.iloc[6, :]
            date_df1 = [df_week.iloc[:, col_num:col_num + 11] for col_num, col in enumerate(df_week.columns) if type(col) == str and 'date' in col.lower()]
            for df in date_df1:
                df.columns = shared_header_1
                df = df[df['Date'] != 'Date']
                sheet_df1.append(df.dropna(subset=['Sample ID']))
        else:
            df_week.columns = df_week.iloc[14, :]
            date_df2 = [df_week.iloc[:, col_num:col_num + 12] for col_num, col in enumerate(df_week.columns) if type(col) == str and 'date' in col.lower()]
            for df in date_df2:
                df.columns = shared_header_2
                df = df[df['Date'] != 'Date']
                sheet_df2.append(df.dropna(subset=['Sample ID']))
   
    ca1 = pd.concat(sheet_df1)
    ca2 = pd.concat(sheet_df2)
    cam_act = pd.concat([ca1, ca2]).drop_duplicates()
    cam_both = pd.concat([cam_act, cam_arch]).drop_duplicates()

    output_fname = util.clin_ops + "Long-Form CAM Schedule.xlsx"
    cam_both.to_excel(output_fname, index=False)

    print("Long-Form CAM Schedule written to", output_fname)
