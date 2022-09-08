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
       
    shared_header = ['Date', 'Time', 'Coordinator Initials', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Processing location', 'Internal Notes']
    cam_arch = []    
    all_dfs = []
    sheet_df_arch = []

    print("Cam df size: {}".format(len(cam_archive)))
    for sname, df_week in cam_archive.items():
                
        if df_week.shape[0] < 20:
            continue
        df_week.columns = df_week.iloc[6, :]

        date_dfs = [df_week.iloc[:, col_num:col_num + 11] for col_num, col in enumerate(df_week.columns) if type(col) == str and 'date' in col.strip().lower()]
        
     
        for df in date_dfs:
        
            df.columns = shared_header
            
            sample_id_col = [val for val in df.columns if 'Sample ID' in val][0]
            
            df = df[df['Date'] != 'Date']
            
            sheet_df_arch.append(df.dropna(subset=[sample_id_col]))

    cam_arch = pd.concat(sheet_df_arch)

    cam_arch.drop_duplicates(inplace=True)
  
    cam_arch.to_excel(util.clin_ops + 'Long-Form CAM Archive.xlsx', index=False)

    print("Cam Archive Complete")
     
    cam_active = pd.read_excel(util.clin_ops + 'CAM Clinic Schedule.xlsx', sheet_name=None, header=None)
                   
    shared_header_1 = ['Date', 'Time', 'Coordinator Initials', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Processing location', 'Internal Notes']
    
    shared_header_2 = ['Date', 'Time', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Time collected', 'Phlebotomist', 'Coordinator Initials', 'Internal Notes']
    
    all_dfs = []
    sheet_df1 = []
    sheet_df2 =[]
    blank=[]

    print("Cam df size: {}".format(len(cam_active)))

    for sname, df_week in cam_active.items():
        
        if df_week.shape[0] < 20:
            continue
        elif df_week.shape[0] < 200:
            df_week.columns = df_week.iloc[6, :]
            med=1
            # print(df_week.columns)
        else:
            df_week.columns = df_week.iloc[14, :]
            med=0

        if med == 1:
            date_df1 = [df_week.iloc[:, col_num:col_num + 11] for col_num, col in enumerate(df_week.columns) if type(col) == str and 'date' in col.strip().lower()]
        else:
            date_df2 = [df_week.iloc[:, col_num:col_num + 12] for col_num, col in enumerate(df_week.columns) if type(col) == str and 'date' in col.strip().lower()]
        
        if med == 1:
            for df in date_df1:
            
                df.columns = shared_header_1
                        
                sample_id_col = [val for val in df.columns if 'Sample ID' in val][0]
            
                df = df[df['Date'] != 'Date']
            
                sheet_df1.append(df.dropna(subset=[sample_id_col]))

        elif med == 0:
            for df in date_df2:
            
                df.columns = shared_header_2
            
                sample_id_col = [val for val in df.columns if 'Sample ID' in val][0]
            
                df = df[df['Date'] != 'Date']
            
                sheet_df2.append(df.dropna(subset=[sample_id_col]))
   
    ca1 = pd.concat(sheet_df1)

    ca2 = pd.concat(sheet_df2)  

    cam_act = pd.concat([ca1, ca2])

    cam_act.drop_duplicates(inplace=True)
  
    cam_act.to_excel(util.clin_ops + 'Long-Form CAM active.xlsx', index=False)

    cam_both = pd.concat([cam_act, cam_arch])

    cam_both.drop_duplicates(inplace=True)

    cam_both.to_excel(util.clin_ops + "Long form Cam schedule.xlsx", index=False)

    print("done")
