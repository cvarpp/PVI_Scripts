# -*- coding: utf-8 -*-
"""
Created on Thu May 26 11:11:43 2022

@author: bmona
"""
import pandas as pd
import util
from datetime import date

if __name__ == '__main__':
            
    cam_archive = pd.read_excel(util.clin_ops + 'CAM Archive/CAM Archive.xlsx', sheet_name=None)
       
    shared_header = ['Date', 'Time', 'Coordinator Initials', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Processing location', 'Internal Notes']
    cam_arch = []    
    all_dfs = []
    sheet_dfs = []
    blank=[] 
    y=0
    
    for sname, df_week in cam_archive.items():
        
        # Returns a tuple of the dataset and compares position 0 of the tuple list to the #
        
        if df_week.shape[0] < 20:
            continue
        elif df_week.shape[0] < 200:
            df_week.columns = df_week.iloc[6, :]
            # print(df_week.columns)
        else:
            df_week.columns = df_week.iloc[14, :]
        # THE BIG CHEESE RIGHT HERE!
        # generates dfs containted within "a call" (wrong word look up later) 
        # check the numbered index location of the column adds 11 to it to form the range of the table (how wide the table of interest)
        # for the col_num and col (in parallel) assign numbers to columns in the data frame. then selecting based on if the col is a string and Date is in the column name.
        # only confusing part is the fact that you can do this all in one line
        date_dfs = [df_week.iloc[:, col_num:col_num + 11] for col_num, col in enumerate(df_week.columns) if type(col) == str and 'Date' in col]
        
        #for the df that is data dfs
        
        for df in date_dfs:
        
            df.columns = shared_header
            
            sample_id_col = [val for val in df.columns if 'Sample ID' in val][0]
            
            df = df[df['Date'] != 'Date']
            
            sheet_dfs.append(df.dropna(subset=[sample_id_col]))
                         
    cam_arch = pd.concat(sheet_dfs).drop_duplicates()
  
    cam_arch.to_excel(util.clin_ops + 'Long-Form CAM Archive.xlsx', index=False)

    print("Cam Archive Complete")

    #where the script lives:    
    
    cam_active = pd.read_excel(util.clin_ops + 'CAM Clinic Schedule.xlsx', sheet_name=None)
                  
    shared_header = ['Date', 'Time', 'Coordinator Initials', 'Patient Name', 'Study',
        'Visit Type / Samples Needed', 'New or Follow-up?', 'Participant ID',
        'Sample ID', 'Processing location', 'Internal Notes']
    
    all_dfs = []
    sheet_dfs = []
    blank=[]

    print("Cam df size: {}".format(len(cam_active)))
    
    for sname, df_week in cam_active.items():
        
        if df_week.shape[0] < 20:
            continue
        elif df_week.shape[0] < 200:
            df_week.columns = df_week.iloc[6, :]
            # print(df_week.columns)
        else:
            df_week.columns = df_week.iloc[14, :]

        # Returns a tuple of the dataset and compares position 0 of the tuple list to the #
        
        # THE BIG CHEESE RIGHT HERE!
        # generates dfs contained within "a call" (wrong word look up later) 
        # check the numbered index location of the column adds 11 to it to form the range of the table (how wide the table of interest)
        # for the col_num and col (in parallel) assign numbers to columns in the data frame. then selecting based on if the col is a string and Date is in the column name.
        # only confusing part is the fact that you can do this all in one line
        
        date_dfs = [df_week.iloc[:, col_num:col_num + 11] for col_num, col in enumerate(df_week.columns) if type(col) == str and 'Date' in col]
        #for the df that is data dfs
        
        for df in date_dfs:

            #set column headers to the common header list
            df.columns = shared_header
            
            #val must be a common value function.
            #bruh he has val 3 times why?
            
            sample_id_col = [val for val in df.columns if 'Sample ID' in val][0]
            
            df = df[df['Date'] != 'Date']
            
            sheet_dfs.append(df.dropna(subset=[sample_id_col]))

    cam_act = pd.concat(sheet_dfs).drop_duplicates()
  
    cam_act.to_excel(util.clin_ops + 'Long-Form CAM active.xlsx', index=False)

    print("Cam Active complete")

    cam_both = pd.concat([cam_arch,cam_act]).drop_duplicates()

    cam_both.to_excel(util.clin_ops + 'Long-Form CAM_both {}.xlsx'.format(date.today().strftime("%m.%d.%y"), index=False))

    print("Cam Calenders Compiled")
    print('filepath:Clinical Research Study Operations/Long-Form CAM_both {}'.format(date.today().strftime("%m.%d.%y")))
