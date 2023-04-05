#%%
import pandas as pd
import numpy as np
import openpyxl as opx
import argparse
import util

Mast_list = pd.read_excel(util.tracking + "Sample ID Master List.xlsx", sheet_name='Master Sheet', header=0)
plog = pd.read_excel(util.print_log, sheet_name='LOG', header=4)
Intake = pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=util.header_intake)
plog.columns=plog.iloc[0]
q = plog.columns.tolist()
q[4] = 'Box Max'
plog.columns = q
# %%
#despite the header statement in the load it keeps the annoying header 
plog.drop(plog.index[0], axis=0,inplace=True)
plog['idx'] = plog.index

future_df = {'Cohort': [], 'Box #': [], 'idx': []}

for rname, row in plog.iterrows():
    
    if row['Box numbers']!= row['Box numbers'] or row['Box Max'] != row['Box Max']:
        continue
    if type(row['Box numbers']) == str or type(row['Box Max']) == str:
        continue 
    
    for box_num in range(int(row['Box numbers']), int(row['Box Max'])+1):

        future_df['Cohort'].append(row['Study'])

        future_df['Box #'].append(box_num)

        future_df['idx'].append(rname)
            
df_perbox = pd.DataFrame(future_df).join(plog.loc[:, 'Screw cap tubes brand and batch #' : 'Notes'], on='idx')

# %%

Intake['Intake ind'] = True
Mast_list['Mast ind'] = True
df_perbox['plog ind'] = True

pre_combine = Mast_list.merge(df_perbox, on='Box #')

pre_combine['Sample ID'] = pre_combine['Study ID']

Combine = pre_combine.merge(Intake, on='Sample ID')
# %%

Combine.to_csv( util.script_output + 'plog link.csv', sep=',')
# %%
