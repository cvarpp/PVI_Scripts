import pandas as pd
import numpy as np
import openpyxl as opx
import argparse
import util
from helpers import query_intake


if __name__ == '__main__':
    # Intake Log - used kits,  Printing Log - printed kits,  Sample ID Master List
    # Notes: mast_list: Location - Kit Type, Study ID - Sample ID;
            #plog: Study ID - Kit Type
    mast_list = (pd.read_excel(util.tracking + "Sample ID Master List.xlsx", sheet_name='Master Sheet', header=0)
                .rename(columns={'Location': 'Kit Type', 'Study ID': 'Sample ID'}))
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=5)
            .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max', 'Study': 'Kit Type'})
            .drop(0))
    intake = query_intake()

    future_df = {'Box #': [], 'idx': []}

    for rname, row in plog.dropna(subset=['Box Min', 'Box Max']).iterrows():
        try:
            box_min = int(row['Box Min'])
            box_max = int(row['Box Max'])
        except:
            continue
        for box_num in range(box_min, box_max+1):
            future_df['Box #'].append(box_num)
            future_df['idx'].append(rname)
    

    # Left join unused_kits & df_perbox on 'Kit Type' & 'Box #'
    df_perbox = pd.DataFrame(future_df).join(plog, on='idx').set_index(['Kit Type', 'Box #'])
    df_perbox['Date Printed'] = df_perbox['Date Printed'].ffill()

    # Identify unused kit numbers for re-printing or discard
    used_kit_ids = intake.index.unique()
    # used_kit_ids = intake['Sample ID'].unique()
    unused_kits = mast_list[~mast_list['Sample ID'].astype(str).str.strip().str.upper().isin(used_kit_ids)].copy()
    # unused_kits['Printed'] = False

    # Add annotations for discrepant sample IDs
    unused_kits = unused_kits.join(df_perbox, on=['Kit Type', 'Box #'], rsuffix='_printing')
    # print([c for c in unused_kits.columns if '_printing' in c]) # debugging, keep track of overlapping columns
    # unused_kits['Date Printed'] = unused_kits['Date Printed'].fillna(unused_kits['Date Printed_printing'])
    # unused_kits['Date Printed'] = unused_kits['Date Printed'].fillna('N/A')
    # unused_kits['Num Kits Printed'] = unused_kits['# of screw cap tubes kits']
    # unused_kits['Num Kits Used'] = np.where(unused_kits['Used'].isnull(), 0, 1)
    unused_kits.drop(columns=['Box Min', 'Box Max'], inplace=True)

    # Output to Excel
    unused_kits.dropna(subset='Date Printed').to_excel(util.proc + 'unused_kits.xlsx', index=False)  # Unused, Printed
    # print("Total Unused Kit Count:", unused_kits.dropna(subset='Date Printed').shape[0]) # Debugging, how many unprinted kits do we have


