import pandas as pd
import argparse
import util
from helpers import query_dscf, query_intake


### Pull from Sample ID Master List & Printing Log & Data Sample Collection Form & Processing Notebook
### Create Master List indexed by (sample ID, material) with additional columns (catalog #, lot #, expiration)

# Sample ID Master List - Master Sheet
# Printing Log - LOG
# Data Sample Collection Form - Lot # Sheet, BSL2 Samples, BSL2+ Samples
# Processing Notebook - Lot #s
# Intake - Sample Intake Log


if __name__ == '__main__':

      # Rename mast_list: Location - Study ID, Study ID - Sample ID
      mast_list = (pd.read_excel(util.tracking + "Sample ID Master List.xlsx", sheet_name='Master Sheet', header=0)
                   .rename(columns={'Location': 'Study', 'Study ID': 'Sample ID'}))
    
      plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=5)
              .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max'})
              .drop(0))
    
      dscf_bsl2 = query_dscf()
      dscf_bsl2p = query_dscf()
      dscf_lot = pd.read_excel(util.dscf, sheet_name='Lot # Sheet', header=0)

      proc_lot = pd.read_excel(util.proc_ntbk, sheet_name='Lot #s', header=0)

      intake = query_intake()


      #### plog & mast_list --> merged_df1
      
      # Create plog_dict to store 'Box #' & 'index'
      plog_dict = {'Box #': [], 'index': []}
      plog_failed_conversion = []
        
      for rname, row in plog.iterrows():
            if pd.isna(row['Box Min']) or pd.isna(row['Box Max']):
                  continue
            try:
                  box_min = int(row['Box Min'])
                  box_max = int(row['Box Max'])
            except:
                  plog_failed_conversion.append(row)
                  continue
            for box_num in range(box_min, box_max + 1):
                  plog_dict['Box #'].append(box_num)
                  plog_dict['index'].append(rname)

      # Output Excel 1: plog with failed conversions (box # not in box range)
      plog_failed_df = pd.DataFrame(plog_failed_conversion)
      plog_failed_df.to_excel(util.proc + 'print_log_box_num_issue.xlsx', index=False)


      # Create plog_df from plog_dict, join plog, drop 'index', set index ('Study', 'Box #')
      plog_df = pd.DataFrame(plog_dict).join(plog, on='index').drop('index', axis='columns').set_index(['Study', 'Box #'])

      # Drop NaN from 'Study' in mast_list
      mast_list_clean = mast_list.dropna(subset=['Study'])

      # Merge mast_list_clean & plog_df
      merged_df1 = mast_list_clean.join(plog_df, on=['Study', 'Box #'], lsuffix='_mast', rsuffix='_plog')
      # merged_df1_1 = pd.merge(mast_list_clean, plog_df, on=['Kit Type', 'Box #'])    # (3502, 38) obviously wrong but why?

      # # test
      # print(mast_list_clean.shape)  # (17935, 9)
      # print(merged_df1.shape)  # (17043, 38)
      # print(merged_df1.drop_duplicates(subset=['Sample ID', 'Box #']).shape)  # (17043, 38)


      #### plog + mast_list (merged_df1) & intake
      
      # Left join unused_kits & plog_df on 'Study' & 'Box #'
      #plog_df = pd.DataFrame(plog_dict).join(mast_list, on='index').set_index(['Study', 'Box #'])
      plog_df['Date Printed'] = plog_df['Date Printed'].ffill()

      # Identify unused kit numbers for re-printing or discard
      used_kit_ids = intake.index.unique()
      # used_kit_ids = intake['Sample ID'].unique()
      unused_kits = mast_list[~mast_list['Sample ID'].astype(str).str.strip().str.upper().isin(used_kit_ids)].copy()
      # unused_kits['Printed'] = False

      # Add annotations for discrepant sample IDs
      unused_kits = unused_kits.join(plog_df, on=['Study', 'Box #'], rsuffix='_printing')
      unused_kits.drop(columns=['Box Min', 'Box Max'], inplace=True)

      # Output Excel 2: Printed but Unused Kits
      unused_kits.dropna(subset='Date Printed').to_excel(util.proc + 'print_log_unused_kits.xlsx', index=False)  # Unused, Printed



