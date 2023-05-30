import pandas as pd
import argparse
import util
from helpers import query_dscf, query_intake


### Create Master List indexed by (sample ID, material) with additional columns (catalog #, lot #, expiration)
# Sample ID Master List - Master Sheet
# Printing Log - LOG
# Intake - Sample Intake Log
# Data Sample Collection Form - Lot # Sheet, BSL2 Samples, BSL2+ Samples


if __name__ == '__main__':

      # Rename mast_list: Location - Study ID, Study ID - Sample ID
      mast_list = (pd.read_excel(util.tracking + "Sample ID Master List.xlsx", sheet_name='Master Sheet', header=0)
                   .rename(columns={'Location': 'Study', 'Study ID': 'Sample ID'}))
    
      plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=5)
              .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max'})
              .drop(0))
    
      dscf_bsl2 = query_dscf()
      dscf_bsl2p = query_dscf()
      dscf_lot = pd.read_excel(util.dscf, sheet_name='Lot # Sheet', header=0).iloc[:, :9]

      # proc_lot = pd.read_excel(util.proc_ntbk, sheet_name='Lot #s', header=0)

      intake = query_intake()
      

      #### plog & mast_list --> merged_df1 ==>> Box # Not in Range
      
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

      # Output 1: Printing log box # range issue (failed conversions)
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
      # print(merged_df1.columns)
      # print(merged_df1.drop_duplicates(subset=['Sample ID', 'Box #']).shape)  # (17043, 38)



      #### plog & mast_list & intake ==>> Printed Unused Kits
      
      # Left join unused_kits & plog_df on 'Study' & 'Box #'

      # plog_df = pd.DataFrame(plog_dict).join(mast_list, on='index').set_index(['Study', 'Box #'])
      plog_df['Date Printed'] = plog_df['Date Printed'].ffill()

      # Identify unused kit numbers for re-printing or discard
      used_kit_ids = intake.index.unique()
      # used_kit_ids = intake['Sample ID'].unique()
      unused_kits = mast_list[~mast_list['Sample ID'].astype(str).str.strip().str.upper().isin(used_kit_ids)].copy()
      # unused_kits['Printed'] = False

      # Add annotations for discrepant sample IDs
      unused_kits = unused_kits.join(plog_df, on=['Study', 'Box #'], rsuffix='_printing')
      unused_kits.drop(columns=['Box Min', 'Box Max'], inplace=True)

      # Output 2: Printed Unused Kits
      unused_kits.dropna(subset='Date Printed').to_excel(util.proc + 'print_log_unused_kits.xlsx', index=False)

      # print(dscf_lot.columns)
      # print(dscf_lot.shape)  # (224, 19)


      ### dscf_lot ==> Reformat Lots

      # Reformat lots on dscf_lot indexed by (Date Used, Material)
      reformatted_dscf_lot = (dscf_lot.drop_duplicates(['Date Used', 'Material'])
                              .set_index(['Date Used', 'Material'])
                              .fillna('Unavailable')
                              .unstack()    # reshape df by moving index to column headers
                              .resample('1d')
                              .asfreq()    # fill in missing dates in resampled df
                              .ffill()    # forward-fill missing dates, with most recent non-missing one
                              .stack()    # revert df shape, bring column back to index
                              .reset_index())
      
      # Keep columns: Date Used, Material, Long-Form Date, Lot Number, EXP Date, Catalog Mumber,Samples Affected/ COMMENTS
      reformatted_columns = ['Date Used', 'Material', 'Long-Form Date', 'Lot Number', 'EXP Date', 'Catalog Number', 'Samples Affected/ COMMENTS']
      reformatted_dscf_lot = reformatted_dscf_lot[reformatted_columns]

      # Output 3: Reformatted Lots
      reformatted_dscf_lot.to_excel(util.proc + 'print_log_reformatted_lots.xlsx', index=False)



      ### Create Master List indexed by (sample ID, material) with columns (catalog #, lot #, expiration)

      # btw not all BSL2 & BSL2+ samples are in Sample ID Master List

      dscf_bsl2 = dscf_bsl2[["Date Processing Started", "Project", "Sample ID"]]
      dscf_bsl2p = dscf_bsl2p[["Date Processing Started", "Project", "Sample ID"]]

      print(dscf_bsl2.shape)  #(10880, 3)
      print(dscf_bsl2p.shape)  # (10880, 3)


      











