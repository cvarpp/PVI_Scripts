import pandas as pd
import argparse
import util
from helpers import query_dscf, query_intake


# Sample ID Master List - Master Sheet
# Printing Log - LOG
# Intake - Sample Intake Log
# Data Sample Collection Form - Lot # Sheet, all samples
# Processing Notebook - Lot #s


if __name__ == '__main__':

      # Rename mast_list: Location - Kit Type, Study ID - Sample ID; plog: Study - Kit Type
      mast_list = (pd.read_excel(util.tracking + "Sample ID Master List.xlsx", sheet_name='Master Sheet', header=0)
                   .rename(columns={'Location': 'Kit Type', 'Study ID': 'Sample ID'}))
    
      plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=5)
              .rename(columns={'Box numbers':'Box Min', 'Unnamed: 4':'Box Max', 'Study':'Kit Type'})
              .drop(0))
      
      all_samples = query_dscf()

      dscf_lot = pd.read_excel(util.dscf, sheet_name='Lot # Sheet', header=0).iloc[:, :9]

      proc_lot = pd.read_excel(util.proc_ntbk, sheet_name='Lot #s', header=0)

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


      # Create plog_df from plog_dict, join plog, drop 'index', set index ('Kit Type', 'Box #')
      plog_df = pd.DataFrame(plog_dict).join(plog, on='index').drop('index', axis='columns').set_index(['Kit Type', 'Box #'])

      # Drop NaN from 'Kit Type'
      mast_list_clean = mast_list.dropna(subset=['Kit Type'])

      # Merge mast_list_clean & plog_df
      merged_df1 = mast_list_clean.join(plog_df, on=['Kit Type', 'Box #'], lsuffix='_mast', rsuffix='_plog')
      # merged_df1_1 = pd.merge(mast_list_clean, plog_df, on=['Kit Type', 'Box #'])    # (3502, 38) obviously wrong but why?

      # print(mast_list_clean.shape)  # (17935, 9)
      # print(merged_df1.shape)  # (17043, 38)
      # print(merged_df1.columns)
      # print(merged_df1.drop_duplicates(subset=['Sample ID', 'Box #']).shape)  # (17043, 38)


      #### plog & mast_list & intake ==>> Printed Unused Kits
      
      # Left join unused_kits & plog_df on 'Kit Type' & 'Box #'

      # plog_df = pd.DataFrame(plog_dict).join(mast_list, on='index').set_index(['Kit Type', 'Box #'])
      plog_df['Date Printed'] = plog_df['Date Printed'].ffill()

      # Identify unused kit numbers for re-printing or discard
      used_kit_ids = intake.index.unique()
      # used_kit_ids = intake['Sample ID'].unique()
      unused_kits = mast_list[~mast_list['Sample ID'].astype(str).str.strip().str.upper().isin(used_kit_ids)].copy()
      # unused_kits['Printed'] = False

      # Add annotations for discrepant sample IDs
      unused_kits = unused_kits.join(plog_df, on=['Kit Type', 'Box #'], rsuffix='_printing')
      unused_kits.drop(columns=['Box Min', 'Box Max'], inplace=True)

      # Output 2: Printed but Unused Kits
      unused_kits.dropna(subset='Date Printed').to_excel(util.proc + 'print_log_unused_kits.xlsx', index=False)

      # print(dscf_lot.columns)
      # print(dscf_lot.shape)  # (224, 19)


      ### dscf_lot & proc_lot ==> Concatenated Lots per Date

      # Reformat lots on dscf_lot 7 proc_lot, indexed by (Date Used, Material)
      dscf_lot_reformatted = (dscf_lot.drop_duplicates(['Date Used', 'Material'])
                              .dropna(subset=['Date Used'])
                              .set_index(['Date Used', 'Material'])
                              .fillna('Unavailable')
                              .unstack()    # reshape df by moving index to column headers
                              .resample('1d')
                              .asfreq()    # fill in missing dates in resampled df
                              .ffill()    # forward-fill missing dates, with most recent non-missing one
                              .stack()    # revert df shape, bring column back to index
                              .reset_index())
      dscf_lot_reformatted['Material'] = dscf_lot_reformatted['Material'].apply(lambda x: x.strip().upper().title())

      # proc_lot_reformatted = (proc_lot.drop_duplicates(['Date Used', 'Material'])
      #                         .set_index(['Date Used', 'Material']).fillna('Unavailable').unstack().resample('1d')
      #                         .asfreq().ffill().stack().reset_index())
      # proc_lot_reformatted['Material'] = proc_lot_reformatted['Material'].apply(lambda x: x.strip().upper().title())

      proc_lot_reformatted = (proc_lot.drop_duplicates(['Date Used', 'Material'])
                              .dropna(subset=['Date Used'])
                              .set_index(['Date Used', 'Material']).fillna('Unavailable')
                              .unstack().resample('1d').asfreq().ffill().stack().reset_index())
      proc_lot_reformatted['Material'] = proc_lot_reformatted['Material'].apply(lambda x: x.strip().upper().title())


      
      # Keep columns: Date Used, Material, Long-Form Date, Lot Number, EXP Date, Catalog Mumber, Samples Affected/ COMMENTS
      selected_columns = ['Date Used', 'Material', 'Lot Number', 'EXP Date', 'Catalog Number', 'Samples Affected/ COMMENTS']
      dscf_lot_new = dscf_lot_reformatted[selected_columns]
      proc_lot_new = proc_lot_reformatted[selected_columns]

      # Concatenate dscf_lot & proc_lot
      concatenated_lot = pd.concat([dscf_lot_new, proc_lot_new])
      concatenated_lot.reset_index(drop=True, inplace=True)

      # # Checkpoint: proc_lot starts later than dscf_lot so overlap should be empty
      # overlap = pd.merge(dscf_lot_new, proc_lot_new, how='inner')
      # if overlap.empty:
      #       print("No overlap between dscf_lot and proc_lot.")
      # else:
      #       print("Overlap detected between dscf_lot and proc_lot. Need check.")
      # overlap.to_excel(util.proc + 'print_log_test_overlap.xlsx', index=False)

      # test (delete if lot_per_day no error)
      concatenated_lot.to_excel(util.proc + 'print_log_concatenated_lots.xlsx', index=False)

      exit(0)


      # Lot per day
      concatenated_lot['Date Used'] = pd.to_datetime(concatenated_lot['Date Used'])
      concatenated_lot = concatenated_lot.drop_duplicates()
      concatenated_lot.set_index(['Date Used', 'Material'], inplace=True)
      
      lot_dates = pd.DataFrame({'Date Used': pd.date_range(concatenated_lot['Date Used'].min(), concatenated_lot['Date Used'].max(), freq='D')})
      lot_per_day = concatenated_lot.set_index('Date Used').groupby('Material').ffill().resample('D').ffill().loc[lot_dates['Date Used']]

      # Output 3: Lot used per day
      lot_per_day.to_excel(util.proc + 'print_log_lot_per_date.xlsx', index=False)

      # test
      print(lot_per_day.shape)

      print(all_samples.columns)


      exit(0)


      # merged_df1 (mast_list + plog) & sample ==> Sample with Date Printed

      # Delete later, make df col shorter for easy read use only
      dscf_bsl2 = dscf_bsl2[["Date Processing Started", "Project", "Sample ID"]]
      dscf_bsl2p = dscf_bsl2p[["Date Processing Started", "Project", "Sample ID"]]

      # Merge merged_df1 with dscf_bsl2
      merged_bsl2 = merged_df1.join(dscf_bsl2.set_index('Sample ID'), on='Sample ID', rsuffix='_dscf_bsl2')
      merged_bsl2_alter = pd.merge(merged_df1, dscf_bsl2, on='Sample ID', suffixes=('', '_dscf_bsl2'))

      # Merge merged_df1 with dscf_bsl2p
      merged_bsl2p = merged_df1.join(dscf_bsl2p.set_index('Sample ID'), on='Sample ID', rsuffix='_dscf_bsl2p')
      merged_bsl2p_alter = pd.merge(merged_df1, dscf_bsl2p, on='Sample ID', suffixes=('', '_dscf_bsl2p'))

      # print(merged_bsl2.columns)  # (18032, 40)
      # print(merged_bsl2_alter.columns)  # (11680, 40)
      # print(merged_bsl2.columns)  # (18032, 40)
      # print(merged_bsl2.columns)  # (11680, 40)
      ### Discrepancy comes from: samples not in mast_list ???


      # Check sample's date, find matched lot







      ### Create list indexed by (sample ID, material) with columns (catalog #, lot #, expiration)

      summary_list = merged_data.set_index(['Sample ID', 'Material'])[['Catalog #', 'Lot #', 'Expiration']]

      # Output 4: List indexed by (sample ID, material) with (catalog #, lot #, expiration)
      reformatted_dscf_lot.to_excel(util.proc + 'print_log_reformatted_lots.xlsx', index=False)



      











