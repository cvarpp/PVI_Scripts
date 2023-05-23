import pandas as pd
import argparse
import util
from helpers import query_dscf


### Pull from Sample ID Master List & Printing Log & Data Sample Collection Form & Processing Notebook
### Create Master List indexed by (sample ID, material) with additional columns (catalog #, lot #, expiration)

# Sample ID Master List - Master Sheet
# Printing Log - LOG
# Data Sample Collection Form - Lot # Sheet, BSL2 Samples, BSL2+ Samples
# Processing Notebook - Lot #s


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


        ### plog & mast_list

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

        # Output plog with failed conversions to Excel
        if len(plog_failed_conversion) > 0:
              plog_failed_df = pd.DataFrame(plog_failed_conversion)
              #plog_failed_df.to_excel(util.proc + 'plog_failed_conversion.xlsx', index=False)
        else:
              print('No failed conversion in Printing Log.')
    

        # Create plog_df from plog_dict, join plog, drop 'index', set index ('kit Type', 'Box #')
        plog_df = pd.DataFrame(plog_dict).join(plog, on='index').drop('index', axis='columns').set_index(['Study', 'Box #'])

        # Drop NaN from 'Kit Type' in mast_list
        mast_list_clean = mast_list.dropna(subset=['Study'])

        # Merge mast_list_clean & plog_df
        merged_df1 = mast_list_clean.join(plog_df, on=['Study', 'Box #'], lsuffix='_mast', rsuffix='_plog')
        # merged_df1_1 = pd.merge(mast_list_clean, plog_df, on=['Kit Type', 'Box #'])    # (3502, 38) obviously wrong but why?

        # # test
        # print(mast_list_clean.shape)  # (17935, 9)
        # print(merged_df1.shape)  # (17043, 38)
        # print(merged_df1.drop_duplicates(subset=['Sample ID', 'Box #']).shape)  # (17043, 38)


        ### dscf & proc_ntbk
        merged_df2 = pd.merge(dscf_lot, proc_lot, on=['Date Used', 'Lot Number'], suffixes=('_dscf', '_proc'))
        # merged_df2 = pd.merge(dscf_lot, proc_lot, on=['Date Used', 'Lot Number'], lsuffix='_dscf', rsuffix='_proc')

        merged_df2 = merged_df2.dropna(subset=['Date Used'])

        print(merged_df2.columns)


        ### Fill out merged_df2
        transformed_df2 = (merged_df2
                           .reset_index()
                           .drop_duplicates(subset=['Date Used', 'Material_dscf'])  # Remove duplicate entries
                           .set_index(['Date Used', 'Material_dscf'])
                           .fillna('Unavailable')
                           .unstack()
                           .resample('1d')
                           .asfreq()
                           .ffill()
                           .stack()
                           .reset_index())
        transformed_df2.to_excel(util.proc + 'print_log_test.xlsx', index=False)












    
   