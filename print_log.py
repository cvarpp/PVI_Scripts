import pandas as pd
import argparse
import util

### Pull from sample id master list and printing log to create a list representing catalog #, lot #, and expiration date 
### for different resources indexed by sample ID 

if __name__ == '__main__':
    # Read Sample ID Master List & Printing Log & Data Sample Collection Form & Processing Notebook
    # Notes: mast_list: Location - Kit Type, Study ID - Sample ID; plog: Study ID - Kit Type
    mast_list = (pd.read_excel(util.tracking + "Sample ID Master List.xlsx", sheet_name='Master Sheet', header=0)
                .rename(columns={'Location': 'Kit Type', 'Study ID': 'Sample ID'}))
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=5)
            .rename(columns={'Box numbers': 'Box Min', 'Unnamed: 4': 'Box Max', 'Study': 'Kit Type'})
            .drop(0))
    dscf = (pd.read_excel(util.dscf, sheet_name='Lot # Sheet', header=0).rename(columns={'Cohort': 'Kit Type'}))
    proc_ntbk = pd.read_excel(util.proc_ntbk, sheet_name='Lot #s', header=0)

    # Join mast_list & plog & dscf on 'Kit Type', Join proc_ntbk on Date Used & Lot Number
    merged_df = pd.merge(mast_list, plog, on='Kit Type')
    merged_df = pd.merge(merged_df, dscf, on='Kit Type')
    merged_df = pd.merge(merged_df, proc_ntbk, on=['Date Used', 'Lot Number'])

    # Select columns to master list
    master_list = ['Sample ID', 'Kit Type', 'Date Assigned', 'Box #', 'Date Printed', 'Date Used', 'Lot Number', 'EXP Date', 'Catalog Number']
    merged_df = merged_df[master_list]

    # Output to Excel
    merged_df.to_excel(util.proc + 'print_log_master.xlsx', index=False)



    
   