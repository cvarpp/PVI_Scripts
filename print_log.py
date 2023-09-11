import pandas as pd
import argparse
import util
from helpers import query_dscf, query_intake, clean_sample_id

if __name__ == '__main__':

      mast_list = (pd.read_excel(util.tracking + "Sample ID Master List.xlsx", sheet_name='Master Sheet', header=0)
                   .rename(columns={'Location': 'Kit Type', 'Study ID': 'Sample ID'})
                   .dropna(subset=['Kit Type'])
                   .assign(sample_id=clean_sample_id))
    
      plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=0)
              .rename(columns={'Box numbers':'Box Min', 'Unnamed: 4':'Box Max', 'Study':'Kit Type'})
              .drop(0))

      all_samples = query_dscf().reset_index().loc[:, ['sample_id', 'Date Processing Started']].copy()
      dscf_lot = pd.read_excel(util.dscf, sheet_name='Lot # Sheet', header=0).iloc[:, :9]
      proc_lot = pd.read_excel(util.proc_ntbk, sheet_name='Lot #s', header=0)
      intake = query_intake()
      
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
      pd.DataFrame(plog_failed_conversion).to_excel(util.proc + 'script_data/print_log_box_num_issue.xlsx', index=False)
      plog_df = pd.DataFrame(plog_dict).join(plog, on='index').drop('index', axis='columns').set_index(['Kit Type', 'Box #'])
      plog_df['Date Printed'] = plog_df['Date Printed'].ffill()
      plog_df.to_excel(util.proc + 'script_data/printing_lots.xlsx')

      unused_kits = (mast_list[~mast_list['sample_id'].isin(intake.index.unique())]
                        .join(plog_df, on=['Kit Type', 'Box #'], rsuffix='_printing')
                        .drop(columns=['Box Min', 'Box Max'])
                        .dropna(subset='Date Printed'))
      unused_kits.to_excel(util.proc + 'script_data/print_log_unused_kits.xlsx', index=False)

      selected_columns = ['Date Used', 'Material', 'Lot Number', 'EXP Date', 'Catalog Number', 'odate', 'Samples Affected/ COMMENTS']
      lots_reformatted = (pd.concat([dscf_lot, proc_lot]).assign(Material=lambda df: df['Material'].str.strip().str.title())
                              .assign(odate=lambda df: df['Date Used'])
                              .sort_values(by=['Date Used', 'Material'])
                              .drop_duplicates(['Date Used', 'Material'])
                              .dropna(subset=['Date Used'])
                              .set_index(['Date Used', 'Material']).fillna('Unavailable')
                              .unstack().resample('1d').asfreq().ffill().stack().reset_index()
                              .loc[:, selected_columns])
      lots_reformatted.to_excel(util.proc + 'script_data/lots_by_dates.xlsx', index=False)

      all_samples['Date Processing Started'] = pd.to_datetime(all_samples['Date Processing Started'])
      samples_with_lot = (pd.merge(all_samples, lots_reformatted, left_on='Date Processing Started', right_on='Date Used')
                              .drop_duplicates(subset=['sample_id', 'Material'])
                              .set_index(['sample_id', 'Material']))
      samples_with_lot.to_excel(util.proc + 'script_data/print_log_sample_with_lot.xlsx')

