import pandas as pd
import numpy as np
import seaborn as sns
import util
from helpers import query_dscf
import argparse
import sys
from results.PARIS import paris_results

if __name__ == '__main__':
    df = paris_results().sort_values(by=['Participant ID', 'Date'])
    vaccine_dates = ['First Dose Date', 'Second Dose Date', 'Boost Date', 'Boost 2 Date', 'Boost 3 Date', 'Boost 4 Date', 'Boost 5 Date']
    vaccine_types = ['Vaccine Type', 'Boost Type', 'Boost 2 Type', 'Boost 3 Type', 'Boost 4 Type', 'Boost 5 Type']
    df['Last Known Vaccine Date'] = df.loc[:, vaccine_dates].fillna(pd.to_datetime('1/1/1900')).max(axis=1)
    df['Last Known Vaccine Type'] = df.loc[:, vaccine_types].ffill(axis=1).iloc[:, -1]
    poi = df[df['Last Known Vaccine Date'] > pd.to_datetime('9/18/23')]['Participant ID'].unique()
    df['Days to Last Known Vaccine'] = (pd.to_datetime(df['Date']) - df['Last Known Vaccine Date']).dt.days.astype(int)
    dfoi = df[df['Participant ID'].isin(poi) & (df['Days to Last Known Vaccine'] >= -120)]
    proc_cols = ['Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials', 'Log2AUC', 'Log2COV22']
    proc_names = ['Serum Volume', 'PBMCs per Aliquot (x10^6)', 'PBMC Vial Count', 'Log2AUC', 'Log2COV22']
    dfoi = dfoi.join(query_dscf(sid_list=dfoi['Sample ID'].unique()).loc[:, proc_cols[:-2]], on='Sample ID').copy()
    output_filename = util.paris + 'datasets/xbb/as_of_{}.xlsx'.format(pd.Timestamp.today().strftime("%m.%d.%y"))
    per_person = dfoi.drop_duplicates(subset='Participant ID').loc[:, ['Participant ID', 'Last Known Vaccine Date', 'Last Known Vaccine Type']]
    df_last_prevax = dfoi[dfoi['Days to Last Known Vaccine'] <= 0].drop_duplicates(subset='Participant ID', keep='last').set_index('Participant ID')
    prevax_available = df_last_prevax.index.to_numpy()
    per_person['Last Pre-XBB Sample ID'] = per_person['Participant ID'].apply(lambda pid: df_last_prevax.loc[pid, 'Sample ID'] if pid in prevax_available else 'N/A')
    per_person['Last Pre-XBB Days'] = per_person['Participant ID'].apply(lambda pid: df_last_prevax.loc[pid, 'Days to Last Known Vaccine'] if pid in prevax_available else np.nan)
    df_last_postvax = dfoi[dfoi['Days to Last Known Vaccine'] > 0].drop_duplicates(subset='Participant ID', keep='last').set_index('Participant ID')
    postvax_available = df_last_postvax.index.to_numpy()
    per_person['Last Post-XBB Sample ID'] = per_person['Participant ID'].apply(lambda pid: df_last_postvax.loc[pid, 'Sample ID'] if pid in postvax_available else 'N/A')
    per_person['Last Post-XBB Days'] = per_person['Participant ID'].apply(lambda pid: df_last_postvax.loc[pid, 'Days to Last Known Vaccine'] if pid in postvax_available else np.nan)
    postvax_count = dfoi[dfoi['Days to Last Known Vaccine'] > 0].groupby('Participant ID').count()
    per_person['Post-XBB Sample Count'] = per_person['Participant ID'].apply(lambda pid: postvax_count.loc[pid, 'Sample ID'] if pid in postvax_available else 0)
    df_for_sids = dfoi.set_index('Sample ID').copy()
    for col, new_name in zip(proc_cols, proc_names):
        per_person[new_name + ' Pre'] = per_person['Last Pre-XBB Sample ID'].apply(lambda sid: df_for_sids.loc[sid, col] if sid != 'N/A' else 'N/A')
        per_person[new_name + ' Post'] = per_person['Last Post-XBB Sample ID'].apply(lambda sid: df_for_sids.loc[sid, col] if sid != 'N/A' else 'N/A')
    # per_person['Pre-XBB PBMC Count'] = per_person['Last Pre-XBB Sample ID'].apply(lambda sid: df_for_sids.loc[sid, ''])

    with pd.ExcelWriter(output_filename) as writer:
        dfoi.to_excel(writer, sheet_name='By Sample', index=False)
        per_person.to_excel(writer, sheet_name='By Participant', index=False)
    print("XBB People written to {}".format(output_filename))
