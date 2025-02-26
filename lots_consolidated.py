import pandas as pd
import argparse
import util
from helpers import query_dscf, query_intake, clean_sample_id

def compress_lots():
    dscf_lot = pd.read_excel(util.dscf, sheet_name='Lot # Sheet', header=0).iloc[:, :9]
    proc_lot = pd.read_excel(util.proc_ntbk, sheet_name='Lot #s', header=0)

    selected_columns = ['Date Used', 'Material', 'Lot Number', 'EXP Date',
                        'Catalog Number', 'odate', 'Samples Affected/ COMMENTS']
    lots_reformatted = (pd.concat([dscf_lot, proc_lot]).assign(
                            Material=lambda df: df['Material'].str.strip().str.title(),
                            odate=lambda df: df['Date Used'])
                            .dropna(subset=['Date Used'])
                            .sort_values(by=['Date Used', 'Material'])
                            .drop_duplicates(['Date Used', 'Material'])
                            .set_index(['Date Used', 'Material']).fillna('Unavailable')
                            .unstack().resample('1d').asfreq().ffill().stack(future_stack=True).reset_index()
                            .loc[:, selected_columns])
    return lots_reformatted

if __name__ == '__main__':
    lots_reformatted = compress_lots()
    output_loc = util.proc + 'script_data/lots_by_dates.xlsx'
    lots_reformatted.to_excel(output_loc, index=False)
    print("Lots written to", output_loc)

