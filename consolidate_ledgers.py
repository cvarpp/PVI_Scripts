import pandas as pd
import numpy as np
import glob
import util
import sys

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("usage: python consolidate_ledgers.py output_fname")
        exit(1)
    output_fname = sys.argv[1]
    ledger_dir = util.onedrive + 'Simon Lab - Management/Finance/Accounting/Ledgers'
    id_col = 'Employee / P.O. / Ref'
    to_concat = []
    for fname in glob.glob('{}/*xls*'.format(ledger_dir)):
        header_df = pd.read_excel(fname, header=None)
        month_ending = header_df.iloc[1, 2]
        report_id = header_df.iloc[0, 2]
        account_number = header_df.iloc[4, 3]
        main_df = pd.read_excel(fname, header=8)
        main_df = main_df.loc[:, ['Employee / ', 'Date', 'Current Month']].rename(
            columns={'Employee / ': id_col}
        ).dropna(subset=id_col).assign(
            month_ending=month_ending,
            report_id=report_id,
            account_number=account_number
        )
        to_concat.append(main_df)
    output_df = pd.concat(to_concat)
    with pd.ExcelWriter(output_fname) as writer:
        output_df.to_excel(writer, sheet_name='All', freeze_panes=(1,1), index=None)
        output_df[output_df[id_col].astype(str).str[:1] == 'S'].to_excel(writer, sheet_name='S Beginning', freeze_panes=(1,1))
