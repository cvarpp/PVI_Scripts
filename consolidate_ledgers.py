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
    accounting_dir = util.onedrive + 'Simon Lab - Management/Finance/Accounting/'
    ledger_dir = accounting_dir + 'Ledgers/'
    out_dir = accounting_dir + 'ledgers_consolidated/'
    id_col = 'Employee / P.O. / Ref'
    to_concat = []
    for fname in glob.glob('{}/*xls*'.format(ledger_dir)):
        header_df = pd.read_excel(fname, header=None)
        month_ending = header_df.iloc[1, 2]
        report_id = header_df.iloc[0, 2]
        account_number = header_df.iloc[4, 3]
        main_df = pd.read_excel(fname, header=8)
        main_df = main_df.rename(columns={'Year to Date': 'Account to Date'}).loc[:, ['Employee / ', 'Unnamed: 1', 'Date', 'Current Month', 'Account to Date', 'Journal ID', 'Description      ']].rename(
            columns={'Employee / ': id_col, 'Unnamed: 1': 'Inc/Exp Code', 'Description      ': 'Description'}
        ).assign(
            month_ending=month_ending,
            report_id=report_id,
            account_number=account_number
        )
        to_concat.append(main_df)
    output_df = pd.concat(to_concat)
    s_df = output_df[output_df[id_col].astype(str).str[:1] == 'S']
    payables_df = s_df[s_df['Journal ID'] == 'Payables-Purchase Invoices']
    accrual_df = s_df[s_df['Journal ID'] == 'Receipt Accounting-Period End Accrual']
    other_idxer = list(set(s_df.index.to_numpy()) - (set(payables_df.index.to_numpy()) | set(accrual_df.index.to_numpy())))
    other_df = s_df.loc[other_idxer, :]
    reverses_df = output_df[(output_df['Journal ID'] == 'Receipt Accounting-Period End Accrual') & (output_df['Description'].astype(str).str.lower().str.contains('reverses'))].copy()
    reverses_df[id_col] = reverses_df[id_col].fillna(reverses_df['Description'])
    with pd.ExcelWriter(out_dir + output_fname) as writer:
        output_df.to_excel(writer, sheet_name='All', freeze_panes=(1,1), index=None)
        payables_df.to_excel(writer, sheet_name='Payable POs', freeze_panes=(1,1), index=None)
        accrual_df.to_excel(writer, sheet_name='Accrual', freeze_panes=(1,1), index=None)
        reverses_df.to_excel(writer, sheet_name='Reverses', freeze_panes=(1,1), index=None)
        other_df.to_excel(writer, sheet_name='Not payable PO or accrual', freeze_panes=(1,1), index=None)
        s_df.to_excel(writer, sheet_name='S Beginning', freeze_panes=(1,1), index=None)
