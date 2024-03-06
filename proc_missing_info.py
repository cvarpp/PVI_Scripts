import pandas as pd
import datetime
import os
import argparse
import util
import helpers


if __name__ == '__main__':
    argparser = argparse.ArgumentParser(description='Assign blame. Benignly.')
    argparser.add_argument('-r', '--recency', type=int, default=14, help='Number of days in the past to consider recent')
    args = argparser.parse_args()

    df_all = (pd.read_excel(util.proc_ntbk, sheet_name='Specimen Dashboard', header=1)
                .assign(date=lambda df: df['Date Processed'].apply(helpers.coerce_date)))
    df = df_all[(df_all['date'] >= datetime.today() - datetime.timedelta(days=args.recency))
                 & (df_all['date'] < datetime.today())].copy()
    
    init_cols = ['Processed by (initials)', 'Cells Processed By', 'Who Moved Samples to LN?',
                 'Plasma Aliquoted by', 'Serum Aliquoted by', 'Saliva Aliquoted by']
    df['any_responsibility'] = df.loc[:, init_cols].apply(lambda row: '; '.join(row.values.astype(str)), axis=1)
    for col in init_cols:
        df['Missing: ' + col] = df[col] == 0
        df[col] = df[col].where(df[col] != 0, df['any_responsibility'])
    init_info = pd.read_excel(util.proc + 'script_data/initial_mapping.xlsx').set_index('Initials')
    future_output = {'Inits': [], 'Name': [], 'Sample ID': [], 'Sheet': [], 'Column': [], 'Current Value': [], 'Issue': [], 'Fixed': []}
    issues = []
    issue_inits = []
    # for col
    for issue, init_source in zip(issues, issue_inits):
        pass
