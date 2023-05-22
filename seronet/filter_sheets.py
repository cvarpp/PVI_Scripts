import pandas as pd
import argparse

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Filter all sheets in an excel file based on a given column and value set.')
    argParser.add_argument('-i', '--input_excel', action='store', required=True)
    argParser.add_argument('-f', '--filter_col', action='store', required=True)
    argParser.add_argument('-v', '--filter_vals', action='store', required=True, type=lambda path: pd.read_excel(path))
    args = argParser.parse_args()

    input_dfs = pd.read_excel(args.input_excel, sheet_name=None)
    filter_vals = args.filter_vals[args.filter_col].astype(str).unique()
    with pd.ExcelWriter('{}_filtered.xlsx'.format(args.input_excel.split('.xls')[0])) as writer:
        for sname, df in input_dfs.items():
            try:
                df = df[df[args.filter_col].astype(str).isin(filter_vals)]
                df.to_excel(writer, sheet_name=sname, index=False, na_rep='N/A')
            except:
                print(sname)
                continue
