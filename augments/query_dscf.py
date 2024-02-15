from helpers import query_dscf
import pandas as pd


if __name__ == '__main__':
    df = pd.read_excel('local_secrets/input.xlsx')
    df2 = query_dscf(sid_list=df['Sample ID'].astype(str).str.strip().str.upper().unique())
    df2.to_excel('local_secrets/proc_info.xlsx')
