import pandas as pd
import numpy as np
import os
import sys
import helpers
import warnings
import util

def legacy_combine_qpcrs():
    input_folder = util.psp + 'qPCR results/qPCR individual files'
    to_join = []
    for dirpath, _dirnames, fnames in os.walk(input_folder):
        for i, fname in enumerate(fnames):
            fpath = os.path.join(dirpath, fname)
            all_dfs = pd.read_excel(fpath, sheet_name=None, header=None)
            for sname, df in all_dfs.items():
                cols_of_interest = df[df.astype(str).apply(lambda s: ~s.str.contains('HMPV') & (s.str.contains('PV') | s.str.contains('PX')))].dropna(how='all', axis=1).columns
                if cols_of_interest.size > 0:
                    new_header = df.dropna(subset=[cols_of_interest[0]]).index[0]
                    new_col_name = df.loc[new_header, cols_of_interest[0]]
                    df = pd.read_excel(fpath, sheet_name=sname, header=new_header)
                    df = df.dropna(subset=[new_col_name]).assign(File=fname, Sheet=sname)
                    df['PV Index'] = df[new_col_name].astype(str).str.strip().str.upper()
                    df = df.drop_duplicates(subset='PV Index').set_index('PV Index')
                    to_join.append(df)
    if len(to_join) == 0:
        df_out = None
    else:
        df_out = to_join[0]
        for df in to_join[1:]:
            df_out = df_out.combine_first(df)
    return df_out

def combine_extracts():
    '''
    Concatenates all Excel files in the Micronics folder, assumes format from
    Micronics plate reader with PVID in second to last column ("Free Text").
    Does not recurse into subdirectories.
    '''
    to_concat = []
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*", module='openpyxl')
        for dirpath, _dirnames, fnames in os.walk(util.extractions + 'Micronics'):
            for fname in fnames:
                if 'xls' in fname:
                    df = pd.read_excel(dirpath + os.sep + fname).iloc[:, :7].copy()
                    df['Sheet Name'] = fname[:-5]
                    cols = [col for col in df.columns]
                    assert cols[-2] == 'Status', fname
                    cols[-3] = 'PV'
                    df.columns = cols
                    df = df.dropna(subset='PV')
                    df['PV'] = df['PV'].astype(str).str.strip().str.upper()
                    to_concat.append(df)
    df_out = pd.concat(to_concat).dropna(how='all', axis=1)
    return df_out

def combine_qpcrs():
    pass

def sars_farm():
    pass

def flu_farm():
    pass

def sequence_info():
    pass

def link_log():
    all_dfs = pd.read_excel(util.psp + 'NPS Rack Status.xlsx', sheet_name=None)
    to_concat = []
    for sname, df in all_dfs.items():
        if 'Sample ID ' in df.columns:
            df['Sample ID'] = df['Sample ID ']
        if 'Sample ID' in df.columns:
            cols_to_swap = [col for col in df[~(df.astype(str).apply(lambda s: s.str.contains('PV')) | df.astype(str).apply(lambda s: s.str.contains('PX')) | df.astype(str).apply(lambda s: s.str.contains('SE'))) & df.astype(str).apply(lambda s: s.str.len() >= 8)].dropna(how='all', axis=1).columns if col not in ['Rack']]
            if 'Accession' not in df.columns and 0 < len(cols_to_swap) < 3:
                df['Accession'] = df[cols_to_swap[0]] # a little hacky, hand verified for archive and should be irrelevant for new racks
            df['Sheet Name'] = sname
            to_concat.append(df)
    df_out = pd.concat(to_concat).assign(PV=lambda df: df['Sample ID']).dropna(subset='PV').dropna(how='all', axis=1)
    df_out['PV'] = df_out['PV'].astype(str).str.strip().str.upper()
    return df_out

def neut_pull():
    pass

if __name__ == '__main__':
    extraction_df = combine_extracts().set_index('PV')
    linking_df = link_log().set_index('PV')
    df = extraction_df.join(linking_df, rsuffix='_linking')
    col_info = pd.read_excel(util.psp + 'script_data/column_names.xlsx', sheet_name='Dashboard')
    cols_to_drop = [col for col in col_info[col_info['Action'] == 'Drop']['Column Name'].to_numpy() if col in df.columns]
    extract_cols_to_drop = [col for col in col_info[col_info['Action'] == 'Drop']['Column Name'].to_numpy() if col in extraction_df.columns]
    link_cols_to_drop = [col for col in col_info[col_info['Action'] == 'Drop']['Column Name'].to_numpy() if col in linking_df.columns]
    with pd.ExcelWriter(util.psp + 'script_output/psp_combined_draft.xlsx') as writer:
        df.drop(cols_to_drop, axis='columns').to_excel(writer, sheet_name='Combined')
        extraction_df.drop(extract_cols_to_drop, axis='columns').to_excel(writer, sheet_name='Extractions')
        linking_df.drop(link_cols_to_drop, axis='columns').to_excel(writer, sheet_name='Linking')
