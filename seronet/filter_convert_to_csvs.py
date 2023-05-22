import pandas as pd
import sys
import glob
import os
import argparse

def shift_biospecimen(val, shift_df):
    res_id = val[:9]
    if res_id not in shift_df.index:
        return val
    else:
        return val[:-2] + str(int(val[-2:]) - shift_df.loc[res_id, 'Shift']).zfill(2)

def shift_visits(visit, res_id, shift_df):
    if res_id not in shift_df.index:
        return visit
    else:
        return visit - shift_df.loc[res_id, 'Shift']

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Read Excel files in input_dir and create corresponding csv files in output_dir. Filter and update biospecimen IDs along the way.')
    argParser.add_argument('-i', '--input_dir', action='store', required=True)
    argParser.add_argument('-o', '--output_dir', action='store', required=True)
    args = argParser.parse_args()
    assert os.path.exists(args.input_dir + '/biospecimen.xlsx')
    ref_df = pd.read_excel(args.input_dir + '/biospecimen.xlsx', sheet_name='Removed Biospecs').loc[:, ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Biospecimen_ID']]
    shift_df = pd.read_excel(args.input_dir + '/biospecimen.xlsx', sheet_name='Shifts').set_index('Research_Participant_ID')
    ref_biospec = ref_df['Biospecimen_ID'].unique()
    ref_visit_ppl = set([(row['Research_Participant_ID'], row['Visit_Number']) for _, row in ref_df.iterrows()])
    for fname in glob.glob('{}/*xls*'.format(args.input_dir)):
        print("Converting", fname)
        tmp_df = pd.read_excel(fname, na_filter=False, keep_default_na=False)
        if 'submission' in fname:
            tmp_df.to_csv('{}/{}.csv'.format(args.output_dir, fname.split(os.sep)[-1].split('.')[0]), index=False)
            print(fname, "converted!")
            continue
        if 'Biospecimen_ID' in tmp_df.columns:
            tmp_df = tmp_df[tmp_df['Biospecimen_ID'].isin(ref_biospec)].copy()
        elif 'baseline' in fname:
            tmp_df = tmp_df[tmp_df.apply(lambda row: (row['Research_Participant_ID'], 'Baseline(1)') in ref_visit_ppl, axis=1)].copy()
        else:
            tmp_df = tmp_df[tmp_df.apply(lambda row: (row['Research_Participant_ID'], row['Visit_Number']) in ref_visit_ppl, axis=1)].copy()
        if 'baseline' in fname:
            tmp_df.to_csv('{}/{}.csv'.format(args.output_dir, fname.split(os.sep)[-1].split('.')[0]), index=False)
            print(fname, "converted!")
            continue
        if 'Biospecimen_ID' in tmp_df.columns:
            tmp_df['Biospecimen_ID'] = tmp_df['Biospecimen_ID'].apply(lambda val: shift_biospecimen(val, shift_df))
        if 'Visit_Number' in tmp_df.columns:
            tmp_df['Visit_Number'] = tmp_df.apply(lambda row: shift_visits(row['Visit_Number'], row['Research_Participant_ID'], shift_df), axis=1)
        if 'Aliquot_ID' in tmp_df.columns:
            tmp_df['Aliquot_ID'] = tmp_df['Biospecimen_ID'].str[:] + tmp_df['Aliquot_ID'].str[-2:]
        tmp_df.to_csv('{}/{}.csv'.format(args.output_dir, fname.split(os.sep)[-1].split('.')[0]), index=False)
        print(fname, "converted!")
    print("Done!")
