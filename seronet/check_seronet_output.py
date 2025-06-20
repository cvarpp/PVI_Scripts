import pandas as pd
import argparse
# import util
import os
from collections import Counter

def validate_files(input_folder, output_path, use_xlsx):
    fnames = ['baseline.csv','follow_up.csv','covid_history.csv','covid_vaccination_status.csv','treatment_history.csv','biospecimen.csv','biospecimen_test_result.csv','aliquot.csv','reagent.csv','equipment.csv','consumable.csv','cancer_cohort.csv','hiv_cohort.csv','organ_transplant_cohort.csv','autoimmune_cohort.csv','study_design.csv', 'shipping_manifest.csv']
    if use_xlsx:
        ending = '.xlsx'
        reader = pd.read_excel
    else:
        ending = '.csv'
        reader = pd.read_csv
    if not os.path.exists(input_folder + 'submission' + ending):
        print("No submission file in input folder. Fatal error")
        exit(1)
    submit_raw = (reader(input_folder + 'submission' + ending, header=None, names=['Variable', 'Value', 'Extra'])
                    .set_index('Variable')
                    .drop('Extra', axis=1))
    print(submit_raw.head())
    to_check = []
    for fname in fnames:
        if submit_raw.loc[fname, 'Value'] == 'X':
            to_check.append(fname.split('.')[0])
    num_ppl = int(submit_raw.loc['Number of Research Participants', 'Value'])
    num_specs = int(submit_raw.loc['Number of Biospecimens', 'Value'])
    print(num_ppl, num_specs)
    if len(os.listdir()) != len(to_check) + 1:
        print("Mismatch between directory contents and submission file.")
        print("Files just in directory:")
        for fname in sorted(os.listdir()):
            base_name = fname.split('.')[0]
            if base_name not in to_check and base_name != 'submission':
                print(base_name)
        print("Files marked in submission but missing:")
        for fname in sorted(to_check):
            if fname + '.csv' not in os.listdir() and fname + '.xlsx' not in os.listdir() and fname + '.xlsm' not in os.listdir():
                print(fname)
    id_cols = ['Research_Participant_ID', 'Cohort', 'Visit_Number', 'Biospecimen_ID', 'Biospecimen_Type', 'Aliquot_ID']
    visit_found = {}
    biospecimen_found = {}
    aliquot_found = {}
    visit_biospecimens = {}
    all_cohorts = set()
    cohort_counts = {}
    for fname in to_check:
        fpath = input_folder + fname + ending
        try:
            df = reader(fpath).rename(columns={'Current Label': 'Aliquot_ID'})
        except:
            df = reader(input_folder + fname + '.xlsm').rename(columns={'Current Label': 'Aliquot_ID'})
        coi = [col for col in df.columns if col in id_cols]
        df_counts = df.loc[:, coi]
        for _, row in df_counts.iterrows():
            if 'Research_Participant_ID' in coi and 'Visit_Number' in coi:
                idx = (row['Research_Participant_ID'], str(row['Visit_Number']))
                if idx not in visit_found:
                    visit_found[idx] = set()
                visit_found[idx].add(fname)
                if 'Biospecimen_Type' in coi:
                    if idx not in visit_biospecimens:
                        visit_biospecimens[idx] = set()
                    visit_biospecimens[idx].add(row['Biospecimen_Type'])
            if 'Biospecimen_ID' in coi:
                idx = row['Biospecimen_ID']
                if idx not in biospecimen_found:
                    biospecimen_found[idx] = set()
                biospecimen_found[idx].add(fname)
            if 'Aliquot_ID' in coi:
                idx = row['Aliquot_ID']
                if idx not in aliquot_found:
                    aliquot_found[idx] = set()
                aliquot_found[idx].add(fname)
        if 'Cohort' in coi:
            cohort_counts[fname] = Counter(df_counts['Cohort'].to_numpy())
            for cohort in df_counts['Cohort'].drop_duplicates():
                all_cohorts.add(cohort)
    future_output = {}
    future_df = {'Participant_ID': [], 'Visit': [], 'Serum': [], 'PBMC': []}
    future_df.update({fname: [] for fname in to_check})
    for idx, fnames in visit_found.items():
        participant, visit = idx
        for fname in to_check:
            if fname in fnames:
                future_df[fname].append('X')
            else:
                future_df[fname].append(pd.NA)
        future_df['Participant_ID'].append(participant)
        future_df['Visit'].append(visit)
        if idx in visit_biospecimens.keys():
            specimens = visit_biospecimens[idx]
            for spec_type in ['Serum', 'PBMC']:
                if spec_type in specimens:
                    future_df[spec_type].append('X')
                else:
                    future_df[spec_type].append(pd.NA)
        else:
            for spec_type in ['Serum', 'PBMC']:
                future_df[spec_type].append('Unknown')
    future_output['visit_found'] = pd.DataFrame(future_df).sort_values(['Participant_ID', 'Visit']).dropna(axis=1, how='all')
    future_df = {'Biospecimen_ID': []}
    future_df.update({fname: [] for fname in to_check})
    for biospecimen, fnames in biospecimen_found.items():
        for fname in to_check:
            if fname in fnames:
                future_df[fname].append('X')
            else:
                future_df[fname].append(pd.NA)
        future_df['Biospecimen_ID'].append(biospecimen)
    future_output['biospecimen_found'] = pd.DataFrame(future_df).sort_values(['Biospecimen_ID']).dropna(axis=1, how='all')
    future_df = {'Aliquot_ID': []}
    future_df.update({fname: [] for fname in to_check})
    for aliquot, fnames in aliquot_found.items():
        for fname in to_check:
            if fname in fnames:
                future_df[fname].append('X')
            else:
                future_df[fname].append(pd.NA)
        future_df['Aliquot_ID'].append(aliquot)
    future_output['aliquot_found'] = pd.DataFrame(future_df).sort_values(['Aliquot_ID']).dropna(axis=1, how='all')
    future_df = {'File': []}
    future_df.update({cohort: [] for cohort in all_cohorts})
    for fname in to_check:
        if fname in cohort_counts.keys():
            future_df['File'].append(fname)
            for cohort in all_cohorts:
                future_df[cohort].append(cohort_counts[fname][cohort])
    future_output['cohort_count'] = pd.DataFrame(future_df)
    with pd.ExcelWriter(output_path) as writer:
        for sname, df in future_output.items():
            df.to_excel(writer, sheet_name=sname, index=False)
    print("Summary written to {}".format(output_path))
    n1 = future_output['visit_found'].drop_duplicates(subset=['Participant_ID']).shape[0]
    n2 = num_ppl
    if n1 != n2:
        print("Participant Counts:", n1, "does not equal", n2, "(!!!)")
    n1 = future_output['biospecimen_found'].drop_duplicates(subset=['Biospecimen_ID']).shape[0]
    n2 = num_specs
    if n1 != n2:
        print("Biospecimen Counts:", n1, "does not equal", n2)



if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Check a folder of SERONET excel files or csv')
    argParser.add_argument('-b', '--batch', action='store', help='Name of Batch', required=True)
    argParser.add_argument('-f', '--folder', action='store', help='Fully specified filepath to folder containing files', required=True)
    argParser.add_argument('-x', '--xlsx', action='store_true', help='Files in folder are in xls(x) format')
    args = argParser.parse_args()
    # input_folder = util.seronet_data + "Batch " + args.batch + os.sep + args.folder + os.sep
    input_folder = args.folder + os.sep
    # output_folder = util.seronet_qc
    # output_path = output_folder + '{}.xlsx'.format(args.batch)
    output_path = '{}.xlsx'.format(args.batch)
    validate_files(input_folder, output_path, args.xlsx)
