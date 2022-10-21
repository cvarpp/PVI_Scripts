import pandas as pd
import numpy as np
import datetime
import argparse
import util
from d4_all_data import pull_from_source
from ecrabs_seronet import make_ecrabs
from clinical_forms import write_clinical
import os

if __name__ == '__main__':
    
    argParser = argparse.ArgumentParser(description='Make files for monthly data submission.')
    argParser.add_argument('-u', '--update', action='store_true')
    args = argParser.parse_args()
    if args.update:
        all_data = pull_from_source()
        ecrabs = make_ecrabs(all_data)
        write_clinical(ecrabs, 'monthly_tmp')
    else:
        all_data = pd.read_excel(util.unfiltered, keep_default_na=False)
    dfs_clin = pd.read_excel(util.cross_d4 + 'monthly_tmp.xlsx', sheet_name = None, keep_default_na=False)
    ppl_cols = ['Research_Participant_ID', 'Age', 'Race', 'Ethnicity', 'Sex_At_Birth', 'Sunday_Prior_To_Visit_1']
    baseline_dates = all_data.drop_duplicates(subset='Seronet ID').set_index('Seronet ID').loc[:, 'Date']
    baseline_sundays = baseline_dates - np.mod(baseline_dates.dt.weekday + 1, 7) * datetime.timedelta(days=1)
    ppl_data = dfs_clin['Baseline'].loc[:, ppl_cols[:-1]].join(baseline_sundays, on='Research_Participant_ID').rename(columns={'Date': ppl_cols[-1]})
    output_outer = util.cross_d4 + 'Accrual/'
    output_inner = output_outer + '{}/'.format(datetime.date.today())
    if not os.path.exists(output_inner):
        os.makedirs(output_inner)
    ppl_data.to_excel(output_inner + 'Accrual_Participant_Info.xlsx', index=False, na_rep='N/A')
    vax_cols = ['Research_Participant_ID', 'Visit_Number', 'Vaccination_Status', 'SARS-CoV-2_Vaccine_Type', 'SARS-CoV-2_Vaccination_Date_Duration_From_Visit1']
    orig_date = 'SARS-CoV-2_Vaccination_Date_Duration_From_Index'
    vax_data = dfs_clin['Vax'].loc[:, vax_cols[:-1] + [orig_date]]
    index_to_baseline = all_data.drop_duplicates(subset='Seronet ID').set_index('Seronet ID').loc[:, 'Days from Index']
    vax_data[vax_cols[-1]] = vax_data[orig_date].apply(lambda val: 0 if val == 'N/A' else val) - vax_data['Research_Participant_ID'].apply(lambda val: index_to_baseline[val])
    vax_data.loc[(vax_data[orig_date] == 'N/A'), vax_cols[-1]] = 'N/A'
    vax_data.drop(orig_date, axis=1).to_excel(output_inner + 'Accrual_Vaccination_Status.xlsx', index=False, na_rep='N/A')
    sample_cols = ['Site_Cohort_Name', 'Primary_Cohort', 'Research_Participant_ID', 'Visit_Number', 'Visit_Date_Duration_From_Visit_1', 'SARS_CoV_2_Infection_Status', 'Serum_Volume_For_FNL', 'Serum_Shipped_To_FNL', 'PBMC_Concentration', 'Num_PBMC_Vials_For_FNL', 'PBMC_Shipped_To_FNL', 'Unscheduled_Visit', 'Unscheduled_Visit_Purpose', 'Lost_To_FollowUp', 'Final_Visit', 'Collected_In_This_Reporting_Period']
    print(all_data.loc[:, ['Seronet ID', 'Cohort', 'Days from Index', 'Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials']].head())
