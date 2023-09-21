import numpy as np
import pandas as pd
import datetime
import argparse
import util
from helpers import query_intake, query_dscf

def seronet_dems(dem_cols):
    iris_data = util.iris_folder + 'IRIS for D4 Long.xlsx'
    mars_data = util.mars_folder + 'MARS for D4 Long.xlsx'
    titan_data = util.titan_folder + 'TITAN for D4 Long.xlsx'
    priority_data = util.prio_folder + 'PRIORITY for D4 Long.xlsx'
    gaea_data = util.gaea_folder + 'GAEA for D4 Long.xlsx'
    dfs = []
    for source_file in [iris_data, mars_data, titan_data, priority_data, gaea_data]:
        df = pd.read_excel(source_file, sheet_name='Baseline Info')
        dfs.append(df.loc[:, dem_cols])
    return pd.concat(dfs)

def all_dems():
    dem_cols = ['Participant ID', 'Age', 'Sex_At_Birth', 'Race', 'Ethnicity', 'Height', 'Weight', 'BMI']
    paris_info = pd.read_excel(util.projects + 'PARIS/Demographics.xlsx').loc[:, ['Subject ID', 'Age at Baseline', 'NIH Race', 'NIH Ethnicity', 'NIH Sex']].rename(
        columns={'Subject ID': 'Participant ID',
                 'NIH Race': 'Race',
                 'NIH Ethnicity': 'Ethnicity',
                 'NIH Sex': 'Sex_At_Birth',
                 'Age at Baseline': 'Age'}
    )
    seronet_info = seronet_dems(dem_cols)
    shield_info = pd.read_excel(util.projects + 'SHIELD/SHIELD Tracker.xlsx', sheet_name='Summary').loc[:, ['Participant ID', 'Sex at birth', 'Race \n(from Epic)', 'Ethnicity \n(from Epic)', 'Hispanic or Latino?', 'Age']].rename(
        columns={'Sex at birth': 'Sex_At_Birth',
                 'Race (from Epic': 'Race',
                 'Hispanic or Latino?': 'Ethnicity'}
    )
    umbrella_info = pd.read_excel(util.umbrella_tracker, sheet_name='Summary').assign(
        Age=lambda df: ((pd.to_datetime(df['Baseline Visit Date'], errors='coerce') - pd.to_datetime(df['DOB'], errors='coerce')).dt.days / 465.2475).fillna(-1).astype(int).replace(-1, np.nan)
    ).rename(
        columns={'Subject ID': 'Participant ID',
                 'Sex': 'Sex_At_Birth'}
    )
    hd_info = pd.read_excel(util.clin_ops + 'Healthy Donors/Enrollment Log.xlsx', sheet_name='Demographics').rename(
        columns={'Patient ID': 'Participant ID',
                 'Gender': 'Sex_At_Birth'} # This is not really correct but race/ethnicity was also named wrong, will have to check
    )
    dem_dfs = [paris_info, seronet_info, shield_info, umbrella_info, hd_info]
    output_df = pd.concat(dem_dfs).loc[:, dem_cols].fillna('Unknown').drop_duplicates(subset='Participant ID')
    return output_df

if __name__ == '__main__':
    intake_df = query_intake(include_research=True).reset_index().sort_values(by=['participant_id', 'Date Collected'])
    total_visits = intake_df.drop_duplicates(subset=['participant_id', 'Date Collected']).groupby('participant_id').count().iloc[:, :1]
    total_visits.columns = ['Visit Count']
    last_visit = intake_df.drop_duplicates(subset=['participant_id'], keep='last').set_index('participant_id').loc[:, ['Date Collected', 'sample_id']]
    last_visit.columns = ['Date of Most Recent Visit', 'Most Recent Sample ID']
    first_visit = intake_df.drop_duplicates(subset=['participant_id'], keep='first').set_index('participant_id').loc[:, ['Date Collected', 'sample_id']]
    first_visit.columns = ['Baseline Date', 'Baseline Sample ID']
    visit_info = total_visits.join(last_visit).join(first_visit)
    visit_info['Days Followed'] = (visit_info['Date of Most Recent Visit'] - visit_info['Baseline Date']).dt.days + 1
    visit_info['Visits per Week'] = 7 * visit_info['Visit Count'] / visit_info['Days Followed']
    output_df = all_dems().set_index('Participant ID').join(visit_info).reset_index()
    hd = output_df[output_df['Participant ID'].astype(str).str.contains('16772')]
    hd_summary = pd.crosstab(index=hd['Race'], columns=[hd['Ethnicity'], hd['Sex_At_Birth']])
    umbrella = output_df[output_df['Participant ID'].astype(str).str.contains('16791')]
    umbrella_summary = pd.crosstab(index=umbrella['Race'], columns=[umbrella['Ethnicity'], umbrella['Sex_At_Birth']])
    paris = output_df[output_df['Participant ID'].astype(str).str.contains('03374')]
    paris_summary = pd.crosstab(index=paris['Race'], columns=[paris['Ethnicity'], paris['Sex_At_Birth']])
    with pd.ExcelWriter(util.cross_project + 'All Dems.xlsx') as writer:
        output_df.to_excel(writer, sheet_name='All Long', index=False)
        hd_summary.to_excel(writer, sheet_name='Healthy Donor Summary')
        umbrella_summary.to_excel(writer, sheet_name='Umbrella Summary')
        paris_summary.to_excel(writer, sheet_name='PARIS Summary')
