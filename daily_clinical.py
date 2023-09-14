import numpy as np
import pandas as pd
import datetime
import argparse
import util

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

if __name__ == '__main__':
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