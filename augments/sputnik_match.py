import pandas as pd
import numpy as np
import util
from helpers import query_dscf
import datetime
from results.PARIS import paris_results
from cam_convert import transform_cam

if __name__ == '__main__':
    match_df = pd.read_excel(util.paris + 'datasets/arg_match/to_match.xlsx')
    match_df['Inf to Vax'] = match_df['Days from Pre-Vaccine Infection to 1st Vaccine Dose']
    match_df['Days Between Doses'] = match_df['Days after First Dose'] - match_df['Days after Second Dose']
    match_counts = pd.crosstab(index=match_df['Participant ID'], columns=match_df['Timepoint']) > 0
    print(match_counts)
    match_visits = match_df.loc[:, ['Participant ID', 'Days after First Dose']
                                ].astype(str).groupby('Participant ID').agg(lambda df: df.str.cat(sep=','))
    match_df['Visits'] = match_df['Participant ID'].apply(lambda pid: match_visits.loc[pid, 'Days after First Dose'])
    naive = match_df[match_df['Inf to Vax'].isna()].copy()
    naive_ppl = naive.drop_duplicates(subset='Participant ID')
    conv = match_df.dropna(subset=['Inf to Vax']).copy()
    conv_ppl = conv.drop_duplicates(subset='Participant ID')
    print(conv_ppl['Inf to Vax'])
    paris_df = paris_results().sort_values(by=['Participant ID', 'Date'])
    paris_df = paris_df[paris_df['Days to 1st Vaccine Dose'].apply(lambda val: -90 <= val <= 250) &
                        (paris_df['Vaccine Type'].isin(['Pfizer', 'Moderna'])) &
                        ((paris_df['Days to 1st Vaccine Dose'] <= 0) | (paris_df['Days to 1st Vaccine Dose'] < paris_df['Days to Infection 1'].fillna(400)))].copy()
    paris_ppl = paris_df.drop_duplicates(subset='Participant ID')
    paris_ppl['Pre-Vax'] = paris_ppl['Days to 1st Vaccine Dose'] <= 0
    d15s = paris_df[(paris_df['Days to 1st Vaccine Dose'] >= 9) & (paris_df['Days to 2nd'] <= 0)]
    paris_ppl['D15'] = paris_ppl['Participant ID'].isin(d15s['Participant ID'].unique())
    d42s = paris_df[paris_df['Days to 1st Vaccine Dose'].apply(lambda val: 39 <= val <= 70)]
    paris_ppl['D42'] = paris_ppl['Participant ID'].isin(d42s['Participant ID'].unique())
    d120s = paris_df[paris_df['Days to 1st Vaccine Dose'].apply(lambda val: 100 <= val <= 150)]
    paris_ppl['D120'] = paris_ppl['Participant ID'].isin(d120s['Participant ID'].unique())
    d180s = paris_df[paris_df['Days to 1st Vaccine Dose'].apply(lambda val: 170 <= val <= 230)]
    paris_ppl['D180'] = paris_ppl['Participant ID'].isin(d180s['Participant ID'].unique())
    paris_ppl['Count'] = paris_ppl.iloc[:, -5:].astype(int).sum(axis=1)
    paris_ppl['Days Between Doses'] = paris_ppl['Days to 1st Vaccine Dose'] - paris_ppl['Days to 2nd']
    paris_visits = paris_df.loc[:, ['Participant ID', 'Days to 1st Vaccine Dose', 'Sample ID']
                                ].astype(str).groupby('Participant ID').agg(lambda df: df.str.cat(sep=','))
    paris_ppl['Visits'] = paris_ppl['Participant ID'].apply(lambda pid: paris_visits.loc[pid, 'Days to 1st Vaccine Dose'])
    paris_ppl['Samples'] = paris_ppl['Participant ID'].apply(lambda pid: paris_visits.loc[pid, 'Sample ID'])
    paris_ppl = paris_ppl[paris_ppl['D42']].sort_values(by=['Count'], ascending=[False])
    paris_naive = paris_ppl[paris_ppl['Infection Pre-Vaccine?'] == 'no'].copy()
    paris_conv = paris_ppl[paris_ppl['Infection Pre-Vaccine?'] == 'yes'].copy()
    paris_conv['Inf to Vax'] = paris_conv['Days to Infection 1'] - paris_conv['Days to 1st Vaccine Dose']
    naive_out = paris_naive.set_index('Participant ID').loc[:, ['Pre-Vax', 'D15', 'D42', 'D120', 'D180', 'Gender', 'Age',
                                    'Days Between Doses', 'Vaccine Type', 'Count', 'Visits']]
    conv_out = paris_conv.set_index('Participant ID').loc[:, ['Pre-Vax', 'D15', 'D42', 'D120', 'D180', 'Gender', 'Age',
                                  'Days Between Doses', 'Vaccine Type', 'Inf to Vax', 'Count', 'Visits']].sort_values(
                                      by=['Count', 'Inf to Vax'], ascending=[False, True]
                                  )
    print(paris_ppl.groupby(['Pre-Vax', 'D15', 'D42', 'D120', 'D180', 'Gender']).count().iloc[:, :1])
    print(paris_ppl.shape)
    print(paris_ppl[paris_ppl.iloc[:, -5:].all(axis=1)].shape)
    print(paris_ppl[paris_ppl.iloc[:, -5:].all(axis=1)].groupby(['Infection Pre-Vaccine?', 'Gender']).count().iloc[:, :1])
    print(naive_ppl.groupby('Sex at Birth').count().iloc[:, :1])
    print(conv_ppl.groupby('Sex at Birth').count().iloc[:, :1])
    with pd.ExcelWriter(util.paris + 'datasets/arg_match/matching_intermediate.xlsx') as writer:
        naive_ppl.to_excel(writer, sheet_name='Naive to Match', index=False)
        naive_out.to_excel(writer, sheet_name='Naive Options')
        conv_ppl.to_excel(writer, sheet_name='Conv to Match', index=False)
        conv_out.to_excel(writer, sheet_name='Conv Options')
