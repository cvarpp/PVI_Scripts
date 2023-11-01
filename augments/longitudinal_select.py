
import pandas as pd
import numpy as np
import util
from helpers import query_dscf
import datetime
import seaborn as sns
from matplotlib import pyplot as plt
from results.PARIS import paris_results


if __name__ == '__main__':
    df = paris_results().reset_index().sort_values(by=['Participant ID', 'Date'])
    baselines = df.drop_duplicates(subset='Participant ID').set_index('Participant ID')
    df['Days from Baseline'] = df.apply(lambda row: int((row['Date'] - baselines.loc[row['Participant ID'], 'Date']).days), axis=1)
    participant_info = df.drop_duplicates(subset='Participant ID', keep='last').loc[:, ['Participant ID', 'Days from Baseline', 'Infection Pre-Vaccine?', 'Age', 'Gender']].rename(columns={'Days from Baseline': 'Days to Last Follow-Up'})
    visit_counts = df.groupby('Participant ID').count()
    participant_info['Total Visits'] = participant_info['Participant ID'].apply(lambda pid: visit_counts.loc[pid, 'Log2AUC'])
    df['3-week Intervals from Baseline'] = df['Days from Baseline'] // 21
    stringent_counts = df.drop_duplicates(subset=['Participant ID', '3-week Intervals from Baseline']).groupby('Participant ID').count()
    participant_info['Stringent Visit Count'] = participant_info['Participant ID'].apply(lambda pid: stringent_counts.loc[pid, 'Log2AUC'])
    length_followed_filter = participant_info['Days to Last Follow-Up'] >= 900
    number_of_visits_filter = participant_info['Stringent Visit Count'] >= 20
    age_filter = participant_info['Age'].apply(lambda val: 25 <= val <= 50)
    infection_filter = participant_info['Infection Pre-Vaccine?'].isin(['yes', 'no'])
    selection = participant_info[length_followed_filter & number_of_visits_filter &
                                 age_filter & infection_filter].copy()
    print(selection.groupby(['Gender', 'Infection Pre-Vaccine?']).count().iloc[:, :1])
    assert selection[(selection['Gender'] == 'Male') & (selection['Infection Pre-Vaccine?'] == 'no')].shape[0] == 14
    assert selection[(selection['Gender'] == 'Male') & (selection['Infection Pre-Vaccine?'] == 'yes')].shape[0] == 6
    male_naive = selection[(selection['Gender'] == 'Male') & (selection['Infection Pre-Vaccine?'] == 'no')]['Participant ID'].unique()
    male_hybrid = selection[(selection['Gender'] == 'Male') & (selection['Infection Pre-Vaccine?'] == 'yes')]['Participant ID'].unique()
    potential_female_naive = selection[(selection['Gender'] == 'Female') & (selection['Infection Pre-Vaccine?'] == 'no')]['Participant ID'].unique()
    potential_female_hybrid = selection[(selection['Gender'] == 'Female') & (selection['Infection Pre-Vaccine?'] == 'yes')]['Participant ID'].unique()
    rng = np.random.default_rng(seed=874178436837243)
    female_naive = rng.choice(potential_female_naive, 14, replace=False)
    female_hybrid = rng.choice(potential_female_hybrid, 6, replace=False)
    pids = np.concatenate((male_naive, male_hybrid, female_naive, female_hybrid))
    final = selection[selection['Participant ID'].isin(pids)]
    all_samples = df[df['Participant ID'].isin(pids)]
    output_folder = util.paris + 'datasets/NaNNAA/'
    with pd.ExcelWriter(output_folder + 'selection.xlsx') as writer:
        final.to_excel(writer, sheet_name='Participants', index=False)
        all_samples.to_excel(writer, sheet_name='Samples', index=False)
    for pid in all_samples['Participant ID'].unique():
        to_plot = all_samples[all_samples['Participant ID'] == pid]
        fig, ax = plt.subplots(figsize=(12,8))
        sns.scatterplot(data=to_plot, x='Date', y='Log2AUC')
        sns.lineplot(data=to_plot, x='Date', y='Log2AUC')
        infection_status = 'Naive' if to_plot.iloc[0, :]['Infection Pre-Vaccine?'] == 'no' else 'Hybrid Immunity'
        plt.title('{}, {}, {}'.format(pid, to_plot.iloc[0, :]['Gender'], infection_status), fontsize=16)
        plt.ylim((0, 18))
        yticks = np.arange(0,18,2)
        plt.yticks(yticks, np.exp2(yticks).astype(int), fontsize=14)
        plt.ylabel('SARS-CoV-2 Spike-Binding Antibodies (AUC)', fontsize=16)
        vax_dates = ['First Dose Date', 'Second Dose Date', 'Boost Date', 'Boost 2 Date', 'Boost 3 Date', 'Boost 4 Date', 'Boost 5 Date']
        inf_dates = ['Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date']
        for vax_col in vax_dates:
            vax = to_plot.iloc[0, :][vax_col]
            if not pd.isna(pd.to_datetime(vax, errors='coerce')):
                ax.axvline(vax, color=sns.color_palette('dark')[0], linestyle='--')
        for inf_col in inf_dates:
            inf = to_plot.iloc[0, :][inf_col]
            if not pd.isna(pd.to_datetime(inf, errors='coerce')):
                ax.axvline(inf, color=sns.color_palette('dark')[1], linestyle='--')
        plt.tight_layout()
        plt.savefig(output_folder + '{}.png'.format(pid), dpi=300)
        plt.close()

