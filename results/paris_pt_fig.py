import warnings
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.lines as mlines
import os
import pandas as pd
import argparse
import util
from helpers import query_intake
import datetime

def load_and_clean_sheet(excel_file, sheet_name, index_col='Participant ID', **excel_kwargs):
    df = pd.read_excel(excel_file, sheet_name=sheet_name, **excel_kwargs)
    if index_col != 'Participant ID':
        df = df.rename(columns={index_col: 'Participant ID'})
    df = df.dropna(subset=['Participant ID'])
    df['Participant ID'] = df['Participant ID'].astype(str).str.strip().str.upper()
    return df.set_index('Participant ID')

def pull_tracker():
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*extension is not supported.*")
        paris_tracker = pd.ExcelFile(util.paris + 'Patient Tracking - PARIS.xlsx')
        paris_data = load_and_clean_sheet(paris_tracker, sheet_name='Subgroups', header=4)
        paris_main = load_and_clean_sheet(paris_tracker, sheet_name='Main', index_col='Subject ID', header=8)
        paris_dems = load_and_clean_sheet(util.projects + 'PARIS/Demographics.xlsx', sheet_name='inputs', index_col='Subject ID')
        paris_data = (
            paris_data.rename(columns={'Study Status': 'Status'})
            .join(paris_dems, how='left', rsuffix='_dem')
            .join(paris_main, how='left', rsuffix='_main')
            .assign(Study='PARIS')
        )
        paris_data.rename(columns={'E-mail_main':'Email', 'Date of Birth':'DOB', 'Gender':'Sex', 'Status':'Status_sg', 'Study Status':'Status',
                                'Sample ID_main':'Baseline Sample ID', 'First Dose Date':'Dose #1 Date', 'Second Dose Date':'Dose #2 Date', 
                                'Boost Date':'Dose #3 Date', 'Boost 2 Date':'Dose #4 Date', 'Boost 3 Date':'Dose #5 Date', 
                                'Boost 4 Date':'Dose #6 Date', 'Boost 5 Date':'Dose #7 Date', 'Boost 6 Date':'Dose #8 Date',
                                'Vaccine Type':'Dose #1 Type', 'Boost Type':'Dose #3 Type', 'Boost 2 Type':'Dose #4 Type', 'Boost 3 Type':'Dose #5 Type', 
                                'Boost 4 Type':'Dose #6 Type', 'Boost 5 Type':'Dose #7 Type', 'Boost 6 Type':'Dose #8 Type'}, inplace=True)
        paris_data['Dose #2 Type'] = paris_data['Dose #1 Type'].copy()
        paris_data['Ethnicity'] = paris_data['Ethnicity: Hispanic or Latino'].apply(lambda val: 'Hispanic or Latino' if val == 'Yes' else 'Not Hispanic or Latino')
        return paris_data
    
def pull_immune_events(pid):
    tracker = pull_tracker()
    pt_info = tracker.loc[pid]
    vaccine_date_cols = ['Dose #1 Date', 'Dose #2 Date', 'Dose #3 Date', 'Dose #4 Date', 'Dose #5 Date', 
                                'Dose #6 Date', 'Dose #7 Date', 'Dose #8 Date']
    infection_date_cols = ['Infection 1 Date', 'Infection 2 Date', 'Infection 3 Date', 'Infection 4 Date']
    vax_dates = pt_info[vaccine_date_cols]
    inf_dates = pt_info[infection_date_cols]
    return vax_dates, inf_dates

def main(pid):
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        results = query_intake(participants=pid, include_research=True)
        results['Quantitative'].replace(to_replace=1, value=2.5, inplace=True)

        results.sort_values(by=['Date Collected'], inplace=True)
        all_covqa=results['Quantitative']
        first_covqa = results.loc[all_covqa.first_valid_index(), 'Date Collected']
        last_covqa = results.loc[all_covqa.last_valid_index(), 'Date Collected']

        if pd.to_datetime(results['Date Collected'].values[-1]) >= pd.to_datetime('1.18.23'):
            #Calculate Relative Figure Widths Based on Date
            covqa_length = (pd.to_datetime(last_covqa)-pd.to_datetime(first_covqa)).days

            all_cov22=results['COV22']
            first_cov22 = results.loc[all_cov22.first_valid_index(), 'Date Collected']
            last_cov22 = results.loc[all_cov22.last_valid_index(), 'Date Collected']
            cov22_length = (pd.to_datetime(last_cov22)-pd.to_datetime(first_cov22)).days

            total_length = covqa_length+cov22_length
            covqa_proportion = covqa_length/total_length
            cov22_proportion = cov22_length/total_length

            fig, axes = plt.subplots(1, 2,figsize=(10, 5), dpi=300, layout='tight', width_ratios=[covqa_proportion, cov22_proportion])

            def plot_raster(ax, dates, color):
                plt.sca(ax)
                for date in dates:
                    if pd.to_datetime(date) > pd.to_datetime('1.1.2020'):
                            ax.axvline(x=pd.to_datetime(date, errors='coerce'), color=color, lw=2, linestyle='--')
                plt.xlim(x_limits)
            
            #Left Plot
            ax=axes[0]
            plt.sca(ax)
            sns.lineplot(data=results, x='Date Collected', y='Quantitative', color='green', ax=ax, zorder=9, linewidth=2)
            sns.scatterplot(data=results, x='Date Collected', y='Quantitative', color='green', ax=ax, linewidth=.5, edgecolor='white', zorder=10, s=50)
            plt.title('Kantaro SeroKlir Assay', fontsize=14)
            xticks = ['7.1.20', '1.1.21', '7.1.21', '1.1.22', '7.1.22', '1.1.23', '7.1.23']
            xticklabels = ['', '2021', '', '2022', '',  '2023', '']
            plt.xticks([pd.to_datetime(xtick) for xtick in xticks], xticklabels, fontsize=11)
            pad = pd.Timedelta(weeks=4)
            ax.set(xlim=(first_covqa-pad, last_covqa+pad))
            plt.xlabel('Date', fontsize=11)
            plt.yscale('log', base=10)
            plt.minorticks_off()
            yticks=[5,10,100,1000]
            yticklabels=['LOD', '10', '100', '1000']
            plt.yticks(yticks, yticklabels, fontsize=11)
            plt.ylabel('Antibody Concentration (AU/mL)', fontsize=12)

            x_limits = ax.get_xlim()
            vax_dates, inf_dates = pull_immune_events(pid)
            plot_raster(ax,vax_dates,'blue')
            plot_raster(ax,inf_dates,'orange')
            vax_line = mlines.Line2D([], [], color='blue', linewidth=2, linestyle='--', label='COVID Vaccine')
            inf_line = mlines.Line2D([], [], color='orange', linewidth=2, linestyle='--', label='SARS-CoV-2 Infection')
            plt.legend(handles=[vax_line, inf_line], fontsize=11)

            #Right Plot
            ax=axes[1]
            plt.sca(ax)
            sns.lineplot(data=results, x='Date Collected', y='COV22', color='purple', ax=ax, zorder=9, linewidth=2)
            sns.scatterplot(data=results, x='Date Collected', y='COV22', color='purple', ax=ax, linewidth=.5, edgecolor='white', zorder=10, s=50)
            plt.title('Abbott AdviseDx Assay', fontsize=14)
            xticks = ['1.1.23', '7.1.23', '1.1.24', '7.1.24', '1.1.25', '7.1.25']
            xticklabels = ['2023', '', '2024', '',  '2025', '']
            plt.xticks([pd.to_datetime(xtick) for xtick in xticks], xticklabels, fontsize=11)
            ax.set(xlim=(first_cov22-pad, last_cov22+pad))
            plt.xlabel('Date', fontsize=11)
            axes[1].yaxis.set_label_position('right')
            axes[1].yaxis.tick_right()
            plt.yscale('log', base=10)
            plt.minorticks_off()
            yticks=[100,1000,10000,50000]
            yticklabels=['100', '1000', '10000', 'ULOQ']
            plt.yticks(yticks, yticklabels, fontsize=11)
            qalower, qaupper = axes[0].get_ylim()
            def convert_limits(covqa):
                cov22 = 20.88063*covqa**1.132526536
                return cov22
            ax.set(ylim=(convert_limits(qalower),convert_limits(qaupper)))
            plt.ylabel('Antibody Concentration (AU/mL)', fontsize=12)

            x_limits = ax.get_xlim()
            plot_raster(ax,vax_dates,'blue')
            plot_raster(ax,inf_dates,'orange')

            plt.savefig(os.path.expanduser('~') + '/Downloads/paris_final_{}.png'.format(pid), dpi=300)
            plt.show()

        else:
            fig, axes = plt.subplots(1, 1,figsize=(7, 5), dpi=300, layout='tight')

            def plot_raster(dates, color):
                for date in dates:
                    if pd.to_datetime(date) > pd.to_datetime('1.1.2020'):
                            plt.axvline(x=pd.to_datetime(date, errors='coerce'), color=color, lw=2, linestyle='--')
                plt.xlim(x_limits)
            
            sns.lineplot(data=results, x='Date Collected', y='Quantitative', color='green', zorder=9, linewidth=2)
            sns.scatterplot(data=results, x='Date Collected', y='Quantitative', color='green', linewidth=.5, edgecolor='white', zorder=10, s=50)
            plt.title('Kantaro SeroKlir Assay', fontsize=14)
            xticks = ['7.1.20', '1.1.21', '7.1.21', '1.1.22', '7.1.22', '1.1.23', '7.1.23']
            xticklabels = ['', '2021', '', '2022', '',  '2023', '']
            plt.xticks([pd.to_datetime(xtick) for xtick in xticks], xticklabels, fontsize=11)
            pad = pd.Timedelta(weeks=4)
            plt.xlim(first_covqa-pad, last_covqa+pad)
            plt.xlabel('Date', fontsize=11)
            plt.yscale('log', base=10)
            plt.minorticks_off()
            yticks=[5,10,100,1000]
            yticklabels=['LOD', '10', '100', '1000']
            plt.yticks(yticks, yticklabels, fontsize=11)
            plt.ylabel('Antibody Concentration (AU/mL)', fontsize=12)

            x_limits = plt.xlim()
            vax_dates, inf_dates = pull_immune_events(pid)
            plot_raster(vax_dates,'blue')
            plot_raster(inf_dates,'orange')
            vax_line = mlines.Line2D([], [], color='blue', linewidth=2, linestyle='--', label='COVID Vaccine')
            inf_line = mlines.Line2D([], [], color='orange', linewidth=2, linestyle='--', label='SARS-CoV-2 Infection')
            plt.legend(handles=[vax_line, inf_line], fontsize=11)

            plt.savefig(os.path.expanduser('~') + '/Downloads/paris_final_{}.png'.format(pid), dpi=300)
            plt.show()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Generate a figure displaying all antibody results for a given PARIS participant.')
    parser.add_argument('pid', type=str, help='PARIS Participant ID in 03374-XXX format')
    args = parser.parse_args()

    main(args.pid)