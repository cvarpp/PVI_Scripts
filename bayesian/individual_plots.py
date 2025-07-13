import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
import matplotlib as mpl
import seaborn as sns
import statsmodels.api as sm

lod = 36

def parse_plates(workbook_name, sheet_name=0):
    """
    Reads in an Excel sheet with one row per sample containing scaled fluorescence values and returns a long-form dataframe

    Parameters:
    - workbook_name (str): The name or path of the Excel workbook containing normalized entry inhibition data
    - sheet_name (str or int, optional): The name or index of the sheet to read. Defaults to the first sheet.

    Returns:
    - df_long: A long-form DataFrame with one row per well
    """
    df_long = pd.read_excel(workbook_name, sheet_name=sheet_name).set_index(
        ['Sample ID', 'Plate', 'Column']).stack().reset_index().rename(
        columns={'level_3': 'Dilution', 0: 'Normalized'})

    df_long['Well'] = np.log(df_long['Dilution'] / 10) / np.log(3)
    df_long['log_dilution'] = np.log(df_long['Dilution'])

    print(df_long.head(1).T)
    return df_long

if __name__ == '__main__':
    scaled_input = parse_plates(workbook_name='Long_scv1.xlsx', sheet_name=0)
    df_summary = pd.read_excel('scv1_bayes_fit.xlsx', sheet_name='Full', index_col=0)
    scaled_input['top'] = df_summary.loc['top', 'mean']
    scaled_input['bottom'] = scaled_input['Plate'].apply(
        lambda x: df_summary.loc[f'bottom[{x-1}]', 'mean'])
    scaled_input['hill'] = scaled_input['Sample ID'].apply(
        lambda x: df_summary.loc[f'hill[{int(x=="Nevirapine")}]', 'mean'])
    scaled_input['id50'] = scaled_input['Sample ID'].apply(
        lambda x: df_summary.loc[f'id50[{x}]', 'mean'])
    pred_cols = [col for col in df_summary.index if col.startswith('y_hat[')]
    scaled_input['predicted'] = df_summary.loc[pred_cols, 'mean'].values
    scaled_input['residual'] = scaled_input['Normalized'] - scaled_input['predicted']
    for pid in scaled_input['Plate'].unique():
        test_df = scaled_input[scaled_input['Plate'] == pid].copy()
        xvals = np.linspace(-10, 20, 300)
        yvals_nev = [test_df.iloc[0, :]['bottom'] + (test_df.iloc[0, :]['top'] - test_df.iloc[0, :]['bottom']) / (1 + np.exp((xval - 5) * test_df.iloc[0, :]['hill'])) for xval in xvals]
        yvals = [test_df.iloc[0, :]['bottom'] + (test_df.iloc[0, :]['top'] - test_df.iloc[0, :]['bottom']) / (1 + np.exp((xval - 5) * test_df.iloc[-1, :]['hill'])) for xval in xvals]
        pal = sns.color_palette("husl", len(test_df['Sample ID'].unique()))
        fig, axes = plt.subplots(4, 3, figsize=(10, 8), sharex=True, sharey=True, layout='constrained')
        for i, sample_id in enumerate(test_df['Sample ID'].unique()):
            sample_data = test_df[test_df['Sample ID'] == sample_id]
            ax = axes[i // 3, i % 3]
            plt.sca(ax)
            ax.set_title(sample_id)
            sns.scatterplot(data=sample_data, x='log_dilution', y='Normalized', color=pal[i], ax=ax)
            if sample_id == 'Nevirapine':
                sns.lineplot(x=xvals - 5 + sample_data.iloc[0, :]['id50'], y=yvals_nev, color=pal[i], ax=ax)
            else:
                sns.lineplot(x=xvals - 5 + sample_data.iloc[0, :]['id50'], y=yvals, color=pal[i], ax=ax)
            ax.axvline(x=sample_data.iloc[0, :]['id50'], color='grey', linestyle='-', zorder=-10)
            id50 = np.exp(sample_data.iloc[0, :]['id50'])
            if id50 >= lod:
                ax.text(0.975, 0.95, f"ID50\n{id50:.0f}", transform=ax.transAxes, fontsize=10, verticalalignment='top', horizontalalignment='right')
            else:
                ax.text(0.975, 0.95, "Below LOD", transform=ax.transAxes, fontsize=10, verticalalignment='top', horizontalalignment='right', fontweight='bold')
            xticklabels = np.array([10, 30, 90, 270, 810, 2430, 7290])
            xticks = np.log(xticklabels)
            ax.vlines(sample_data['log_dilution'], ymin=sample_data['predicted'], ymax=sample_data['Normalized'], color=pal[i], linestyle='-', zorder=-5)
            plt.xticks(xticks, np.arange(7) + 2, fontsize=10)
            plt.xlabel('Well')
            plt.ylabel('Fluorescence\n(scaled)')
        plt.xlim(1, 10)
        plt.suptitle(f'Plate {pid}')
        plt.savefig(f'fits/scv1_plate_{pid}_fit.png', dpi=300)
        plt.show()
    print(scaled_input.iloc[::20].head(10))