import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
import matplotlib as mpl
import seaborn as sns
import statsmodels.api as sm
import os

import arviz as az
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import pymc as pm

import argparse
import re

import pytensor

pytensor.config.cxx = '/usr/bin/clang++'

import pytensor.tensor as pt
print(f"Running on PyMC v{pm.__version__}")

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

    return df_long


def fit_logistic(df, exp_tag, save_diagnostics=True):
    sample_list = df['Sample ID'].unique()
    sample_count = len(sample_list)
    sample_lookup = dict(zip(sample_list, range(sample_count)))
    samples = df["sample_code"] = df['Sample ID'].replace(sample_lookup).values
    log_dilution = df['log_dilution'].to_numpy()
    plates = df['Plate'].to_numpy() - 1
    plate_ids = df['Plate'].unique() - 1
    early_top = np.array([100] * df['Sample ID'].unique().size)
    late_bottom = np.array([0] * df['Sample ID'].unique().size)

    plate_cols = df['Plate_Col'] = (df['Plate'].to_numpy() - 1) * 12 + df['Column'].to_numpy() - 1
    plate_col_list = df['Plate_Col'].unique()

    coords = {"sample": sample_list, "plate": plate_ids, "plate_col": plate_col_list, "obs_id": np.arange(df.shape[0]), 'nevirapine': np.arange(2)}

    nevs = df['Nev'] = (df['Sample ID'] == 'Nevirapine').astype(int)
    scaled = df['Normalized'].to_numpy()

    with pm.Model(coords=coords) as logistic_model:
        log_dil = pm.Data("log_dil", log_dilution, dims="obs_id")
        sample_idx = pm.Data("sample_idx", samples, dims="obs_id")
        plate_idx = pm.Data("plate_idx", plates, dims="obs_id")
        _plate_col_idx = pm.Data("plate_col_idx", plate_cols, dims="obs_id")
        nev_idx = pm.Data("nev_idx", nevs, dims="obs_id")
        y_ceil = pm.Data("y_ceil", early_top, dims="sample")
        y_floor = pm.Data("y_floor", late_bottom, dims="sample")

        top = pm.Normal('top', mu=100, sigma=1)
        hill = pm.Normal('hill', mu=2, sigma=0.8, dims="nevirapine")
        sigma_1 = 1
        sigma_2 = pm.Exponential("sigma_2", lam=0.1)

        # mu_b = pm.Normal('mu_b', mu=0, sigma=50)
        # sigma_b = pm.Exponential("sigma_b", lam=0.5)
        # z_b = pm.Normal("z_b", mu=0, sigma=1, dims="plate")
        # bottom = pm.Deterministic("bottom", mu_b + z_b * sigma_b, dims="plate")
        bottom = pm.Normal("bottom", mu=0, sigma=20, dims="plate")

        id50 = pm.Normal('id50', mu=6, sigma=3, dims="sample")

        y_hat = pm.Deterministic("y_hat", bottom[plate_idx] + (top - bottom[plate_idx]) / (1 + pt.exp((log_dil - id50[sample_idx]) * hill[nev_idx])), dims='obs_id')
        y_hat_ceil = pm.Deterministic("y_hat_ceil", 100 / (1 + pt.exp((0 - id50) * 3)), dims='sample')
        y_hat_floor = pm.Deterministic("y_hat_floor", 100 / (1 + pt.exp((15 - id50) * 3)), dims='sample')

        sigma_proper = pm.Deterministic("sigma_proper", sigma_1 + sigma_2 * (top - y_hat) / (top - bottom[plate_idx]))

        points = pm.Normal("yvals", mu=y_hat, sigma=sigma_proper, observed=scaled)
        points_ceil = pm.Normal("yvals_ceil", mu=y_hat_ceil, sigma=1, observed=y_ceil)
        points_floor = pm.Normal("yvals_floor", mu=y_hat_floor, sigma=1, observed=y_floor)
        logistic_trace = pm.sample(2000, tune=2000, cores=4)

    az.plot_trace(logistic_trace, var_names=('mu', 'sigma_2', 'sigma_b', 'bottom', 'top', 'hill'), filter_vars='like')
    plt.tight_layout()
    if save_diagnostics:
        os.makedirs(f'logs', exist_ok=True)
        plt.savefig(f'logs/{exp_tag}.png', dpi=100)
    plt.close()

    df_summary = az.summary(logistic_trace, round_to=2)
    pred_rows = [row for row in df_summary.index if row[:6] == 'y_hat[']
    df['Pred'] = df_summary.loc[pred_rows, 'mean'].to_numpy()
    df['Resid'] = df['Normalized'] - df['Pred']
    sigma_rows = [row for row in df_summary.index if row[:13] == 'sigma_proper[']
    df['Sigma Pred'] = df_summary.loc[sigma_rows, 'mean'].to_numpy()
    df['Z Score'] = df['Resid'] / df['Sigma Pred']
    return df_summary

def plot_plates(scaled_input, df_summary, exp_tag, lod=10):
    """
    Plots the fitted logistic curves for every sample by plate.
    """
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
    os.makedirs(f'fits/{exp_tag}', exist_ok=True)
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
        plt.savefig(f'fits/{exp_tag}/plate_{pid}_fit.png', dpi=300)
        plt.close()
    return scaled_input

if __name__ == '__main__':
    argparser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    argparser.add_argument('-v', '--verbose', action='store_true')
    argparser.add_argument('-win', '--workbook_in', action='store', required=True)
    argparser.add_argument('-s', '--sheet', action='store', default=0)
    argparser.add_argument('-wout', '--workbook_out', action='store', required=True)
    argparser.add_argument('-exp', '--exp_tag', action='store', default='tmp')
    args = argparser.parse_args()
    scaled_input = parse_plates(workbook_name=args.workbook_in, sheet_name=args.sheet)
    df_summary = fit_logistic(scaled_input, args.exp_tag)
    scaled_input = plot_plates(scaled_input, df_summary, args.exp_tag)
    id50_idx = [idx for idx in df_summary.index if 'id50' in idx]
    mapper = {}
    for col in id50_idx:
        mapper[col] = re.search(r".*\[(.*)\]", col)[1]
    df_id50s = df_summary.loc[id50_idx, :].rename(index=mapper)
    df_id50s['id50'] = np.exp(df_id50s['mean'])
    with pd.ExcelWriter(args.workbook_out) as writer:
        df_id50s.to_excel(writer, sheet_name='ID50s')
        df_summary.to_excel(writer, sheet_name='Full')
        scaled_input.to_excel(writer, sheet_name='Annotated Input')