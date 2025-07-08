import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
import matplotlib as mpl
import seaborn as sns
import statsmodels.api as sm

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

    print(df_long.head(1).T)
    return df_long


def fit_logistic(df, save_diagnostics=True):
    sample_list = df['Sample ID'].unique()
    sample_count = len(sample_list)
    sample_lookup = dict(zip(sample_list, range(sample_count)))
    samples = df["sample_code"] = df['Sample ID'].replace(sample_lookup).values
    log_dilution = df['log_dilution'].to_numpy()
    plates = df['Plate'].to_numpy() - 1
    plate_ids = np.arange(5)
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
        plate_col_idx = pm.Data("plate_col_idx", plate_cols, dims="obs_id")
        nev_idx = pm.Data("nev_idx", nevs, dims="obs_id")
        y_ceil = pm.Data("y_ceil", early_top, dims="sample")
        y_floor = pm.Data("y_floor", late_bottom, dims="sample")

        top = pm.Normal('top', mu=100, sigma=1)
        hill = pm.Normal('hill', mu=1.1, sigma=0.6, dims="nevirapine")
        # sigma_1 = pm.Exponential("sigma_1", lam=0.5)
        sigma_2 = pm.Exponential("sigma_2", lam=0.1)

        mu_b = pm.Normal('mu_b', mu=0, sigma=50)
        sigma_b = pm.Exponential("sigma_b", lam=0.5)
        z_b = pm.Normal("z_b", mu=0, sigma=1, dims="plate_col")
        bottom = pm.Deterministic("bottom", mu_b + z_b * sigma_b, dims="plate_col")

        id50 = pm.Normal('id50', mu=6, sigma=3, dims="sample")

        y_hat = pm.Deterministic("y_hat", bottom[plate_col_idx] + (top - bottom[plate_col_idx]) / (1 + pt.exp((log_dil - id50[sample_idx]) * hill[nev_idx])), dims='obs_id')
        y_hat_ceil = pm.Deterministic("y_hat_ceil", 100 / (1 + pt.exp((0 - id50) * 3)), dims='sample')
        y_hat_floor = pm.Deterministic("y_hat_floor", 100 / (1 + pt.exp((15 - id50) * 3)), dims='sample')

        sigma_proper = pm.Deterministic("sigma_proper", 1 + sigma_2 * (top - y_hat) / (top - bottom[plate_col_idx]))

        points = pm.Normal("yvals", mu=y_hat, sigma=sigma_proper, observed=scaled)
        points_ceil = pm.Normal("yvals_ceil", mu=y_hat_ceil, sigma=1, observed=y_ceil)
        points_floor = pm.Normal("yvals_floor", mu=y_hat_floor, sigma=1, observed=y_floor)
        logistic_trace = pm.sample(2000, tune=3000, cores=4)

    az.plot_trace(logistic_trace)
    plt.tight_layout()
    if save_diagnostics:
        plt.savefig('tmp.png', dpi=100)
    plt.show()

    df_summary = az.summary(logistic_trace, round_to=2)
    return df_summary

if __name__ == '__main__':
    argparser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    argparser.add_argument('-v', '--verbose', action='store_true')
    argparser.add_argument('-win', '--workbook_in', action='store', required=True)
    argparser.add_argument('-s', '--sheet', action='store', default=0)
    argparser.add_argument('-wout', '--workbook_out', action='store', required=True)
    args = argparser.parse_args()
    scaled_input = parse_plates(workbook_name=args.workbook_in, sheet_name=args.sheet)
    df_summary = fit_logistic(scaled_input, save_diagnostics=True)
    id50_idx = [idx for idx in df_summary.index if 'id50' in idx]
    mapper = {}
    for col in id50_idx:
        mapper[col] = re.search(r".*\[(.*)\]", col)[1]
    with pd.ExcelWriter(args.workbook_out) as writer:
        df_summary.loc[id50_idx, :].rename(index=mapper).to_excel(writer, sheet_name='ID50s')
        df_summary.to_excel(writer, sheet_name='Full')