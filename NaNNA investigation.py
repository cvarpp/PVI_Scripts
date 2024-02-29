#%%
import pandas as pd
import re
import numpy as np
import openpyxl as opx
import util
import helpers

#%%

intake = pd.read_excel(util.intake, sheet_name="Sample Intake Log")
proc_note = pd.read_excel(util.proc_ntbk, sheet_name="Intake")

# %%
Trouble_makers =  pd.read_excel(util.tracking + "Released Samples/VIVA/NaNNA 1.7.23.xlsx", sheet_name="Missing Combined")

Samples=Trouble_makers['Sample ID'].tolist()
sample_lookup = helpers.query_dscf(sid_list=Trouble_makers['Sample ID'])
sample_lookup.to_csv('~/Downloads/quick query.csv', sep=',')
intake_filt = intake

# %%


