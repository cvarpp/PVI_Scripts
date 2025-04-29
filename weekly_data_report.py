# %%
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import util

# %%
df_proc = pd.read_excel(util.proc_ntbk, sheet_name='Specimen Dashboard', header=1)
df_intake = pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=0)
df_merged = pd.merge(df_proc, df_intake[['Sample ID', 'Study']], on='Sample ID', how='left')
df_report = df_merged[['Sample ID', 'Study', 'Date Processing Started', '% Viability', 'Total Cell Count (x10^6)', 'Cell Tube Volume (mL)']].dropna()

# Remove 0/NA/MISSING
df_report_cleaned = df_report[
    (df_report['Date Processing Started'] != '00:00:00') &
    (df_report['Total Cell Count (x10^6)'] != 'MISSING') &
    (df_report['% Viability'] != 'MISSING') & 
    (df_report['Cell Tube Volume (mL)'] != 0) & 
    (df_report['% Viability'] != 0)
]

# Convert 'Date Processing Started' to datetime, coerce errors to NaT & drop NaT
df_report_cleaned['Date Processing Started'] = pd.to_datetime(df_report_cleaned['Date Processing Started'], errors='coerce')
df_report_cleaned = df_report_cleaned.dropna(subset=['Date Processing Started'])

# Convert cell count & viability to numeric, drop NaN
df_report_cleaned.loc[:, 'Total Cell Count (x10^6)'] = pd.to_numeric(df_report_cleaned['Total Cell Count (x10^6)'], errors='coerce')
df_report_cleaned.loc[:, '% Viability'] = pd.to_numeric(df_report_cleaned['% Viability'], errors='coerce')
df_report_cleaned.loc[:, 'Cell Tube Volume (mL)'] = pd.to_numeric(df_report_cleaned['Cell Tube Volume (mL)'], errors='coerce')
df_report_cleaned = df_report_cleaned.dropna(subset=['% Viability', 'Total Cell Count (x10^6)', 'Cell Tube Volume (mL)'])

# Calculate Cell Count per mL of whole blood
df_report_cleaned['Cell Count (x10^6) per mL of Whole Blood'] = df_report_cleaned['Total Cell Count (x10^6)'] / df_report_cleaned['Cell Tube Volume (mL)']
df_report_cleaned.tail()

# %%
current_date = pd.Timestamp.now().normalize()
one_week_ago = current_date - pd.Timedelta(days=current_date.weekday())  # Most recent Monday to today
# one_week_ago = current_date - pd.Timedelta(days=(current_date.weekday() + 7))  # Previous Monday to today
# one_week_ago = current_date - pd.Timedelta(weeks=1) # Exact one week earlier
four_weeks_ago = current_date - pd.Timedelta(weeks=4)

df_past_week = df_report_cleaned[(df_report_cleaned['Date Processing Started'] >= one_week_ago) &
                                 (df_report_cleaned['Date Processing Started'] <= current_date)]
df_past_four_weeks = df_report_cleaned[(df_report_cleaned['Date Processing Started'] > four_weeks_ago) & 
                                       (df_report_cleaned['Date Processing Started'] <= current_date)]

study_category = {
    'PARIS': 'orange',
    'MARS': 'blue',
    'UMBRELLA - PVI': 'purple',
    'HEALTHY DONORS': 'green',
    'GAEA': 'red',
    'APOLLO': 'maroon',
    'Others': 'gray'
}

df_past_week['Category'] = df_past_week['Study'].apply(lambda x: x if x in study_category else 'Others')
df_past_four_weeks['Category'] = df_past_four_weeks['Study'].apply(lambda x: x if x in study_category else 'Others')

df_past_week = df_past_week.sort_values(by='Date Processing Started')
df_past_four_weeks = df_past_four_weeks.sort_values(by='Date Processing Started')

df_past_week.tail()

# %%
# ### Plot: % Viability in Past Week
plt.figure(figsize=(10, 6))
sns.stripplot(x='Date Processing Started', y='% Viability', hue='Category', 
              data=df_past_week, jitter=True, palette=study_category)
plt.axhline(y=90, color='blue', linestyle='--', linewidth=1)
plt.title('% Viability in Past Week')
plt.xlabel('Date')
plt.ylabel('% Viability')
plt.show()

# %%
# ### Plot: % Viability in Past 4 Weeks
plt.figure(figsize=(12, 8))
sns.stripplot(x='Date Processing Started', y='% Viability', hue='Category', 
              data=df_past_four_weeks, jitter=True, palette=study_category)
plt.axhline(y=90, color='blue', linestyle='--', linewidth=1)
plt.title('% Viability in Past 4 Weeks')
plt.xlabel('Date')
plt.ylabel('% Viability')
plt.legend(title='Study')
plt.xticks(rotation=30)
plt.show()

# %%
# ### Plot: Cell Count in Past Week
plt.figure(figsize=(10, 6))
sns.stripplot(x='Date Processing Started', y='Cell Count (x10^6) per mL of Whole Blood', hue='Category', 
              data=df_past_week, jitter=True, palette=study_category)
plt.title('Cell Count (x10^6) per mL of Whole Blood in Past Week')
plt.xlabel('Date')
plt.ylabel('Cell Count (x10^6) per mL of Whole Blood')
plt.show()

# %%
# ### Plot: Cell Count in Past 4 Weeks
plt.figure(figsize=(12, 8))
sns.stripplot(x='Date Processing Started', y='Cell Count (x10^6) per mL of Whole Blood', hue='Category', 
              data=df_past_four_weeks, jitter=True, palette=study_category)
plt.title('Cell Count (x10^6) per mL of Whole Blood in Past 4 Weeks')
plt.xlabel('Date')
plt.ylabel('Cell Count (x10^6) per mL of Whole Blood')
plt.legend(title='Study')
plt.xticks(rotation=30)
plt.show()

# %%
# ### Sample Distribution - Last Week (Bar Plot)
df_intake['Date Collected'] = pd.to_datetime(df_intake['Date Collected'], errors='coerce')
df_intake_past = df_intake[(df_intake['Date Collected'] >= one_week_ago) & (df_intake['Date Collected'] <= current_date)]

study_data = df_intake_past['Study'].value_counts().reset_index()
study_data.columns = ['Study', 'Number of Samples']

plt.figure(figsize=(10, 6))
sns.barplot(x='Study', y='Number of Samples', data=study_data, palette='Blues')
plt.title('Sample Distribution Across Studies - Last Week')
plt.xlabel('Study')
plt.ylabel('Number of Samples')
plt.xticks(rotation=10)
plt.show()

# %%
# ### Sample Distribution - Last Week (Pie Chart)
plt.figure(figsize=(8, 6))
plt.pie(study_data['Number of Samples'], labels=study_data['Study'], autopct='%1.1f%%',
        colors=plt.cm.Blues(np.linspace(0.3, 0.7, len(study_data['Study']))))
plt.title('Percentage of Samples by Study')
plt.show()
