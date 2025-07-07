# %%
from concurrent.futures.thread import BrokenThreadPool
import pandas as pd
import matplotlib.pyplot as plot
import numpy as np
from datetime import datetime 
import seaborn as sns
import re
import colorcet as cc
import matplotlib.dates as mdates
import os
#%%

#Anything labeled "22" is from the "Now" sheets from NYC which seem to be 2022 onward 

case_seq = pd.read_csv("https://raw.githubusercontent.com/nychealth/coronavirus-data/master/variants/cases-sequenced.csv", header=0)

case_seq22 = pd.read_csv("https://raw.githubusercontent.com/nychealth/coronavirus-data/master/variants/now-cases-sequenced.csv", header=0)

var_epi = pd.read_csv("https://raw.githubusercontent.com/nychealth/coronavirus-data/master/variants/variant-epi-data.csv", 
                      header=0)

var_epi22 = pd.read_csv("https://raw.githubusercontent.com/nychealth/coronavirus-data/master/variants/now-variant-epi-data.csv", header=0)

var_class = pd.read_csv("https://raw.githubusercontent.com/nychealth/coronavirus-data/master/variants/variant-classification.csv", header=0)

variant_reference = pd.read_excel('~/library/CloudStorage/OneDrive-SharedLibraries-TheMountSinaiHospital/Simon Lab - Pathogen Surveillance/Sequencing/Lineage Mappings/variant reference.xlsx', sheet_name='Sheet1', header=0)

outpath = os.path.expanduser("~/library/CloudStorage/OneDrive-SharedLibraries-TheMountSinaiHospital/Simon Lab - Pathogen Surveillance/NYC Auto Charts/")

#%% filling things as a list this is for a later easier automatic conversion!

other=[]
omicron=[]
original_voc=[]
post_omicron=[]

var_epi_percent = var_epi.filter(regex=r'(?!)|percent|date')
var_epi_count = var_epi.filter(regex=r'(?!)|count|date')

percent_list = var_epi_percent.columns.to_list()
count_list = var_epi_count.columns.to_list()

var_epi_sets = [var_epi_count,var_epi_percent]

for dataset in var_epi_sets:        
    mapped_columns = {}
    cols_to_change_total = []
    for taxonomic_name, new_name in zip(variant_reference['scientific name'], variant_reference['common name']):
        cols_to_change = [name for name in dataset.columns if taxonomic_name in name]
        if type(cols_to_change) == "list":
            for item in cols_to_change[-1]:
                mapped_columns.update({item : new_name})
        elif cols_to_change == []:
            pass
        else:
            mapped_columns.update({cols_to_change[0] : new_name})
    dataset.rename(columns=mapped_columns,inplace=True)

    for name, col in dataset.items():
        if "date" in name:
            dataset[name] = pd.to_datetime(col)
        else:
            try:
                col.astype(int)
            except:
                print(f'column not a number: {name}')
                pass
    
day_cols=['Period start date', 'Period end date']

#%%
#stratified percent by years
years = var_epi_percent['Period end date'].dt.year.unique()

year_index = list(np.arange(0,1-(2021-years[-1])))

epi_years = []

palette = list(sns.color_palette(cc.glasbey_warm).as_hex())

color_assignments_percent={}

for n, item in enumerate(var_epi_percent.columns[2:]):
    color_assignments_percent.update({item:palette[n]})

for n, val in enumerate(year_index):
    epi_tmp = var_epi_percent[var_epi_percent['Period end date'].dt.year == 2021 + val].copy()
    if n!=0:
        previous = var_epi_percent[var_epi_percent['Period end date'].dt.year == (2021+val-1)].copy()
        last_add = previous[previous['Period end date'] == previous['Period end date'].max()].copy()
        epi_tmp = pd.concat([last_add,epi_tmp],ignore_index=True)

    epi_years.append(epi_tmp)

Master_fig = plot.figure(layout='constrained', figsize=[20,7*len(var_epi_percent['Period end date'].dt.year.unique())])

subfigs = Master_fig.subfigures(len(var_epi_percent['Period end date'].dt.year.unique()),1, hspace=0.05)

for n, fig in enumerate(subfigs):
    dat = []
    labels = []
    
    df = epi_years[n].copy()
    variants_absent = []
    
    for col in df.columns[2:]:
        if df[col].sum() <= 0:
            variants_absent.append(col)
    
    df = df.drop(labels=variants_absent,axis=1).copy()

    for name in df.columns[2:]:
        labels.append(name)

    for name in labels:
        dat.append(df[name])

    plot_colors=[]
    for column in df.columns[2:]:
        if column in color_assignments_percent.keys():
            plot_colors.append(color_assignments_percent[column])

    ax = fig.subplots(1,1)
    plot.stackplot(df['Period end date'], dat, labels=labels, colors=plot_colors, alpha=0.85, edgecolor='black')
    # sns.lineplot(data=var_epi_percent, x="Period end date", y=col, ax=axes)
    
    fig.suptitle(f"Variants in NYC: {2021+n}", fontsize=24, fontweight='bold')

    ax.set_xlim(df['Period end date'].min(), df['Period end date'].max())

    ticks = [pd.to_datetime(f'{2021+n}-01-28'),pd.to_datetime(f'{2021+n}-02-28'),pd.to_datetime(f'{2021+n}-03-28'),pd.to_datetime(f'{2021+n}-04-28'),pd.to_datetime(f'{2021+n}-05-28'),
        pd.to_datetime(f'{2021+n}-06-28'),pd.to_datetime(f'{2021+n}-07-28'),pd.to_datetime(f'{2021+n}-08-28'),pd.to_datetime(f'{2021+n}-09-28'),pd.to_datetime(f'{2021+n}-10-28'),
            pd.to_datetime(f'{2021+n}-11-28'),pd.to_datetime(f'{2021+n}-12-28')]
    
    labels=[]
    
    for item in ticks:
        labels.append(f"{item.month} \n {item.year}")
    
    ax.set_xticks(ticks)
    ax.set_xticklabels(labels)
  
    ax.set_xticks(ticks,labels,fontsize=15,fontweight='bold')

    ticks = ax.get_yticks()
    labels = ax.get_yticklabels()

    ax.set_yticks(ticks,labels,fontsize=15, fontweight='bold')

    ax.set_xlim(pd.to_datetime(f'{2021+n}-01-01'), df['Period end date'].max())
    ax.set_ylim(0,100)
    
    ax.set_ylabel("Percent of Cases", fontsize=18, fontweight='bold')

    fig.legend(bbox_to_anchor=[1.135,0.5], loc="center right", fontsize=14, fancybox=True)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b\n%Y'))

if __name__ == "__main__":
    try:
        plot.savefig(outpath + "NYC Variants Per Year Starting 2021 percent.png", dpi=300)   
        print("Plot output to sharepoint: PSP NYC Auto Charts")
    except:
        print("Error in sharepoint plotting")
        try:
            plot.savefig(os.path.expanduser("~/Documents/NYC Variants Per Year Starting 2021 Percent.png"), dpi=300)
            print("Local saved to documents folder")
        except:
            print("double Export error showing plot!")
            plot.show()
else:
    plot.show()
# %%
#stratified count by years
years = var_epi_count['Period end date'].dt.year.unique()

year_index = list(np.arange(0,1-(2021-years[-1])))

epi_years = []

palette = list(sns.color_palette(cc.glasbey_warm).as_hex())

color_assignments_count={}

for n, item in enumerate(var_epi_count.columns[2:]):
    color_assignments_count.update({item:palette[n]})

for n, val in enumerate(year_index):
    epi_tmp = var_epi_count[var_epi_count['Period end date'].dt.year == 2021 + val].copy()
    if n!=0:
        previous = var_epi_count[var_epi_count['Period end date'].dt.year == (2021+val-1)].copy()
        last_add = previous[previous['Period end date'] == previous['Period end date'].max()].copy()
        epi_tmp = pd.concat([last_add,epi_tmp],ignore_index=True)

    epi_years.append(epi_tmp)

Master_fig = plot.figure(layout='constrained', figsize=[20,7*len(var_epi_count['Period end date'].dt.year.unique())])

subfigs = Master_fig.subfigures(len(var_epi_count['Period end date'].dt.year.unique()),1, hspace=0.05)

for n, fig in enumerate(subfigs):
    dat = []
    labels = []
    
    df = epi_years[n].copy()
    variants_absent = []
    
    for col in df.columns[2:]:
        if df[col].sum() <= 0:
            variants_absent.append(col)
    
    df = df.drop(labels=variants_absent,axis=1).copy()

    for name in df.columns[2:]:
        labels.append(name)

    for name in labels:
        dat.append(df[name])

    plot_colors=[]
    for column in df.columns[2:]:
        if column in color_assignments_count.keys():
            plot_colors.append(color_assignments_count[column])

    ax = fig.subplots(1,1)
    plot.stackplot(df['Period end date'], dat, labels=labels, colors=plot_colors, alpha=0.85, edgecolor='black')
    # sns.lineplot(data=var_epi_count, x="Period end date", y=col, ax=axes)
    
    fig.suptitle(f"Variants in NYC: {2021+n}", fontsize=24, fontweight='bold')
    
    ticks = [pd.to_datetime(f'{2021+n}-01-28'),pd.to_datetime(f'{2021+n}-02-28'),pd.to_datetime(f'{2021+n}-03-28'),pd.to_datetime(f'{2021+n}-04-28'),pd.to_datetime(f'{2021+n}-05-28'),
        pd.to_datetime(f'{2021+n}-06-28'),pd.to_datetime(f'{2021+n}-07-28'),pd.to_datetime(f'{2021+n}-08-28'),pd.to_datetime(f'{2021+n}-09-28'),pd.to_datetime(f'{2021+n}-10-28'),
            pd.to_datetime(f'{2021+n}-11-28'),pd.to_datetime(f'{2021+n}-12-28')]
    
    labels=[]
    
    for item in ticks:
        labels.append(f"{item.month} \n {item.year}")

    ax.set_xticks(ticks)
    ax.set_xticklabels(labels)
  
    ax.set_xticks(ticks,labels,fontsize=15,fontweight='bold')
    ticks = ax.get_yticks()
    labels = ax.get_yticklabels()

    ax.set_yticks(ticks,labels,fontsize=15, fontweight='bold')
    ax.set_ylabel("Number of Cases", fontsize=18, fontweight='bold')
    ax.set_xlim(pd.to_datetime(f'{2021+n}-01-01'), df['Period end date'].max())

    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b\n%Y'))

    fig.legend(bbox_to_anchor=[1.135,0.5], loc="center right", fontsize=14, fancybox=True)
if __name__ == "__main__":
    try:
        plot.savefig(outpath + f"NYC Variants Per Year Starting 2021 Direct Count {datetime.today()}.png", dpi=300)   
        print("Plot output to sharepoint: PSP NYC Auto Charts")
    except:
        print("Error in sharepoint plotting")
        try:
            plot.savefig(os.path.expanduser("~/Documents/NYC Variants Per Year Starting 2021 Direct Count.png"), dpi=300)
            print("Local saved to documents folder")
        except:
            print("double Export error showing plot!")
            plot.show()
else:
    plot.show()

# %%
# All one Graph! --------------------------------------------

fig, ax = plot.subplots(1,1, figsize=[20,7])

dat = []
labels = []

color_assignments_percent={}

for n, item in enumerate(var_epi_percent.columns[2:]):
    color_assignments_percent.update({item:palette[n]})

df = var_epi_percent.copy()
    
for name in df.columns[2:]:
    labels.append(name)

for name in labels:
    dat.append(df[name])

plot_colors=[]
for column in df.columns[2:]:
    if column in color_assignments_percent.keys():
        plot_colors.append(color_assignments_percent[column])

plot.stackplot(df['Period end date'], dat, labels=labels, colors=plot_colors, alpha=0.85, edgecolor='black')

fig.suptitle(f"COVID Variants in NYC Over Time", fontsize=24)

ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=range(1,12,3)))
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b\n%Y'))

ticks = ax.get_xticks()
labels = ax.get_xticklabels()

ax.set_xticks(ticks,labels,fontweight='bold',fontsize=15)

ax.set_xlim(df['Period end date'].min(), df['Period end date'].max())

ticks = ax.get_yticks()
labels = ax.get_yticklabels()

ax.set_yticks(ticks,labels,fontweight='bold',fontsize=15)

ax.set_ylim(0,100)

ax.set_ylabel("Percent of Cases", fontsize=18, fontweight="bold")

fig.legend(bbox_to_anchor=[0.5,-0.25], loc="lower center", fontsize=14, fancybox=True, ncol=7)

if __name__ == "__main__":
    try:
        plot.savefig(outpath + f"NYC Variants Per Year Starting 2021 Contiguous Percent {datetime.today()}.png", dpi=300)   
        print("Plot output to sharepoint: PSP NYC Auto Charts")
    except:
        print("Error in sharepoint plotting")
        try:
            plot.savefig(os.path.expanduser("~/Documents/NYC Variants Per Year Starting 2021 Contiguous Percent.png"), dpi=300)
            print("Local saved to documents folder")
        except:
            print("double Export error showing plot!")
            plot.show()
else:
    plot.show()

# %%

fig, ax = plot.subplots(1,1, figsize=[20,7])

dat = []
labels = []

color_assignments_count={}

for n, item in enumerate(var_epi_count.columns[2:]):
    color_assignments_count.update({item:palette[n]})

df = var_epi_count.copy()
    
for name in df.columns[2:]:
    labels.append(name)

for name in labels:
    dat.append(df[name])

plot_colors=[]
for column in df.columns[2:]:
    if column in color_assignments_count.keys():
        plot_colors.append(color_assignments_count[column])

plot.stackplot(df['Period end date'], dat, labels=labels, colors=plot_colors, alpha=0.85, edgecolor='black')
# sns.lineplot(data=var_epi_count, x="Period end date", y=col, ax=axes)

ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=range(1,12,3)))
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b\n%Y'))

ticks = ax.get_xticks()
labels = ax.get_xticklabels()

ax.set_xticks(ticks,labels,fontweight='bold',fontsize=15)

ax.set_xlim(df['Period end date'].min(), df['Period end date'].max())

ticks = ax.get_yticks()
labels = ax.get_yticklabels()

ax.set_yticks(ticks,labels,fontweight='bold',fontsize=15)

fig.suptitle(f"COVID Variants in NYC Over Time", fontsize=24)

ax.set_ylabel("Number of Cases", fontsize=18, fontweight="bold")

ax.set_xlim(df['Period end date'].min(), df['Period end date'].max())

fig.legend(bbox_to_anchor=[0.5,-0.225], loc="lower center", fontsize=14, fancybox=True, ncol=7)

if __name__ == "__main__":
    try:
        plot.savefig(outpath + f"NYC Variants Per Year Starting 2021 Contiguous Count {datetime.today()}.png", dpi=300)   
        print("Plot output to sharepoint: PSP NYC Auto Charts")
    except:
        print("Error in sharepoint plotting")
        try:
            plot.savefig(os.path.expanduser("~/Documents/NYC Variants Per Year Starting 2021 Contiguous Count.png"), dpi=300)
            print("Local saved to documents folder")
        except:
            print("double Export error showing plot!")
            plot.show()
else:
    plot.show()


# %%
