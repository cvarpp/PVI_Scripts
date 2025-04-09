import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
import seaborn as sns
import matplotlib as mpl
import warnings
import colorcet as cc
from matplotlib import lines as mlines
import statsmodels.api as sm
import codecs
import math
import helpers as help

#From Brian -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

def data_column_grouper(data_frame):
    integer=[]
    floats=[]
    mixed=[]
    
    for i, type in enumerate(data_frame.dtypes):
        if type == int:
            integer.append(data_frame.columns[i])
        elif type == float:
            floats.append(data_frame.columns[i])
        elif type == object:
            mixed.append(data_frame.columns[i])    
    
    print("integer", "floats", "objects")
    return(integer, floats , mixed)
    
def data_split(filepath='default', excel_sheet='Sheet1', head=0):
    
    data_frame = help.data_import(filepath, excel_sheet=excel_sheet, head=head)
    
    integer, floats, mixed = data_column_grouper(data_frame)
    
    return(data_frame, integer, floats , mixed)

def date_finder(filepath="default", frame="default", excel_sheet='Sheet1', head=0):
    if filepath != "default":
        data_frame, int_cols, float_cols, obj_cols = data_split(filepath, excel_sheet=excel_sheet, head=head)
    else:
    
        int_cols, float_cols, obj_cols = data_column_grouper(frame)
        data_frame = frame
    
    dates=[]

    for name, col in data_frame.items():
        Success=[]
        for i, item in enumerate(col):
            if len(Success) >= 0.5*len(col):
                pd.to_datetime(col)
                print(f"{name} could be a date!")
                break
            try:
                Success.append(pd.to_datetime(item))
            except:
                pass
        dates.append(name)
        print(f"{name} complete")
    return(dates)

# %%-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# exploration utilities
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
def graph_maker(quant=1, columns=1, style="plots", y_size=5, x_size=7, example=False):
    axes=[]
    sfig=[]
    if quant % columns == 0:
        if style.lower() in ["plots"]:
            main_figure, axis_start = plt.subplots(nrows=int(quant/columns), ncols=columns, tight_layout=True, figsize=[int(columns)*x_size, int(quant/columns)*y_size])
            for axis_row in axis_start:
                try:
                    for ax in axis_row:
                        axes.append(ax)
                except:
                    axes.append(axis_row)
            
            if example!=False:
                numbered_charts = axes #These are here because I would like to use a deepcopy but it makes the original unusable
                for i, ax in enumerate(numbered_charts):
                    ax.text(s=str(i), x=0.5, y=0.5, fontsize=22, fontweight='bold', ha='center')
                plt.show()
                

        if style.lower() in ["figs", "figures"]:
            main_figure = plt.figure(layout='constrained', figsize=[int(columns)*x_size, int(quant/columns)*y_size])
            subfigs = main_figure.subfigures(int(quant/columns), columns, wspace=0.07)
            for figs in subfigs:
                try:
                    for fig in figs:
                        sfig.append(fig)
                except:
                    sfig.append(figs)

            if example!=False:
                numbered_charts = sfig
                for i, fig in enumerate(numbered_charts):
                    ax = fig.subplots(1,1)
                    ax.text(s=str(i), x=0.5, y=0.5, fontsize=22, fontweight='bold', ha='center')
                plt.show()
            
    else:
        print("not an even grid!")
        
        bottom = quant % columns
        main_figure = plt.figure(layout='constrained', figsize=[int(columns)*x_size, int(math.ceil(quant/columns))*y_size])
        subfigs = main_figure.subfigures(nrows=int(math.ceil(quant/columns)), wspace=0.07)
        
        if style.lower() in ["plots"]:
            for i, figs in enumerate(subfigs):
                if i <= math.floor(quant/columns)-1:
                    ax = figs.subplots(nrows=1, ncols=columns)
                    for axis in ax:
                        axes.append(axis)
                else:
                    ax = figs.subplots(nrows=1, ncols=bottom)
                    try:
                        for axis in ax:
                            axes.append(axis)
                    except:
                        axes.append(ax)
            if example!=False:
                numbered_charts = axes #These are here because I would like to use a deepcopy but it makes the original unusable
                for i, ax in enumerate(numbered_charts):
                    ax.text(s=str(i), x=0.5, y=0.5, fontsize=22, fontweight='bold', ha='center')
                plt.show()
            
                
        if style.lower() in ["figs", "figures"]:
            for i, row_fig in enumerate(subfigs):
                if i <= (math.floor(quant/columns)-1):
                    column_fig = row_fig.subfigures(nrows=1,ncols=columns)
                    for figs in column_fig:
                        sfig.append(figs)
                else:
                    column_fig = row_fig.subfigures(nrows=1,ncols=bottom)
                    try:
                        for figs in column_fig:
                            sfig.append(figs)
                    except:
                        sfig.append(column_fig)


            if example!=False:
                numbered_charts = sfig #These are here because I would like to use a deepcopy but it makes the original unusable
                for i, figs in enumerate(sfig):
                    ax = figs.subplots(1,1)
                    ax.text(s=str(i), x=0.5, y=0.5, fontsize=22, fontweight='bold', ha='center')
                plt.show()

    if example != False:
        main_figure, sfig = graph_maker(quant=quant,columns=columns,style=style, example=False, y_size=y_size,x_size=x_size)
        axes = sfig

    if style.lower() in ["figs", "figures"]:
        return(main_figure, sfig)
    
    if style.lower() in ["plots"]:
        return(main_figure, axes)

def bar_explore(filepath="default", df="default", excel_sheet='Sheet1', head=0):
    
    if filepath != "default":
        data_frame, int_cols, float_cols, obj_cols = data_split(filepath, excel_sheet=excel_sheet, head=head)
    else:
        int_cols, float_cols, obj_cols = data_split(df, data_types="all")
        data_frame = df
    
    fig_int, axes_int = graph_maker(quant=len(int_cols), columns=2, style="plots")
    fig_float, axes_float = graph_maker(quant=len(float_cols), columns=2, style="plots")
    fig_obj, axes_obj = graph_maker(quant=len(obj_cols), columns=2, style="plots")

    int_df = data_frame[int_cols]
    float_df = data_frame[float_cols]
    obj_df = data_frame[obj_cols]

    frames = [int_df,float_df,obj_df]
    plot_axes = [axes_int,axes_float,axes_obj]
    
    for df, ax_list in zip(frames, plot_axes):
        for (name, col), ax in zip(df.items(),ax_list):
            if col.dtype == object:
                try:
                    col = col.apply(codecs.decode("UTF-8"))
                except:
                    pass

            if col.dtype == object:
                if len(col.unique()) > 30:
                    ax.text(s=f"Length of unique variables \n {len(col.unique())}", x=0.5, y=0.5, fontsize=18, fontweight='bold', ha='center')
                    ax.title.set_text(name)

                else:
                    sns.countplot(x=col, ax=ax)
                    ax.title.set_text(name)
                    labels = ax.get_xticklabels()
                    ax.set_xticklabels(labels, rotation=30, ha='right')
                   
            elif len(col.unique()) <= 10:
                sns.countplot(x=col, ax=ax)
                ax.title.set_text(name)
            
            else:
                sns.histplot(x=col, ax=ax, kde=True)
                ax.title.set_text(name)

    fig_int.suptitle("Integer", fontsize=30, fontweight="bold")
    fig_float.suptitle("Floats", fontsize=30, fontweight="bold")
    fig_obj.suptitle("Objects", fontsize=30, fontweight="bold")

    plt.show(fig_int)
    plt.show(fig_float)
    plt.show(fig_obj)

    return(data_frame)

def cat_explore(filepath, df=False, excel_sheet='Sheet1', head=0):
    if filepath != "default":
        data_frame, int_cols, float_cols, obj_cols = data_split(filepath, excel_sheet=excel_sheet, head=head)
    else:
        int_cols, float_cols, obj_cols = data_split(df, data_types="all")
        data_frame = df


    maybe_cat_list = []

    for name, col in data_frame.items():
        if len(col.unique()) <= 15:
            maybe_cat_list.append(name)
    
    cat_fig, cat_axes = graph_maker(quant=len(maybe_cat_list), columns=2, style='plots')

    for name, ax in zip(maybe_cat_list, cat_axes):
        sns.countplot(data=data_frame, x=name, ax=ax)
        ax.title.set_text(name)
        labels = ax.get_xticklabels()
        ax.set_xticklabels(labels, rotation=30, ha='right')

    plt.show(cat_fig)

    return(data_frame)

def cont_explore(filepath, df=False, excel_sheet='Sheet1', head=0):
    
    if df != False:
        int_cols, float_cols, obj_cols = data_split(df, data_types="all")
        data_frame = df
    else:
        data_frame, int_cols, float_cols, obj_cols = data_split(filepath, excel_sheet=excel_sheet, head=head)

    maybe_cont_list = []

    for name, col in data_frame.items():
        if len(col.unique()) >= 15 and len(col.unique()) <= (0.75 * len(col)):
            maybe_cont_list.append(name)
    
    cont_fig, cont_axes = graph_maker(quant=len(maybe_cont_list), columns=2, style='plots')

    for name, ax in zip(maybe_cont_list, cont_axes):
        sns.histplot(data=data_frame, x=name, ax=ax)
        ax.title.set_text(name)

    plt.show(cont_fig)

    return(data_frame)
