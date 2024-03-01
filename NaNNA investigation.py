#%%
import pandas as pd
import re
import sys
import argparse
import numpy as np
import openpyxl as opx
import util
import helpers
import PySimpleGUI as sg

outfile = "~/Documents/Test.xlsx"
Split_2=[]
#%%

if __name__ == '__main__':
   
    sg.theme('Dark Blue 17')

    layout = [[sg.Text('Enter 2 files to compare')],
            [sg.Radio('External File?', 'RADIO1',key='Infile',default=True), sg.Radio('Text List?', 'RADIO1', key='Text_list')],
            [sg.Text('File\nName', size=(9, 2)), sg.Input(key='filepath'), sg.FileBrowse()],
            [sg.Text('Sheet\nName', size=(9, 2)), sg.Input(key='sheetname')], [sg.Text("Sample ID", size=(9,2)), sg.Input(default_text="Sample ID", key='sampleid')],
            [sg.Text('Sample\nList', size=(9,2)), sg.Multiline(key="Sample_List",size=(30,7))],
            [sg.Text('Outfile Path'), sg.Input(key='outpath'), sg.FolderBrowse()], 
            [sg.Text('Outfile Name'), sg.Input(key='outfilename')],                  
            [sg.Submit(), sg.Cancel()]]

    window = sg.Window('File Compare', layout)

    event, values = window.read()
    window.close()

    print(values)

    if event == 'Cancel':
        quit()
    else:
        args = helpers.ValuesToClass(values)


#%%
    if args.Text_list == True:

        Split_1=re.split('\n', args.Sample_List)
        
        for sample in Split_1:
            if '\r' in sample:
                quick_split=re.split('\r', sample)
                quick_split = [i for i in quick_split if i]
                for value in quick_split:
                    Split_2.append(value)
            else:
                quick_split=sample
                Split_2.append(quick_split)
        
            print(Split_2)

            Samples=Split_2
    
    if args.Infile == True:
        Sample_List = pd.read_excel(args.filepath, sheet_name=args.sheetname)
        Samples=Sample_List[args.sampleid].tolist()
    
    Intake = helpers.query_intake(include_research=True)
    Intake2 = Intake.query("@Samples in `sample_id`")

    Processing = helpers.query_dscf(sid_list=Samples)

# %%
    
    outfile = args.outpath + '/' + args.outfilename + '.xlsx'

    try:
        book = opx.Workbook()
        writer = pd.ExcelWriter(outfile, engine= 'openpyxl')
        writer.book = book
        Intake2.to_excel( writer , sheet_name='Intake Info')
        Processing.to_excel( writer , sheet_name='DSCF Info')
        writer.close()
        print('exported to: ', outfile)
    except:
        book = opx.Workbook()
        writer = pd.ExcelWriter("~/Documents/Sample Query.xlsx", engine= 'openpyxl')
        writer.book = book
        Intake2.to_excel( writer , sheet_name='Intake Info')
        Processing.to_excel( writer , sheet_name='DSCF Info')
        writer.close()
        print('exported to: ', "~/Documents/Sample Query.xlsx")
