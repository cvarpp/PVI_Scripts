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
x = True
#%%

if __name__ == '__main__':
   
    sg.theme('Dark Blue 17')

    layout = [[sg.Text('Enter 2 files to compare')],
            [sg.Radio('External File?', 'RADIO1',key='Infile',default=True), sg.Radio('Text List?', 'RADIO1', key='Text_list'), sg.Checkbox('Clinical Compliant?', enable_events=True, key='clinical')],
            [sg.Checkbox('MRN', disabled=True, visible=False, key='MRN'), sg.Checkbox('Clinical Info',  disabled=True, visible=False, key='tracker'), sg.Checkbox('Umbrella Info', disabled=True, visible=False, key='Umbrella')],
            [sg.Text('File\nName', size=(9, 2)), sg.Input(key='filepath'), sg.FileBrowse()],
            [sg.Text('Sheet\nName', size=(9, 2)), sg.Input(key='sheetname')], [sg.Text("Sample ID", size=(9,2)), sg.Input(default_text="Sample ID", key='sampleid')],
            [sg.Text('Sample\nList', size=(9,2)), sg.Multiline(key="Sample_List",size=(30,7))],
            [sg.Text('Outfile Name'), sg.Input(key='outfilename')],                  
            [sg.Submit(), sg.Cancel()]]

    window = sg.Window('Sample Query', layout)
    
    while x == True:
        event, values = window.read()

        if event == 'clinical':
            if window['clinical'].get() == True:
                window['MRN'].update(disabled=False, visible=True)
                window['tracker'].update(disabled=False, visible=True)
                window['Umbrella'].update(disabled=False, visible=True)
            else:
                window['MRN'].update(disabled=True, visible=False, value = False)
                window['tracker'].update(disabled=True, visible=False, value = False)
                window['Umbrella'].update(disabled=True, visible=False, value = False)
           
        elif event == 'Cancel':
            quit()
        
        elif event == sg.WIN_CLOSED:
            quit()
        
        if event == "Submit":
            args = helpers.ValuesToClass(values)
            window.close()
            x=False


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

    partID = Intake2['participant_id'].to_list()

    Processing = helpers.query_dscf(sid_list=Samples)

    if args.clinical == True:
        Research = helpers.query_research(sid_list=Samples)

    if args.tracker == True:
        PARIS = pd.read_excel(util.paris_tracker, sheet_name="Main", header=8).query("@partID in `Subject ID`")
        TITAN = pd.read_excel(util.titan_tracker, sheet_name="Tracker", header=4).query("@partID in `Umbrella Corresponding Participant ID`")     
        # MARS =
        # APOLLO =
        # GAEA = 

        if args.Umbrella == True:
            UMBRELLA = pd.read_excel(util.umbrella_tracker, sheet_name="Summary").query("@partID in `Subject ID`")

        # Tracker_list=[MARS,GAEA,PARIS,UMBRELLA]
        # for item in Tracker_list:
        #     try:
        #         item.query("@partID in `Subject ID`")

        #     except:
        #         item.query("@partID in `Participant ID`")
# %%
    if args.clinical== True:
        outfile = util.sample_query + args.outfilename + '.xlsx'
        book = opx.Workbook()
        writer = pd.ExcelWriter(outfile, engine= 'openpyxl')
        writer.book = book
        Intake2.to_excel( writer , sheet_name='Intake Info')
        Processing.to_excel( writer , sheet_name='DSCF Info')
        TITAN.to_excel( writer , sheet_name='Titan')
        PARIS.to_excel( writer , sheet_name='Paris')
        UMBRELLA.to_excel( writer , sheet_name='Umbrella')
        writer.close()
        print('exported to: ', outfile)
