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
            [sg.Checkbox('MRN', disabled=True, visible=False, key='MRN'), sg.Checkbox('Tracker Info', enable_events=True, disabled=True, visible=False, key='tracker'), sg.Checkbox("Research", disabled=True, visible=False, key="research")],
            [sg.Checkbox("All", enable_events=True, visible=False, key='all_trackers' ), sg.Checkbox('Umbrella Info', disabled=True, visible=False, key='umbrella'), sg.Checkbox('Paris Info', disabled=True, visible=False, key='paris'),
                sg.Checkbox('CRP Info', disabled=True, visible=False, key='crp'), sg.Checkbox('MARS Info', disabled=True, visible=False, key='mars')], 
            [sg.Checkbox('TITAN Info', disabled=True, visible=False, key='titan'), sg.Checkbox('GAEA Info', disabled=True, visible=False, key='gaea'), sg.Checkbox('ROBIN Info', disabled=True, visible=False, key='robin'),\
                sg.Checkbox('APOLLO Info', disabled=True, visible=False, key='apollo'), sg.Checkbox('DOVE Info', disabled=True, visible=False, key='dove')],
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
                window['research'].update(disabled=False, visible=True)
            else:
                window['MRN'].update(disabled=True, visible=False, value = False)
                window['tracker'].update(disabled=True, visible=False, value = False)
                window['research'].update(disabled=True, visible=False, value = False)
                window['all_trackers'].update(visible=False)
                window['umbrella'].update(visible=False)
                window['paris'].update(visible=False)
                window['crp'].update(visible=False)
                window['mars'].update(visible=False)
                window['titan'].update(visible=False)
                window['gaea'].update(visible=False)
                window['robin'].update(visible=False)
                window['apollo'].update(visible=False)
                window['dove'].update(visible=False)

        
        if event == 'tracker':
            if window['tracker'].get() == True:
                window['all_trackers'].update(disabled=False, visible=True, value=True)
                window['umbrella'].update(visible=True, value=True)
                window['paris'].update(visible=True, value=True)
                window['crp'].update(visible=True, value=True)
                window['mars'].update(visible=True, value=True)
                window['titan'].update(visible=True, value=True)
                window['gaea'].update(visible=True, value=True)
                window['robin'].update(visible=True, value=True)
                window['apollo'].update(visible=True, value=True)
                window['dove'].update(visible=True, value=True)
            else:
                window['all_trackers'].update(disabled=True, visible=False, value=False)
                window['umbrella'].update(disabled=True, visible=False, value=False)
                window['paris'].update(disabled=True, visible=False, value=False)
                window['crp'].update(disabled=True, visible=False, value=False)
                window['mars'].update(disabled=True, visible=False, value=False)
                window['titan'].update(disabled=True, visible=False, value=False)
                window['gaea'].update(disabled=True, visible=False, value=False)
                window['robin'].update(disabled=True, visible=False, value=False)
                window['apollo'].update(disabled=True, visible=False, value=False)
                window['dove'].update(disabled=True, visible=False, value=False)

        if event == 'all_trackers':
            if window['all_trackers'].get() == True:
                window['umbrella'].update(disabled=True, value=True)
                window['paris'].update(disabled=True, value=True)
                window['crp'].update(disabled=True, value=True)
                window['mars'].update(disabled=True, value=True)
                window['titan'].update(disabled=True, value=True)
                window['gaea'].update(disabled=True, value=True)
                window['robin'].update(disabled=True, value=True)
                window['apollo'].update(disabled=True, value=True)
                window['dove'].update(disabled=True, value=True)
            else:
                window['umbrella'].update(disabled=False)
                window['paris'].update(disabled=False)
                window['crp'].update(disabled=False)
                window['mars'].update(disabled=False)
                window['titan'].update(disabled=False)
                window['gaea'].update(disabled=False)
                window['robin'].update(disabled=False)
                window['apollo'].update(disabled=False)
                window['dove'].update(disabled=False)


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

    if args.research == True:
        Research = helpers.query_research(sid_list=Samples)

    if args.tracker == True:
        tracker_list=[]
        tracker_names=[]

        if args.paris == True:
            PARIS = pd.read_excel(util.paris_tracker, sheet_name="Main", header=8).query("@partID in `Subject ID`")
            tracker_list.append(PARIS)
            tracker_names.append("PARIS")
        
        if args.titan == True:
            TITAN = pd.read_excel(util.titan_tracker, sheet_name="Tracker", header=4).query("@partID in `Umbrella Corresponding Participant ID`")     
            tracker_list.append(TITAN)
            tracker_names.append("TITAN")
        
        if args.mars == True:
            MARS =  pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name="Pt List", header=0).query("@partID in `Participant ID`")
            tracker_list.append(MARS)        
            tracker_names.append("MARS")

        if args.crp == True:
            CRP = pd.read_excel(util.crp_folder + "CRP Patient Tracker.xlsx", sheet_name="Tracker", header=4).query("@partID in `Participant ID`")
            tracker_list.append(CRP)        
            tracker_names.append("CRP")

        if args.apollo == True:
            APOLLO = pd.read_excel(util.apollo_folder + "APOLLO Participant Trakcer.xlsx", sheet_name="Summary", header=0).query("@partID in `Participant ID`")
            tracker_list.append(APOLLO)
            tracker_names.append("APOLLO")

        if args.gaea == True:
            GAEA = pd.read_excel(util.gaea_folder + "GAEA Tracker.xlsx", sheet_name="Summary", header=0).query("@partID in `Participant ID`")
            tracker_list.append(GAEA)
            tracker_names.append("GAEA")

        if args.Umbrella == True:
            UMBRELLA = pd.read_excel(util.umbrella_tracker, sheet_name="Summary").query("@partID in `Subject ID`")
            tracker_list.append(UMBRELLA)
            tracker_names.append("UMBRELLA")

        #ROBIN AND DOVE ARE WEIRD DEAL WITH LATER!

        if args.mrn == False:
            for item in tracker_list:
                try:
                    item.drop('MRN', axis=0)
                except:
                    continue
# %%
    if args.clinical== True:
        outfile = util.sample_query + args.outfilename + '.xlsx'
        with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
            Intake2.to_excel(writer , sheet_name='Intake Info')
            Processing.to_excel(writer , sheet_name='DSCF Info')
            for i, item in enumerate(tracker_list):
                item.to_excel(writer, sheet_name=tracker_names[i])

        print('exported to: ', outfile)
    
    if args.clinical== False:
        outfile = util.tracking + args.outfilename + '.xlsx'
        with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
            Intake2.to_excel( writer , sheet_name='Intake Info')
            Processing.to_excel( writer , sheet_name='DSCF Info')

        print('exported to: ', outfile)
