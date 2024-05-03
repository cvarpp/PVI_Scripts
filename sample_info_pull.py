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
import difflib as dl
from copy import deepcopy

#%%

def parse_input():
    x = True
    check = ''
    sg.theme('Dark Blue 17')

    layout_password = [[sg.Text("Are you Really Clincally Compliant?")],
                       [sg.Text("ID"),sg.Input(key='userid')],
                       [sg.Text("Password"), sg.Input(key='password', password_char='*')],
                       [sg.Submit(), sg.Cancel()]]

    layout_main = [[sg.Text('Sample ID Cohort Consolidator: SICC')],
            [sg.Radio('External File?', 'RADIO1', enable_events=True, key='Infile', default=True), sg.Radio('Text List?', 'RADIO1', enable_events=True, key='Text_list'), sg.Checkbox('Clinical Compliant?', enable_events=True, key='clinical')],
            [sg.Checkbox('MRN', disabled=True, visible=False, key='MRN'), sg.Checkbox('Tracker Info', enable_events=True, disabled=True, visible=False, key='tracker'), sg.Checkbox("Research", disabled=True, visible=False, key="research")],
            [sg.Checkbox("All", enable_events=True, visible=False, key='all_trackers' ), sg.Checkbox('Umbrella Info', disabled=True, visible=False, key='umbrella'), sg.Checkbox('Paris Info', disabled=True, visible=False, key='paris'),
                sg.Checkbox('CRP Info', disabled=True, visible=False, key='crp'), sg.Checkbox('MARS Info', disabled=True, visible=False, key='mars')], 
            [sg.Checkbox('TITAN Info', disabled=True, visible=False, key='titan'), sg.Checkbox('GAEA Info', disabled=True, visible=False, key='gaea'), sg.Checkbox('ROBIN Info', disabled=True, visible=False, key='robin'),\
                sg.Checkbox('APOLLO Info', disabled=True, visible=False, key='apollo'), sg.Checkbox('DOVE Info', disabled=True, visible=False, key='dove')],
            [sg.Text('File\nName', size=(9, 2), key='filepath_text'), sg.Input(key='filepath'), sg.FileBrowse(key='filepath_browse')],
            [sg.Text('Sheet\nName', size=(9, 2), key='sheetname_text'), sg.Input(key='sheetname')], [sg.Text("Sample ID", size=(9,2), key='sampleid_text'), sg.Input(default_text="Sample ID", key='sampleid')],
            [sg.Text('Sample\nList', size=(9,2), key='sample_list_text', visible=False), sg.Multiline(key="Sample_List", disabled=True, size=(40,10), visible=False)],
            [sg.Text('Outfile Name'), sg.Input(key='outfilename')],                  
            [sg.Submit(), sg.Cancel()]]

    study_keys = ['umbrella','paris','crp','mars','titan','gaea','robin','apollo','dove']
    window_main = sg.Window('Sample Query', layout_main)

    while x == True:
        event_main, values = window_main.read()
        if event_main == 'Text_list':
            window_main['filepath'].update(disabled=True, visible=False)
            window_main['filepath_text'].update(visible=False)
            window_main['filepath_browse'].update(disabled=True, visible=False)
            window_main['sheetname'].update(disabled=True, visible=False)
            window_main['sheetname_text'].update(visible=False)
            window_main['sampleid'].update(disabled=True, visible=False)
            window_main['sampleid_text'].update(visible=False)
            window_main['Sample_List'].update(disabled=False, visible=True)
            window_main['sample_list_text'].update(visible=True)
        elif event_main == 'Infile':
            window_main['filepath'].update(disabled=False, visible=True)
            window_main['filepath_text'].update(visible=True)
            window_main['filepath_browse'].update(disabled=False, visible=True)
            window_main['sheetname'].update(disabled=False, visible=True)
            window_main['sheetname_text'].update(visible=True)
            window_main['sampleid'].update(disabled=False, visible=True)
            window_main['sampleid_text'].update(visible=True)
            window_main['Sample_List'].update(disabled=True, visible=False)
            window_main['sample_list_text'].update(visible=False)
        elif event_main == 'clinical':
            if check =='Validated':
                if window_main['clinical'].get() == True:
                    window_main['MRN'].update(disabled=False, visible=True)
                    window_main['tracker'].update(disabled=False, visible=True)
                    window_main['research'].update(disabled=False, visible=True)
                else:
                    window_main['MRN'].update(disabled=True, visible=False, value = False)
                    window_main['tracker'].update(disabled=True, visible=False, value = False)
                    window_main['research'].update(disabled=True, visible=False, value = False)
                    window_main['all_trackers'].update(visible=False)
                    for study in study_keys:
                        window_main[study].update(visible=False)
            else:
                window_password = sg.Window('SICC Login', deepcopy(layout_password))
                event_password, values_password = window_password.read()
                if event_password == 'Cancel':
                    window_main['clinical'].update(value=False)
                    window_password.close()
                elif event_password == sg.WIN_CLOSED:
                    window_main['clinical'].update(value=False)
                    window_password.close() 
                elif event_password == 'Submit':
                    check = helpers.corned_beef(values_password["userid"],values_password['password'])
                    if check == "Validated":
                        if window_main['clinical'].get() == True:
                            window_main['MRN'].update(disabled=False, visible=True)
                            window_main['tracker'].update(disabled=False, visible=True)
                            window_main['research'].update(disabled=False, visible=True)
                        else:
                            window_main['MRN'].update(disabled=True, visible=False, value = False)
                            window_main['tracker'].update(disabled=True, visible=False, value = False)
                            window_main['research'].update(disabled=True, visible=False, value = False)
                            window_main['all_trackers'].update(visible=False)
                            for study in study_keys:
                                window_main[study].update(visible=False)
                        window_password.close()
                    else:
                        window_main['clinical'].update(value=False)
                        window_password.close()
        elif event_main == 'tracker':
            if window_main['tracker'].get() == True:
                window_main['all_trackers'].update(disabled=False, visible=True, value=True)
                for study in study_keys:
                    window_main[study].update(visible=True, value=True)
            else:
                window_main['all_trackers'].update(disabled=True, visible=False, value=False)
                for study in study_keys:
                    window_main[study].update(disabled=True, visible=False, value=False)
        elif event_main == 'all_trackers':
            if window_main['all_trackers'].get() == True:
                for study in study_keys:
                    window_main[study].update(disabled=True, value=True)
            else:
                for study in study_keys:
                    window_main[study].update(disabled=False)
        elif event_main in ['Cancel', sg.WIN_CLOSED]:
            quit()
        elif event_main == "Submit":
            args = helpers.ValuesToClass(values)
            window_main.close()
            x=False
    return args

if __name__ == '__main__':
    args = parse_input()

#%%
    if args.Text_list == True:
        Samples=re.split(r'\r?\n', args.Sample_List)
    
    if args.Infile == True:
        Sample_List = pd.read_excel(args.filepath, sheet_name=args.sheetname)
        Samples=Sample_List[args.sampleid].tolist()
    
    Intake = helpers.query_intake(include_research=True)
    Intake2 = Intake.query("@Samples in `sample_id`")
    
    partID = Intake2['participant_id'].to_list()
    
    Processing = helpers.query_dscf(sid_list=Samples, broad_rename=True)
    Processing2 = Processing[Processing.columns.drop(list(Processing.filter(regex='Unnamed')))]
    
    col_renames=[]
    Processing_Serum = Processing2.filter(regex=r"(?i)(Serum)")

    serum_vol = Processing_Serum['Total volume of serum (mL)'].append(Processing_Serum['Total volume of serum (ml)'].append(Processing_Serum['Volume of Serum Collected (mL)']))
    serum_init = Processing_Serum['Aliquoted by: SERUM'].append(Processing_Serum["Serum Aliquoted by"])
    serum_freeze = Processing_Serum['Serum Freeze Time'].append(Processing_Serum['serum_freeze_time'])

    for item in col_renames:
            item.dropna(inplace=True)
            item = item.reset_index()
            item.drop_duplicates(subset='sample_id', keep='last', inplace=True)




    Processing_Cells_Plasma = Processing2.filter(regex=r"(?i)(Plasma|Cell|PBMC)")
    
    Processing_Saliva = Processing2.filter(regex=r"(?i)(Saliva)")

    Sub_frame_list = [Processing_Serum, Processing_Saliva, Processing_Cells_Plasma]
    # for i , Frame in enumerate(Sub_frame_list):
    #     Frame.dropna(axis=1, how="all", inplace=True)
    #     for col in enumerate(Frame.columns):
    #         if dl.SequenceMatcher(None, col, "Volume of serum (mL)") >= 0.5
    #             print(Frame[col])

    if args.research == True:
        Research = helpers.query_research(sid_list=Samples)
    
    if args.tracker == True:
        tracker_list=[]
        tracker_names=[]
        if args.paris == True:
            paris_part1 = pd.read_excel(util.paris_tracker, sheet_name="Subgroups", header=4).query("@partID in `Participant ID`")
            paris_part2 = pd.read_excel(util.paris_tracker, sheet_name="Participant details", header=0).query("@partID in `Subject ID`")
            paris_part3 = pd.read_excel(util.paris_tracker, sheet_name="Flu Vaccine Information", header=0).query("@partID in `Participant ID`")

            paris_part2['Participant ID'] = paris_part2['Subject ID']
            paris_part2.drop('Subject ID', inplace=True, axis='columns')

            PARIS = pd.concat([paris_part1,paris_part2,paris_part3])
            
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
            APOLLO = pd.read_excel(util.apollo_folder + "APOLLO Participant Tracker.xlsx", sheet_name="Summary", header=0).query("@partID in `Participant ID`")
            tracker_list.append(APOLLO)
            tracker_names.append("APOLLO")
        if args.gaea == True:
            GAEA = pd.read_excel(util.gaea_folder + "GAEA Tracker.xlsx", sheet_name="Summary", header=0).query("@partID in `Participant ID`")
            tracker_list.append(GAEA)
            tracker_names.append("GAEA")
        if args.umbrella == True:
            UMBRELLA = pd.read_excel(util.umbrella_tracker, sheet_name="Summary").query("@partID in `Subject ID`")
            tracker_list.append(UMBRELLA)
            tracker_names.append("UMBRELLA")
        tracker_combined = pd.concat(tracker_list)

        #ROBIN AND DOVE ARE WEIRD DEAL WITH LATER!

# %%
    if args.clinical == True:
        outfile = util.sample_query + args.outfilename + '.xlsx'
        with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
            Intake2.to_excel(writer , sheet_name='Intake Info')
            Processing.to_excel(writer , sheet_name='DSCF Info')
            for i, df in enumerate(tracker_list):
                if df.shape[0] != 0:
                    df.to_excel(writer, sheet_name=tracker_names[i])
        print('exported to:', outfile)
    else:
        outfile = util.tracking + 'Sample ID Query/' + args.outfilename + '.xlsx'
        with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
            Intake2.to_excel( writer , sheet_name='Intake Info')
            Processing.to_excel( writer , sheet_name='DSCF Info')
        print('exported to:', outfile)

# %%
