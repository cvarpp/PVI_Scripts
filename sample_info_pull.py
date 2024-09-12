import pandas as pd
import re
import argparse
import numpy as np
import util
import helpers
import PySimpleGUI as sg
from copy import deepcopy
from datetime import datetime
import os
from sys import exit


def flip_inputs(to_text, window_main):
    window_main['filepath'].update(disabled=to_text, visible=not to_text)
    window_main['filepath_text'].update(visible=not to_text)
    window_main['filepath_browse'].update(disabled=to_text, visible=not to_text)
    window_main['sheetname'].update(disabled=to_text, visible=not to_text)
    window_main['sheetname_text'].update(visible=not to_text)
    window_main['ID_column'].update(disabled=to_text, visible=not to_text)
    window_main['ID_column_text'].update(visible=not to_text)
    window_main['ID_list'].update(disabled=not to_text, visible=to_text)
    window_main['ID_list_text'].update(visible=to_text)

def clinical_view(window_main, study_keys):
    if window_main['clinical'].get() == True:
        window_main['MRN'].update(disabled=False, visible=True)
        window_main['tracker'].update(disabled=False, visible=True)
        window_main['contact'].update(disabled=False, visible=True)
    else:
        window_main['MRN'].update(disabled=True, visible=False, value = False)
        window_main['tracker'].update(disabled=True, visible=False, value = False)
        window_main['contact'].update(disabled=True, visible=False, value = False)
        window_main['all_trackers'].update(visible=False)
        for study in study_keys:
            window_main[study].update(visible=False)

def parse_input():
    check = None
    sg.theme('Dark Teal 12')

    layout_main = [[sg.Text('ID Cohort Consolidator: ICC')],
            [sg.Radio('Samples', 'RADIO5', enable_events=True, key='samples', default=True), sg.Radio('Participants', 'RADIO5', enable_events=True, key='participants')],
            [sg.Radio('External File', 'RADIO1', enable_events=True, key='Infile', default=True), sg.Radio('Text List', 'RADIO1', enable_events=True, key='Text_list'), sg.Checkbox('Clinical Compliant', enable_events=True, key='clinical')],
            [sg.Checkbox('MRN', disabled=True, visible=False, key='MRN'), sg.Checkbox('Tracker Info', enable_events=True, disabled=True, visible=False, key='tracker'), sg.Checkbox("Contact info", disabled=True, visible=False, key="contact")],
            [sg.Checkbox("All", enable_events=True, visible=False, key='all_trackers' ), sg.Checkbox('Umbrella Info', disabled=True, visible=False, key='umbrella'), sg.Checkbox('Paris Info', disabled=True, visible=False, key='paris'),
                sg.Checkbox('CRP Info', disabled=True, visible=False, key='crp'), sg.Checkbox('MARS Info', disabled=True, visible=False, key='mars')], 
            
            [sg.Checkbox('TITAN Info', disabled=True, visible=False, key='titan'), sg.Checkbox('GAEA Info', disabled=True, visible=False, key='gaea'), sg.Checkbox('ROBIN Info', disabled=True, visible=False, key='robin'),\
                sg.Checkbox('APOLLO Info', disabled=True, visible=False, key='apollo'), sg.Checkbox('DOVE Info', disabled=True, visible=False, key='dove')],
            
            [sg.Text('File\nName', size=(9, 2), key='filepath_text'), sg.Input(key='filepath'), sg.FileBrowse(key='filepath_browse')],
            [sg.Text('Sheet\nName', size=(9, 2), key='sheetname_text'), sg.Input(key='sheetname', default_text='Sheet1')], [sg.Text("ID Column", size=(9,2), key='ID_column_text'), sg.Input(default_text="Sample ID", key='ID_column')],
            [sg.Text('ID\nList', size=(9,2), key='ID_list_text', visible=False), sg.Multiline(key="ID_list", disabled=True, size=(40,10), visible=False)],
            [sg.Text('Outfile Name'), sg.Input(key='outfilename')],                  
            [sg.Button("Submit"), sg.Button("Cancel"), sg.Checkbox('Test', disabled=False, visible=True, key='test'), sg.Checkbox('Debug', disabled=False, visible=True, key='debug')]]

    study_keys = ['umbrella','paris','crp','mars','titan','gaea','robin','apollo','dove']
    window_main = sg.Window('ID Info Pull', layout_main)

    while True:
        event_main, values = window_main.read()
                
        if event_main == sg.WIN_CLOSED:
            break
        elif event_main == "Cancel":
            window_main.close()
        elif event_main == "Submit":
            args = helpers.ValuesToClass(values)
            window_main.close()
        elif event_main == 'participants':
            window_main['ID_column'].update(value="Participant ID")
        elif event_main == 'samples':
            window_main['ID_column'].update(value="Sample ID")
        elif event_main == 'Text_list':
            flip_inputs(True, window_main)
        elif event_main == 'Infile':
            flip_inputs(False, window_main)
        elif event_main == 'clinical':
            if check == 'Validated':
                clinical_view(window_main, study_keys)
            else:
                if os.path.exists(util.paris_tracker):
                    password_text = sg.popup_get_text(message='Clinical Access Password:', title="", password_char="*")
                    if helpers.corned_beef(password_text):
                        clinical_view(window_main, study_keys)
                    else:
                        window_main['clinical'].update(value=False)
                        sg.popup_ok("Incorrect username or password. Please try again.", title="", font=14)
                else:
                    window_main['clinical'].update(disabled=True, value=False)
                    sg.popup_ok("You (or your computer) cannot access clinical files. Processing information can still be pulled", title="", font=14)
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
    if args.test:
        args.filepath = util.script_data + 'Sample ID Query Test data set.xlsx'
        args.sheetname = "Sample Query Check"
        date_today = datetime.today()
        args.outfilename = f"Test Data {date_today:%Y-%m-%d %H%M}"
        args.Infile = True
        args.Text_list = False
    assert args.Text_list ^ args.Infile, "Exactly one of input file or text list should be selected"
    assert args.samples ^ args.participants, "Exactly one of sample IDs or participant IDs should be supplied"
    if args.Text_list:
        ID_values=re.split(r'\r?\n', args.ID_list)
    elif args.Infile:
        ID_list = pd.read_excel(args.filepath, sheet_name=args.sheetname)
        ID_values = ID_list[args.ID_column].tolist()
    if args.samples:
        args.samples = ID_values
        args.intake = helpers.query_intake(include_research=True).query("sample_id in @args.samples")
        args.pids = args.intake['participant_id'].to_list()
    elif args.participants:
        args.pids = ID_values
        args.intake = helpers.query_intake(include_research=True).query("participant_id in @args.pids")
        args.samples = args.intake.reset_index()['sample_id'].to_list()
    return args

def pull_trackers(args, trackers):
        # TODO: Add Robin and Dove clinical info
        # TODO: Complete SHIELD info
        pids = args.pids
        if args.paris == True:
            paris_tracker_excel = pd.ExcelFile(util.paris_tracker)
            paris_part1 = pd.read_excel(paris_tracker_excel, sheet_name="Subgroups", header=4).set_index('Participant ID')
            paris_part2 = pd.read_excel(paris_tracker_excel, sheet_name="Participant details", header=0).rename(columns={'Subject ID': 'Participant ID'}).set_index('Participant ID')
            paris_part3 = pd.read_excel(paris_tracker_excel, sheet_name="Flu Vaccine Information", header=0).set_index('Participant ID')
            paris_cols = set(paris_part1.columns) | set(paris_part2.columns) | set(paris_part3.columns)
            paris_cols = [col for col in paris_cols if re.search(r"(?i)Unnamed", str(col)) is not None]
            paris_ids = list(set(pids) & set(paris_part1.index))
            if len(paris_ids) > 0:
                trackers["paris"] = paris_part1.combine_first(paris_part2
                                              ).combine_first(paris_part3).loc[paris_ids, paris_cols].copy()
        if args.titan == True:
            trackers["titan"] = pd.read_excel(util.titan_tracker, sheet_name="Tracker", header=4
                                 ).query("`Umbrella Corresponding Participant ID` in @pids"
                                 ).rename(columns={"Umbrella Corresponding Participant ID": 'Participant ID'}
                                 ).set_index('Participant ID')
        if args.mars == True:
            trackers["mars"] =  pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name="Pt List", header=0
                                 ).query("`Participant ID` in @pids").set_index('Participant ID')
        if args.crp == True:
            trackers["crp"] = pd.read_excel(util.crp_folder + "CRP Patient Tracker.xlsx", sheet_name="Tracker", header=4
                               ).query("`Participant ID` in @pids").set_index('Participant ID')
        if args.apollo == True:
            trackers["apollo"] = pd.read_excel(util.apollo_folder + "APOLLO Participant Tracker.xlsx",
                                   sheet_name="Summary", header=0
                                   ).query("`Participant ID` in @pids").set_index('Participant ID')
        if args.gaea == True:
            trackers["gaea"] = pd.read_excel(util.gaea_folder + "GAEA Tracker.xlsx", sheet_name="Summary", header=0
                                ).query("`Participant ID` in @pids").set_index('Participant ID')
        if args.umbrella == True:
            trackers["umbrella"] = pd.read_excel(util.umbrella_tracker, sheet_name="Summary"
                                    ).query("`Subject ID` in @pids").rename(columns={'Subject ID': 'Participant ID'}
                                    ).set_index('Participant ID')
        for tracker_df in trackers.values():
            cols = tracker_df.columns.to_numpy()
            drop_cols = [col for col in tracker_df.columns if str(col).strip() != col and str(col).strip() in cols]
            tracker_df.drop(drop_cols, axis='columns', inplace=True)
            tracker_df.columns = [str(col).strip() for col in tracker_df.columns]

if __name__ == '__main__':
    try:
        args = parse_input()
    except:
        print("Could not parse input. Unrecoverable error, exiting...")
        exit(1)
    
    processing = helpers.query_dscf(sid_list=args.samples)

    col_mapper = pd.ExcelFile(util.script_data + 'Column Names.xlsx')
    dscf_cols = pd.read_excel(col_mapper,sheet_name="DSCF Column Names", header=0)
    tracker_cols = pd.read_excel(col_mapper,sheet_name="Tracker Columns", header=0)
    contact_cols = pd.read_excel(col_mapper,sheet_name="Contact Columns", header=0)

    dscf_cols_of_interest = dscf_cols['Cleaned Column Names'].unique().tolist()
    dscf_short_cols = dscf_cols[dscf_cols['short or debug'] == 'short']['Cleaned Column Names'].unique().tolist()

    for col in dscf_cols_of_interest:
        source_cols = dscf_cols[dscf_cols['Cleaned Column Names'] == col]['Source Column Names']
        if col not in processing.columns:
            processing[col] = np.nan
        for source_col in source_cols:
            processing[col] = processing[col].fillna(processing[source_col])
    
    processing['compiled comments'] = processing.filter(regex=r"(?i)Comments").astype(str).stack().groupby(level=0).agg(" - ".join)
    processing['compiled initials'] = processing.filter(regex=r"(?i)proc_inits").astype(str).stack().groupby(level=0).agg(" - ".join)

    dscf_cols_of_interest.extend(['compiled comments', 'compiled initials'])
    dscf_short_cols.append('compiled comments')

    processing_debug = processing.loc[:, dscf_cols_of_interest].copy()
    processing_short = processing.loc[:, dscf_short_cols].copy()

    trackers={}
    tracker_cleaned_combined = pd.DataFrame()
    if args.tracker == True:
        pull_trackers(args, trackers)
        if not args.contact:
            for tracker_name, tracker_df in trackers.items():
                drop_cols = [col for col in contact_cols['Column Names'] if col in tracker_df.columns]
                if len(drop_cols) > 0:
                    trackers[tracker_name] = tracker_df.drop(drop_cols, axis=1)
        if not args.MRN:
            for tracker_name, tracker_df in trackers.items():
                mrn_cols = [col for col in tracker_df.columns if re.search(r'(?i)(MRN)', str(col)) is not None]
                if len(mrn_cols) > 0:
                    trackers[tracker_name] = tracker_df.drop(mrn_cols, axis=1)

        tracker_keep_cols = tracker_cols[tracker_cols['Keep Drop Unique'] == 'keep']['Cleaned Column Names'].unique()
        tracker_unique_cols = tracker_cols[tracker_cols['Keep Drop Unique'] == 'unique']['Cleaned Column Names'].unique()
        keep_or_unique = set(tracker_keep_cols) | set(tracker_unique_cols)
        tracker_cleaned_short=[]
        for tracker_name, tracker_df in trackers.items():
            tracker_df['Source'] = tracker_name
            cols_to_clean = []
            for col in keep_or_unique:
                source_cols = tracker_cols[tracker_cols['Cleaned Column Names'] == col]['Source Column Names']
                if len(set(source_cols) & set(tracker_df.columns)) > 0:
                    cols_to_clean.append(col)
            cols_to_add = [col for col in cols_to_clean if col not in tracker_df.columns]
            print(tracker_name)
            tracker_df.loc[:, cols_to_add] = np.nan
            for col in cols_to_clean:
                source_cols = tracker_cols[tracker_cols['Cleaned Column Names'] == col]['Source Column Names']
                for source_col in source_cols:
                    if source_col in tracker_df.columns:
                        tracker_df[col] = tracker_df[col].fillna(tracker_df[source_col])
            intersect_full = list(keep_or_unique & set(tracker_df.columns))
            intersect_short = list(set(tracker_keep_cols) & set(tracker_df.columns))
            trackers[tracker_name] = tracker_df.loc[:, intersect_full].copy()
            tracker_cleaned_short.append(tracker_df.loc[:, intersect_short].copy())
        tracker_cleaned_combined = pd.concat(tracker_cleaned_short).dropna(axis="columns", how="all")

    if args.clinical == True:
        output_folder = util.sample_query
    else:
        output_folder = util.tracking + 'Sample ID Query/'
    if args.test == True:
        output_folder += 'Test Data/'
    outfile = output_folder + args.outfilename + '.xlsx'
    with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
        if args.debug==True:
            processing_debug.to_excel(writer, sheet_name='Processing Debug')
        processing_short.to_excel(writer, sheet_name='Processing Short-Form')
        args.intake.to_excel(writer, sheet_name='Intake Info')
        processing.to_excel(writer, sheet_name='DSCF Info')
        if len(trackers) > 0:
            tracker_cleaned_combined.to_excel(writer, sheet_name='Tracker compiled')
        for tracker_name, tracker_df in trackers.items():
            if tracker_df.shape[0] != 0:
                tracker_df.to_excel(writer, sheet_name=tracker_name)
    print('exported to:', outfile)
