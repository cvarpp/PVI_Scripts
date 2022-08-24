import pandas as pd
import numpy as np
import datetime
from datetime import date
from dateutil import parser
import util
import sys
import pickle

def pull_from_source():
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    visit_type = "Visit Type / Samples Needed"
    iris_data = pd.read_excel(util.iris_folder + 'Participant Tracking - IRIS.xlsx', sheet_name='Main Project', header=4).dropna(subset=['Participant ID'])
    iris_data['Participant ID'] = iris_data['Participant ID'].apply(lambda val: val.strip().upper())
    titan_data = pd.read_excel(util.titan_folder + 'TITAN Participant Tracker.xlsx', sheet_name='Tracker', header=4).dropna(subset=['Umbrella Corresponding Participant ID'])
    titan_data['Participant ID'] = titan_data['Umbrella Corresponding Participant ID'].apply(lambda val: val.strip().upper())
    visit_type = "Visit Type / Samples Needed"
    mars_data = pd.read_excel(util.mars_folder + 'MARS tracker.xlsx', sheet_name='Pt List').dropna(subset=['Participant ID'])
    mars_data['Participant ID'] = mars_data['Participant ID'].apply(lambda val: val.strip().upper())
    participants = [p for p in mars_data['Participant ID'].unique()] + [p for p in titan_data['Participant ID'].unique()] + [p for p in iris_data['Participant ID'].unique()]
    participant_study = {p: 'MARS' for p in mars_data['Participant ID'].unique()}
    participant_study.update({p: 'TITAN' for p in titan_data['Participant ID'].unique()})
    participant_study.update({p: 'IRIS' for p in iris_data['Participant ID'].unique()})
    mars_data.set_index('Participant ID', inplace=True)
    iris_data.set_index('Participant ID', inplace=True)
    titan_data.set_index('Participant ID', inplace=True)
    collection_log = pd.read_excel(util.intake, sheet_name='Collection Log')
    collection_log['idx'] = collection_log['Sample ID'].apply(lambda val: str(val).strip().upper())
    collection_log.set_index('idx', inplace=True)
    samples = pd.read_excel(util.intake, sheet_name='Sample Intake Log', header=util.header_intake, dtype=str)
    newCol = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
    newCol2 = 'Ab Concentration (Units - AU/mL)'
    samplesClean = samples.dropna(subset=['Participant ID'])
    participant_samples = {participant: [] for participant in participants}
    submitted_key = pd.read_excel(util.clin_ops + 'Cross-Project/Seronet Task D4/Data/SERONET Key.xlsx', sheet_name='Source').drop_duplicates(subset=['Participant ID']).set_index('Participant ID')
    for _, sample in samplesClean.iterrows():
        if len(str(sample['Sample ID'])) != 5:
            continue
        participant = str(sample['Participant ID']).strip().upper()
        if str(sample['Study']).strip().upper() == 'PRIORITY':
            if participant not in participant_samples.keys():
                participant_samples[participant] = []
                participant_study[participant] = 'PRIORITY'
        if participant in participant_samples.keys():
            if str(sample[newCol]).strip().upper() == "NEGATIVE":
                sample[newCol2] = "Negative"
            if pd.isna(sample[newCol2]):
                result_new = '-'
            elif type(sample[newCol2]) == int:
                result_new = sample[newCol2]
            elif str(sample[newCol2]).strip().upper() == "NEGATIVE":
                result_new = 1.
            else:
                result_new = sample[newCol2]
            try:
                sample_date = parser.parse(str(sample['Date Collected'])).date()
            except:
                print(sample['Sample ID'], "has invalid date", sample['Date Collected'])
                sample_date = parser.parse('1/1/1900').date()
            participant_samples[participant].append((sample_date, sample['Sample ID'], sample[visit_type], sample[newCol], result_new, sample['Blood Collector Initials']))
            
    research_samples_1 = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Reports & Data/From Krammer Lab/Master Sheet.xlsx', sheet_name='Inputs')
    research_samples_2 = pd.read_excel('~/The Mount Sinai Hospital/Simon Lab - PVI - Personalized Virology Initiative/Reports & Data/From Krammer Lab/Master Sheet.xlsx', sheet_name='Archive')
    research_results = {}
    for _, sample in research_samples_1.iterrows():
        sample_id = str(sample['Sample ID']).strip()
        spike = sample['Spike endpoint']
        if str(spike).strip().upper() == "NEG":
            auc = 1.
        else:
            auc = sample['AUC']
        if not pd.isna(auc):
            research_results[sample_id] = [spike, auc]
    for _, sample in research_samples_2.iterrows():
        sample_id = str(sample['Sample ID']).strip()
        spike = sample['Spike endpoint']
        if str(spike).strip().upper() == "NEG":
            auc = 1.
        else:
            auc = sample['AUC']
        if not pd.isna(auc) and sample_id not in research_results.keys():
            research_results[sample_id] = [spike, auc]
    bsl2p_samples = pd.read_excel(util.dscf, sheet_name='BSL2+ Samples', header=1)
    bsl2_samples = pd.read_excel(util.dscf, sheet_name='BSL2 Samples')
    shared_samples = pd.read_excel(util.shared_samples, sheet_name='Released Samples')
    no_pbmcs = set([str(sid).strip() for sid in shared_samples[shared_samples['Sample Type'] == 'PBMC']['Sample ID'].unique()])
    missing_info = {'Sample ID': [], 'Serum Volume': [], 'PBMC Conc': [], 'PBMC Vial Count': []}
    sample_info = {}
    for _, sample in bsl2p_samples.iterrows():
        sample_id = str(sample['Sample ID']).strip().upper()
        try:
            serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("mlML ").split()[0])
        except:
            try:
                serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("ulUL ").split()[0]) / 1000.
            except:
                print(sample_id, sample['Total volume of serum (ml)'], type(sample['Total volume of serum (ml)']), "is not valid")
                if type(sample['Total volume of serum (ml)']) == str:
                    serum_volume = 0
                else:
                    serum_volume = sample['Total volume of serum (ml)']
        if type(serum_volume) != str and serum_volume > 4.5:
            serum_volume = 4.5
        if sample_id in no_pbmcs:
            pbmc_conc = 0
            vial_count = 0
        else:
            pbmc_conc = sample['# cells per aliquot']
            vial_count = sample['# of aliquots frozen']
            if (type(pbmc_conc) != float or pd.isna(pbmc_conc)) and type(pbmc_conc) != int:
                pbmc_conc = 0
                vial_count = 0
        if type(vial_count) in [int, float] and vial_count > 2:
            vial_count = 2
        if sample_id in collection_log.index:
            coll_time = collection_log.loc[sample_id, 'Time Collected']
        else:
            coll_time = '?'


        rec_time = sample['Time Received']
        proc_time = '???'
        serum_freeze_time = sample['Time put in -80: SERUM']
        cell_freeze_time = sample['Time put in -80: PBMC']
        proc_inits = sample['Processed by (initials)']
        viability = sample['% Viability']
        cpt_vol = sample['CPT/EDTA VOL']
        sst_vol = sample['SST VOL']
        comment = sample['Comments']

        if not ((pd.isna(serum_volume) or serum_volume == 0) and (type(pbmc_conc) == str or pd.isna(pbmc_conc) or pbmc_conc == 0)):
            sample_info[sample_id] = [serum_volume, pbmc_conc, vial_count, coll_time, rec_time, proc_time, serum_freeze_time, cell_freeze_time, proc_inits, viability, cpt_vol, sst_vol, comment]
        else:
            missing_info['Sample ID'].append(sample_id)
            missing_info['Serum Volume'].append(serum_volume)
            missing_info['PBMC Conc'].append(pbmc_conc)
            missing_info['PBMC Vial Count'].append(vial_count)
    for _, sample in bsl2_samples.iterrows():
        sample_id = str(sample['Sample ID']).strip().upper()
        try:
            serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("mlML ").split()[0])
        except:
            try:
                serum_volume = float(str(sample['Total volume of serum (ml)']).strip().strip("ulUL ").split()[0]) / 1000.
            except:
                print(sample_id, sample['Total volume of serum (ml)'], type(sample['Total volume of serum (ml)']), "is invalid")
                if type(sample['Total volume of serum (ml)']) == str:
                    serum_volume = 0
                else:
                    serum_volume = sample['Total volume of serum (ml)']
        if type(serum_volume) != str and serum_volume > 4.5:
            serum_volume = 4.5
        if sample_id in no_pbmcs:
            pbmc_conc = 0
            vial_count = 0
        else:
            pbmc_conc = sample['# cell per aliquot']
            vial_count = sample['# of aliquots frozen']
            if (type(pbmc_conc) != float or pd.isna(pbmc_conc)) and type(pbmc_conc) != int:
                pbmc_conc = 0
                vial_count = 0
        if type(vial_count) in [int, float] and vial_count > 2:
            vial_count = 2
        if sample_id in collection_log.index:
            coll_time = collection_log.loc[sample_id, 'Time Collected']
        else:
            coll_time = sample['Time collected']
        rec_time = sample['Time Received']
        proc_time = sample['Time Started Processing']
        serum_freeze_time = sample['Time put in -80: SERUM']
        cell_freeze_time = sample['Time put in -80: PBMC']
        proc_inits = sample['Processed by (initials)'] # is this our best column?
        viability = sample['% Viability']
        cpt_vol = sample['CPT/EDTA VOL']
        sst_vol = sample['SST VOL']
        comment = sample['COMMENTS']
        if not ((pd.isna(serum_volume) or serum_volume == 0) and (type(pbmc_conc) == str or pd.isna(pbmc_conc) or pbmc_conc == 0)):
            sample_info[sample_id] = [serum_volume, pbmc_conc, vial_count, coll_time, rec_time, proc_time, serum_freeze_time, cell_freeze_time, proc_inits, viability, cpt_vol, sst_vol, comment]
        else:
            missing_info['Sample ID'].append(sample_id)
            missing_info['Serum Volume'].append(serum_volume)
            missing_info['PBMC Conc'].append(pbmc_conc)
            missing_info['PBMC Vial Count'].append(vial_count)
    columns = ['Cohort', 'Seronet ID', 'Vaccine', '1st Dose Date', 'Days to 1st', '2nd Dose Date', 'Days to 2nd', 'Boost Vaccine', 'Boost Date', 'Days to Boost', 'Participant ID', 'Date', 'Post-Baseline', 'Sample ID', 'Visit Type', 'Qualitative', 'Quantitative', 'Spike endpoint', 'AUC', 'Log2AUC', 'Volume of Serum Collected (mL)', 'PBMC concentration per mL (x10^6)', '# of PBMC vials', 'coll_inits', 'coll_time', 'rec_time', 'proc_time', 'serum_freeze_time', 'cell_freeze_time', 'proc_inits', 'viability', 'cpt_vol', 'sst_vol', 'proc_comment']
    data = {col: [] for col in columns}
    participant_data = {}
    for participant, samples in participant_samples.items():
        try:
            samples.sort(key=lambda x: x[0])
        except:
            print(participant, "has samples that won't sort (see below)")
            print(samples)
            print()
            continue
        if len(samples) < 1:
            print(participant, "has no samples")
            print()
            continue
        baseline = samples[0][0]
        participant_data[participant] = {}
        if participant_study[participant] == 'TITAN':
            seronet_id = "14_T{}".format(samples[0][1])
            participant_data[participant]['Vaccine'] = titan_data.loc[participant, 'Vaccine Type']
            try:
                participant_data[participant]['1st Dose Date'] = titan_data.loc[participant, 'Vaccine #1 Date'].date()
            except:
                participant_data[participant]['1st Dose Date'] = titan_data.loc[participant, 'Vaccine #1 Date']
            try:
                participant_data[participant]['2nd Dose Date'] = titan_data.loc[participant, 'Vaccine #2 Date'].date()
            except:
                participant_data[participant]['2nd Dose Date'] = titan_data.loc[participant, 'Vaccine #2 Date']
            try:
                participant_data[participant]['Boost Date'] = titan_data.loc[participant, '3rd Dose Vaccine Date'].date()
            except:
                participant_data[participant]['Boost Date'] = titan_data.loc[participant, '3rd Dose Vaccine Date']
            participant_data[participant]['Boost Vaccine'] = titan_data.loc[participant, '3rd Dose Vaccine Type']
        elif participant_study[participant] == 'IRIS':
            seronet_id = "14_I{}".format(samples[0][1])
            participant_data[participant]['Vaccine'] = iris_data.loc[participant, 'Which Vaccine?']
            try:
                participant_data[participant]['1st Dose Date'] = iris_data.loc[participant, 'First Dose Date'].date()
            except:
                participant_data[participant]['1st Dose Date'] = iris_data.loc[participant, 'First Dose Date']
            try:
                participant_data[participant]['2nd Dose Date'] = iris_data.loc[participant, 'Second Dose Date'].date()
            except:
                participant_data[participant]['2nd Dose Date'] = iris_data.loc[participant, 'Second Dose Date']
            try:
                participant_data[participant]['Boost Date'] = iris_data.loc[participant, 'Third Dose Date'].date()
            except:
                participant_data[participant]['Boost Date'] = iris_data.loc[participant, 'Third Dose Date']
            participant_data[participant]['Boost Vaccine'] = iris_data.loc[participant, 'Third Dose Type']
        elif participant_study[participant] == 'MARS':
            seronet_id = "14_M{}".format(samples[0][1])
            participant_data[participant]['Vaccine'] = mars_data.loc[participant, 'Vaccine Name']
            try:
                participant_data[participant]['1st Dose Date'] = mars_data.loc[participant, 'Vaccine #1 Date'].date()
            except:
                participant_data[participant]['1st Dose Date'] = mars_data.loc[participant, 'Vaccine #1 Date']
            try:
                participant_data[participant]['2nd Dose Date'] = mars_data.loc[participant, 'Vaccine #2 Date'].date()
            except:
                participant_data[participant]['2nd Dose Date'] = mars_data.loc[participant, 'Vaccine #2 Date']
            try:
                participant_data[participant]['Boost Date'] = mars_data.loc[participant, '3rd Vaccine'].date()
            except:
                participant_data[participant]['Boost Date'] = mars_data.loc[participant, '3rd Vaccine']
            participant_data[participant]['Boost Vaccine'] = mars_data.loc[participant, '3rd Vaccine Type ']
        elif participant_study[participant] == 'PRIORITY':
            seronet_id = "14_P{}".format(samples[0][1])
            participant_data[participant]['Vaccine'] = ''
            participant_data[participant]['1st Dose Date'] = ''
            participant_data[participant]['2nd Dose Date'] = ''
            participant_data[participant]['Boost Date'] = ''
            participant_data[participant]['Boost Vaccine'] = ''
        else:
            print(participant, "is not in the study! They are in", participant_study[participant])
            exit(1)
        if participant in submitted_key.index:
            seronet_id = submitted_key.loc[participant, 'Research_Participant_ID']
        for date_, sample_id, visit_type, result, result_new, coll_inits in samples:
            sample_id = str(sample_id).strip()
            data['Cohort'].append(participant_study[participant])
            data['Vaccine'].append(participant_data[participant]['Vaccine'])
            data['1st Dose Date'].append(participant_data[participant]['1st Dose Date'])
            try:
                data['Days to 1st'].append(int((date_ - data['1st Dose Date'][-1]).days))
            except:
                data['Days to 1st'].append('')
            data['2nd Dose Date'].append(participant_data[participant]['2nd Dose Date'])
            try:
                data['Days to 2nd'].append(int((date_ - data['2nd Dose Date'][-1]).days))
            except:
                data['Days to 2nd'].append('')
            data['Boost Date'].append(participant_data[participant]['Boost Date'])
            try:
                data['Days to Boost'].append(int((date_ - data['Boost Date'][-1]).days))
            except:
                data['Days to Boost'].append('')
            data['Boost Vaccine'].append(participant_data[participant]['Boost Vaccine'])
            data['Seronet ID'].append(seronet_id)
            data['Participant ID'].append(participant)
            data['Date'].append(date_)
            data['Post-Baseline'].append((date_ - baseline).days)
            data['Sample ID'].append(sample_id)
            data['Visit Type'].append(visit_type)
            data['Qualitative'].append(result)
            data['Quantitative'].append(result_new)
            if sample_id in research_results.keys():
                res = research_results[sample_id]
            else:
                res = ['-', '-']
            if res[1] == '-':
                data['Log2AUC'].append('-')
            elif float(res[1]) == 0.:
                data['Log2AUC'].append(0.)
            else:
                data['Log2AUC'].append(np.log2(res[1]))
            data['Spike endpoint'].append(res[0])
            data['AUC'].append(res[1])
            # TODO: Change sample_info to a dataframe and iterate through columns instead
            unknown = ['?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', 'Processing Data Unavailable']
            if sample_id in sample_info.keys():
                serum_volume, pbmc_conc, vial_count, coll_time, rec_time, proc_time, serum_freeze_time, cell_freeze_time, proc_inits, viability, cpt_vol, sst_vol, comment = sample_info[sample_id]
            else:
                serum_volume, pbmc_conc, vial_count, coll_time, rec_time, proc_time, serum_freeze_time, cell_freeze_time, proc_inits, viability, cpt_vol, sst_vol, comment = unknown
            if participant_study[participant] == 'PRIORITY':
                pbmc_conc = 0 # May have to change to "N/A"
                vial_count = 0 # May have to change to "N/A"
            data['Volume of Serum Collected (mL)'].append(serum_volume)
            data['PBMC concentration per mL (x10^6)'].append(pbmc_conc)
            data['# of PBMC vials'].append(vial_count)
            data['coll_inits'].append(coll_inits)
            data['coll_time'].append(coll_time)
            data['rec_time'].append(rec_time)
            data['proc_time'].append(proc_time)
            data['serum_freeze_time'].append(serum_freeze_time)
            data['cell_freeze_time'].append(cell_freeze_time)
            data['proc_inits'].append(proc_inits)
            data['viability'].append(viability)
            data['cpt_vol'].append(cpt_vol)
            data['sst_vol'].append(sst_vol)
            data['proc_comment'].append(comment)
    report = pd.DataFrame(data)
    report.to_excel(util.unfiltered, index=False)
    return report

if __name__ == '__main__':
    pull_from_source()