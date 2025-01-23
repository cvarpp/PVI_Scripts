import os

home = os.path.expanduser('~') + os.sep

onedrive = home + os.environ.get('SHAREPOINT_DIR', 'The Mount Sinai Hospital') + os.sep
tracking = onedrive + os.environ.get('TRACKING_DIR', 'Simon Lab - Sample Tracking') + os.sep
pvi = onedrive + os.environ.get('PVI_DIR', 'Simon Lab - PVI - Personalized Virology Initiative') + os.sep
proc = onedrive + os.environ.get('PROCESS_DIR', 'Simon Lab - Processing Team') + os.sep
project_ws = onedrive + os.environ.get('PROJECT_DIR', 'Simon Lab - Project Workspace') + os.sep
psp = onedrive + os.environ.get('PSP_DIR', 'Simon Lab - Pathogen Surveillance') + os.sep
printing = onedrive + os.environ.get('PRINT_DIR', 'Simon Lab - Print Shop') + os.sep

dscf = proc + 'Data Sample Collection Form.xlsx'
proc_ntbk = proc + 'Processing Notebook.xlsx'
inventory_input = proc + 'New Import Sheet.xlsx'

sequencing = psp + 'Sequencing' + os.sep
extractions = sequencing + 'Extraction' + os.sep

shared_samples = tracking + 'Released Samples/Collaborator Samples Tracker.xlsx'
intake = tracking + 'Sample Intake Log Historian.xlsx'

clin_ops = pvi + 'Clinical Research Study Operations' + os.sep
umbrella = clin_ops + 'Umbrella Viral Sample Collection Protocol' + os.sep
secret_charles = pvi + 'Secret Sheets/Charles' + os.sep


reports = pvi + 'Reports & Data' + os.sep
research = reports + 'From Krammer Lab/Master Sheet.xlsx'
projects = reports + 'Projects' + os.sep

tube_print = printing + 'Tube Printing' + os.sep
print_log = tube_print + 'Printing Log.xlsx'

apollo = clin_ops + 'APOLLO' + os.sep
apollo_tracker = apollo + 'APOLLO Participant Tracker.xlsx'

paris = clin_ops + 'PARIS' + os.sep
paris_tracker = paris + 'Patient Tracking - PARIS.xlsx'
paris_datasets = paris + 'Datasets' + os.sep

iris_folder = umbrella + 'IRIS' + os.sep
titan_folder = umbrella + 'TITAN' + os.sep
mars_folder = umbrella + 'MARS' + os.sep
gaea_folder = umbrella + 'GAEA' + os.sep
prio_folder = clin_ops + 'PRIORITY' + os.sep
apollo_folder = clin_ops + 'APOLLO' + os.sep 
crp_folder = umbrella + 'Critical Reference Panel' + os.sep

umbrella_tracker = umbrella + "Patient Tracker - Umbrella Virology Protocol.xlsx"
titan_tracker = titan_folder + "TITAN Participant Tracker.xlsx"
cam_long = clin_ops + "Long-Form CAM Schedule.xlsx"

script_folder = pvi + 'Scripts' + os.sep
script_input = script_folder + 'input' + os.sep
script_output = script_folder + 'output' + os.sep
script_data = script_folder + 'data' + os.sep

proc_d4 = proc + 'D4 Sheets' + os.sep
unfiltered = script_output + 'all_D4.xlsx'
filtered = script_output + 'SERONET_In_Window_Data.xlsx'
report_form = script_output + 'SERONET.csv'

cross_project = clin_ops + 'Cross-Project' + os.sep
sample_query = cross_project + 'Sample ID Query' + os.sep
sharing = cross_project + 'Result Sharing' + os.sep
cross_d4 = cross_project + 'Seronet Task D4' + os.sep
seronet_data = cross_d4 + 'Data' + os.sep
seronet_qc = cross_d4 + 'QC' + os.sep
seronet_vax = cross_d4 + 'Vax' + os.sep

fp_url = 'https://mssm-simonlab.freezerpro.com/api/v2' + os.sep

# Header Heights

header_intake = 0

# Column Names

qual = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
quant = 'Ab Concentration (Units - AU/mL)'
visit_type = "Visit Type / Samples Needed"

# Environment Variables

fp_pass = 'FP_PASS'
fp_user = 'FP_USER'
