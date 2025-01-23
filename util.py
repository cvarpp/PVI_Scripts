import os

home = os.path.expanduser('~') + '/'

onedrive = os.environ.get('SHAREPOINT_DIR', home + 'The Mount Sinai Hospital/')
tracking = os.environ.get('TRACKING_DIR', onedrive + 'Simon Lab - Sample Tracking/')
pvi = os.environ.get('PVI_DIR', onedrive + 'Simon Lab - PVI - Personalized Virology Initiative/')
proc = os.environ.get('PROCESS_DIR', onedrive + 'Simon Lab - Processing Team/')
project_ws = os.environ.get('PROJECT_DIR', onedrive + 'Simon Lab - Project Workspace/')
psp = os.environ.get('PSP_DIR', onedrive + 'Simon Lab - Pathogen Surveillance/')
printing = os.environ.get('PRINT_DIR', onedrive + 'Simon Lab - Print Shop/')

dscf = proc + 'Data Sample Collection Form.xlsx'
proc_ntbk = proc + 'Processing Notebook.xlsx'
inventory_input = proc + 'New Import Sheet.xlsx'

sequencing = psp + 'Sequencing/'
extractions = sequencing + 'Extraction/'

shared_samples = tracking + 'Released Samples/Collaborator Samples Tracker.xlsx'
intake = tracking + 'Sample Intake Log Historian.xlsx'

clin_ops = pvi + 'Clinical Research Study Operations/'
umbrella = clin_ops + 'Umbrella Viral Sample Collection Protocol/'
secret_charles = pvi + 'Secret Sheets/Charles/'


reports = pvi + 'Reports & Data/'
research = reports + 'From Krammer Lab/Master Sheet.xlsx'
projects = reports + 'Projects/'

tube_print = printing + 'Tube Printing/'
print_log = tube_print + 'Printing Log.xlsx'

apollo = clin_ops + 'APOLLO/'
apollo_tracker = apollo + 'APOLLO Participant Tracker.xlsx'

paris = clin_ops + 'PARIS/'
paris_tracker = paris + 'Patient Tracking - PARIS.xlsx'
paris_datasets = paris + 'Datasets/'

iris_folder = umbrella + 'IRIS/'
titan_folder = umbrella + 'TITAN/'
mars_folder = umbrella + 'MARS/'
gaea_folder = umbrella + 'GAEA/'
prio_folder = clin_ops + 'PRIORITY/'
apollo_folder = clin_ops + 'APOLLO/' 
crp_folder = umbrella + 'Critical Reference Panel/'

umbrella_tracker = umbrella + "Patient Tracker - Umbrella Virology Protocol.xlsx"
titan_tracker = titan_folder + "TITAN Participant Tracker.xlsx"
cam_long = clin_ops + "Long-Form CAM Schedule.xlsx"

script_folder = pvi + 'Scripts/'
script_input = script_folder + 'input/'
script_output = script_folder + 'output/'
script_data = script_folder + 'data/'

proc_d4 = proc + 'D4 Sheets/'
unfiltered = script_output + 'all_D4.xlsx'
filtered = script_output + 'SERONET_In_Window_Data.xlsx'
report_form = script_output + 'SERONET.csv'

cross_project = clin_ops + 'Cross-Project/'
sample_query = cross_project + 'Sample ID Query/'
sharing = cross_project + 'Result Sharing/'
cross_d4 = cross_project + 'Seronet Task D4/'
seronet_data = cross_d4 + 'Data/'
seronet_qc = cross_d4 + 'QC/'
seronet_vax = cross_d4 + 'Vax/'

fp_url = 'https://mssm-simonlab.freezerpro.com/api/v2/'

# Header Heights

header_intake = 0

# Column Names

qual = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
quant = 'Ab Concentration (Units - AU/mL)'
visit_type = "Visit Type / Samples Needed"

# Environment Variables

fp_pass = 'FP_PASS'
fp_user = 'FP_USER'
