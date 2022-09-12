import os

home = os.path.expanduser('~') + '/'
onedrive = home + 'The Mount Sinai Hospital/'
proc = onedrive + 'Simon Lab - Processing Team/'
dscf = proc + 'Data Sample Collection Form.xlsx'
pvi = onedrive + 'Simon Lab - PVI - Personalized Virology Initiative/'
tracking = onedrive + 'Simon Lab - Sample Tracking/'
shared_samples = tracking + 'Released Samples/Collaborator Samples Tracker.xlsx'
clin_ops = pvi + 'Clinical Research Study Operations/'
umbrella = clin_ops + 'Umbrella Viral Sample Collection Protocol/'
secret_charles = pvi + 'Secret Sheets/Charles/'
intake = tracking + 'Sample Intake Log.xlsx'

reports = pvi + 'Reports & Data/'
research = reports + 'From Krammer Lab/Master Sheet.xlsx'
projects = reports + 'Projects/'

paris = clin_ops + 'PARIS/'
paris_tracker = paris + 'Patient Tracking - PARIS.xlsx'
paris_datasets = paris + 'Datasets/'

iris_folder = umbrella + 'IRIS/'
titan_folder = umbrella + 'TITAN/'
mars_folder = umbrella + 'MARS/'

umbrella_tracker = umbrella + "Patient Tracker - Umbrella Virology Protocol.xlsx"

script_folder = pvi + 'Scripts/'
script_input = script_folder + 'input/'
script_output = script_folder + 'output/'

proc_d4 = proc + 'D4 Sheets/'
unfiltered = script_output + 'all_D4.xlsx'
filtered = script_output + 'SERONET_In_Window_Data.xlsx'
report_form = script_output + 'SERONET.csv'

cross_project = clin_ops + 'Cross-Project/'
cross_d4 = cross_project + 'Seronet Task D4/'
seronet_data = cross_d4 + 'Data/'
seronet_qc = cross_d4 + 'QC/'

# Header Heights

header_intake = 6

# Column Names


qual = 'Ab Detection S/P Result (Clinical) (Titer or Neg)'
quant = 'Ab Concentration (Units - AU/mL)'
visit_type = "Visit Type / Samples Needed"
