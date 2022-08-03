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
asv_folder = secret_charles + 'ASV 2022/'
output_folder = asv_folder + 'pics/'
intake = tracking + 'Sample Intake Log.xlsx'
research = pvi + 'Reports & Data/From Krammer Lab/Master Sheet.xlsx'
paris = clin_ops + 'PARIS/'
paris_tracker = paris + 'Patient Tracking - PARIS.xlsx'
paris_datasets = paris + 'Datasets/'
reports = pvi + 'Reports & Data/'
projects = reports + 'Projects/'


iris_folder = umbrella + 'IRIS/'
titan_folder = umbrella + 'TITAN/'
mars_folder = umbrella + 'MARS/'

script_folder = pvi + 'Scripts/'
script_input = script_folder + 'input/'
script_output = script_folder + 'output/'

d4 = proc + 'D4 Sheets/'
unfiltered = script_output + 'all_D4.xlsx'
filtered = script_output + 'SERONET_In_Window_Data.xlsx'
report_form = script_output + 'SERONET.csv'

header_intake = 6
