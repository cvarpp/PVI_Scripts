# PVI_Scripts

A set of useful scripts transforming data (input and output in Sharepoint folders synced with OneDrive).


#### GitHub Setup Guide
[GitHub Setup Guide](#create-a-github-account)


Scripts in subfolders (results and seronet, currently) must be run with module syntax, from the outer folder of this repository (until further developments). Specifically, the syntax for running the PARIS results report is:
```
python -m results.PARIS
```
with optional arguments after the script name.
Note that the module syntax does not include the .py extension.
Scripts in the outer folder can be run as previously, including aggregate_inventory:
```
python aggregate_inventory.py
```

For any of these scripts, you can use the `--help` argument to get a short description of the purpose of the function and any other arguments it accepts. The syntax for the PARIS reportings script, for example, would be:
```
python -m results.PARIS --help
```

## Result Reports

These scripts (in the `results` folder) bring together sample information (antibody test results and quantities of serum/plasma/cells/saliva) with participant information (demographics, vaccinations and infections, etc.).

- CRP
    - Includes T cell information. Pulls from the CRP tracker
- general
    - Legacy script. Pulls from participants_of_interest in PVI/Scripts/input. Study-agnostic, only pulls in antibody data
- MARS
    - Includes T cell information. Pulls from MARS for D4 Long
- PARIS
    - Focuses on COVID vaccinations and infections
- TITAN
    - Focuses on splitting out data by transplant type and HIV status. TITAN is uniquely suited to wide-form datasets, with a strong adherence to specific timepoints anchored to the third dose.


## SERONET Scripts

These scripts (in the `seronet` folder) are meant to assist with preparing data for quarterly data submissions and monthly accrual reporting. Some scripts directly pull together data for those various reports, and some are utilities.

### Central Pipeline:
- d4_all_data
    - Pulls together processing information into a single sheet. This script has the longest read time, as it is the only script in the pipeline that queries the main tabs of the intake and processing logs.
- ecrabs
    - A pronounceable acronym for equipment, consumables, reagent, aliquot, biospecimen, shipping manifest. These are the required sheets for the quarterly data submission that do not rely on clinical information (with the exception of an index date for each participant). The script outputs one workbook with a sheet for each of these tables, pulling primarily from the output of d4_all_data
- clinical_forms
    - Lacking a catchy acronym, this script outputs the baseline, follow-up, COVID vaccination, COVID infection, treatment, and any cohort-specific tables (transplant, auto-immune, or cancer). It pulls heavily from the "X for D4 Long" workbooks in each cohort's clinical information folder, reformatting information as needed to map information to a specific visit.
- accrual
    - This script produces the 3 workbooks associated with the monthly accrual report, corresponding to participant information, visit information, and vaccination information. Because it contains both clinical and processing data, it calls the main functions of all 3 preceding scripts before reorganizing and filtering the information as needed. It relies on having up-to-date information in each D4 Long sheet (at a minimum, participant IDs and research IDs must be present in the `Baseline Info` sheet and any cohort-specific sheets in their respective workbook).
### SERONET Utilities
- check_output
    - This will check the presence or absence of visits and biospecimens across each form being submitted, in order to identify any missing information. This is most likely to impact the biospecimen_test_result workbook, which is compiled by the pathology department given a list of order IDs rather than directly by the script.
- convert_to_csvs
    - This script can be run as a module (`python -m seronet.convert_to_csvs`) or from another folder (`python seronet/convert_to_csvs.py`) as it has no local dependencies. This can be helpful because it reads excel files from a fully specified input folder and converts their first sheets into csv files with the same name in a fully specified output folder, so the second construction can save you some typing. This is certainly possible to do manually, but the script can be faster with 2 or more workbooks to convert.
- filter_sheets
    - This may graduate to a pipeline script. The latest iteration of the pipeline has come to rely on the intermediate files created by the monthly accrual report. After a cleaned biospecimen file is constructed for a given data submission, a list of sample IDs can be finalized. This script takes as input a list of valid values (sample IDs), a filter column (`Sample ID`), and an initial workbook (either clinical or processing forms). It outputs a filtered copy of the workbook with the same sheets but only containing the rows with a valid value (`Sample ID`).
- filter_convert_to_csvs
    - This functionality should be folded into the main pipeline. It adjusts the visit ID and biospecimen IDs to reflect the additional out-of-window samples submitted in batch 4E.
- vaccination
    - Converts wide-form vaccination data in the different cohort trackers to the long-form data expected in D4 Long sheets.
- TITAN_meds
    - Converts the wide-form treatment data in the TITAN tracker to the long-form treatments expected in the D4 Long sheet.


## Other Scripts

Not to be underestimated.

- aggregate_inventory
    - Converts completed inventory from the tab-per-box format into the tab-per-sample-type format as well as reformatting for FreezerPro upload.
- cam_convert
    - Converts the wide-form tab-per-week CAM Clinic schedule into a single long-form sheet for easy filtering, searching, and joining
- freezerpro_delete
    - Legacy. An outdated use of the FreezerPro API. Included for reference and possible expansion of aggregate_inventory functionality
- freezerpro_pull
    - Legacy. An outdated use of the FreezerPro API. Included for reference and possible expansion of aggregate_inventory functionality
- helpers.py
    - A variety of helper functions, documented within the script. All querying of common source documents should be done through these functions to ensure clean data.
- par_umb_sharing
    - Compiles unshared results with known associated emails into an excel sheet suitable for mailmerge. After a mailmerge, the sample ID column can be used to update the correct tab in the sample intake log.
- print_log
    - Converts the per-print-session material information in the printing log and the per-bottle information in the processing notebook/dscf lot # sheets into per-sample+material sheet for easy joining for future use in the ecrabs preparation script.
-proc_missing_info
    - Tracks missing information from the processing notebook and associates it with the people best equipped to fill out the information
-util
    - Shared constants, largely related to file organization.