import pandas as pd
import datetime
import os
import argparse
import util
import win32com.client as win32  # For 'outlook' object created & emails sent via Outlook


# Missing Information from Processing Notebook:

## Pull filtered list from specimen dashboard into separate output seet
## Notify individuals based on initials
## Automate script to run every morning


# Set path
if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Make Seronet monthly sample report.')
    argParser.add_argument('-m', '--min_count', action='store', type=int, default=78)
    args = argParser.parse_args()
    
    home = os.path.expanduser('~') + '/'
    onedrive = home + 'The Mount Sinai Hospital/'
    proc = onedrive + 'Simon Lab - Processing Team/'
    proc_ntbk = proc + 'Processing Notebook.xlsx'
    df = pd.read_excel(proc_ntbk, sheet_name='Specimen Dashboard')


    # Cannot figure out what's wrong with missing_df so check dataframe here, delete later
    print(df.columns)

    
    # Set cutoff date for processing
    cutoff_date = datetime.datetime.now() - datetime.timedelta(days=3)    # better set days=3 for Mon-lastFri = 3d

    # Filter for samples: 3 days ago + missing data (sample_complete=No)
    missing_df = df[(df['Date Processed'] <= cutoff_date) & (df['Sample Complete?'] == 'No')]

    # Export filtered list
    output_path = proc + 'Processing Missing Information.xlsx'
    missing_df.to_excel(output_path, index=False)
    
    # # Notify individuals based on initials
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    
    for initials, email in {'YC': 'yuexing.chen@mssm.edu'}.items():
        initials_missing_df = missing_df.loc[missing_df['Processor Initials'] == initials]
        if len(initials_missing_df) > 0:
            mail.To = email
            mail.Subject = f"Missing Processing Information for Samples Processed on or before {cutoff_date}"
            mail.HTMLBody = f"Good morning {initials},<br><br>The following samples are missing processing information:<br><br>{initials_missing_df.to_html(index=False)}<br><br>Thank you,<br>Automatic Email DO NOT REPLY"
            mail.Send()

    print("Daily check is finished. Notification emails have been sent. Have a good day!")

