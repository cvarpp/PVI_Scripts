import pandas as pd
import re
from util import pbmc_shipping, serum_shipping

#Pick Serum or PBMC

df = pd.read_excel(pbmc_shipping + '4.2.24 PBMC Shipment/shipping_manifest.xlsx')
#df = pd.read_excel(serum_shipping + '/shipping_manifest.xlsx')

def validate_data(df):
    for index, row in df.iterrows():
            # Check 1: The 10th character of Current Label based on Material Type
        if row['Material Type'] == 'PBMC':
            assert row['Current Label'][9] == '2', f"Row {index+2}: 10th digit must be 2 for PBMC in 'Current Label'"
        elif row['Material Type'] == 'SERUM':
            assert row['Current Label'][9] == '1', f"Row {index+2}: 10th digit must be 1 for SERUM in 'Current Label'"
        # Check 2: Volume should be 1 if Material Type is PBMC
        if row['Material Type'] == 'PBMC':
            assert row['Volume'] == 1, f"Row {index+2}: Volume should be 1 for PBMC"
        # Check 3: Volume Estimate should always be 'Actual'
        assert row['Volume Estimate'] == 'Actual', f"Row {index+2}: Volume Estimate should always be 'Actual'"
        # Check 4: Study ID should always be 'LP003'
        assert row['Study ID'] == 'LP003', f"Row {index+2}: Study ID should always be 'LP003'" 
        # Check 5: Vial Type should be 'CRYOTUBE 1.8ML' if Material Type is PBMC
        if row['Material Type'] == 'PBMC':
            assert row['Vial Type'] == 'CRYOTUBE 1.8ML', f"Row {index+2}: Incorrect Vial Type for PBMC"

validate_data(df)
