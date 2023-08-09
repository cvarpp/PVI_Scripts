import pandas as pd
import argparse
import util
import os


if __name__ == '__main__':
      
    # Rename mast_list: Location - Kit Type, Study ID - Sample ID; plog: Study - Kit Type
    mast_list = (pd.read_excel(util.tracking + "Sample ID Master List.xlsx", sheet_name='Master Sheet', header=0)
                   .rename(columns={'Location': 'Kit Type', 'Study ID': 'Sample ID'}))
    
    plog = (pd.read_excel(util.print_log, sheet_name='LOG', header=1)
              .rename(columns={'Box numbers':'Box Min', 'Unnamed: 4':'Box Max', 'Study':'Kit Type'})
              .drop(0))
    
    type = ['Standard', 'SERONET', 'PRIORITY']
    folder = {'Standard': 'STANDARD', 'SERONET': 'SERONET FULL', 'PRIORITY': 'SERUM'}

    # DRAFT: start with serum 
    for kit_type in type:
        plog_kit_type = plog[plog['Kit Type'] == kit_type]
        recent_priority = plog_kit_type[plog_kit_type['Study'] == 'PRIORITY'].iloc[-1]
        
        box_start = recent_priority['Box Max'] + 1
        box_end = box_start + 3
        workbook_name = f"{kit_type} {box_start}-{box_end}"
        folder_name = folder_name[kit_type]
        output_path = os.path.join(util.tube_print, 'Future Sheets', folder_name, f"{workbook_name}.xlsx")
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

        
        for sheet_num, sheet_name in enumerate(['1 – Kits', '2 – Tops', '3 – Sides', '4 – 4.5 mL Tops', '5 - 4.5 mL Sides'], start=1):
            df = mast_list[mast_list['Kit Type'] == kit_type].copy()
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        writer.save()

