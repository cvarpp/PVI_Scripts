import os
import pandas as pd
import util

row_labels = list("ABCDEFGHIJ")  # Row A-J

tank3_folder = os.path.join(util.psp, 'Liquid Nitrogen Inventory (USE THIS!!!)', 'TANK3')
output_file = os.path.join(tank3_folder, 'Tank 3 Master Sheet.xlsx')

def read_box(sheet_df, sample_type="PBMC"):
    records = []
    for row_idx, row_letter in enumerate(row_labels):
        for col_idx in range(1, 11):  # Column 1-10
            try:
                cell_value = sheet_df.iloc[row_idx, col_idx - 1]
            except IndexError:
                continue  # Skip incomplete table

            if pd.isna(cell_value) or str(cell_value).strip().lower() == 'x':
                continue  # Skip empty or X

            cell_value = str(cell_value).strip()

            if ' ' in cell_value:
                sample_id, date = cell_value.split(' ', 1)
            else:
                sample_id = cell_value
                date = ''

            record = {
                'Sample ID': cell_value,
                'Sample Type': sample_type,
                'Name': cell_value,
                'Row': row_letter,
                'Column': col_idx,
                'Position': f"{col_idx}/{row_letter}",
                'Date': date.strip()
            }
            records.append(record)
    return records

def process_workbook(filepath, sample_type="PBMC"):
    all_records = []
    excel_file = pd.ExcelFile(filepath)
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        records = read_box(df, sample_type)
        all_records.extend(records)
    return all_records

def main():
    all_records = []

    for file in os.listdir(tank3_folder):
        if file.endswith(".xlsx") and file != 'Tank 3 Master Sheet.xlsx' and not file.startswith('~$'):
            file_path = os.path.join(tank3_folder, file)
            print(f"Processing {file}...")
            records = process_workbook(file_path)
            all_records.extend(records)

    master_df = pd.DataFrame(all_records)
    master_df = master_df[['Sample ID', 'Sample Type', 'Name', 'Row', 'Column', 'Position', 'Date']]

    master_df.to_excel(output_file, index=False)
    print(f"Master file saved to: {output_file}")

if __name__ == "__main__":
    main()
