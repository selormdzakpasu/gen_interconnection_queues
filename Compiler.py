# Author: Selorm Kwami Dzakpasu

from openpyxl import load_workbook
import pandas as pd

# Function to add ISO name in the first column of their respective queues
def add_column(file_path, iso_name, sheet_name=None):
    # Load the workbook
    wb = load_workbook(file_path)
    
    # Default to the first sheet if no sheet name is provided
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]
    
    # Insert a new column at the first position
    ws.insert_cols(1)
    
    # Set the header for the new column
    ws.cell(row=1, column=1).value = "ISO"
    
    # Fill the column with the entity name up to the last filled row
    for row in range(2, ws.max_row + 1):
        ws.cell(row=row, column=1).value = iso_name
    
    # Save the workbook
    wb.save(file_path)
    print(f"Column 'ISO' added successfully to {ws.title} in {file_path}")


# Function to combine all ISO Queues
def concatenate_excel_files(file_paths, output_file):
    # List to hold DataFrames
    dataframes = []
    
    # Read each file and append the DataFrame to the list
    for file in file_paths:
        df = pd.read_excel(file)
        dataframes.append(df)
    
    # Concatenate all DataFrames
    combined_df = pd.concat(dataframes, ignore_index=True)
    
    # Write the combined DataFrame to a new Excel file
    combined_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Files have been concatenated into {output_file}")

# File paths and corresponding ISO names
files_and_iso = [
    ("Processed Queues/CAISO_Queue.xlsx", "CAISO"),
    ("Processed Queues/ERCOT_Queue.xlsx", "ERCOT"),
    ("Processed Queues/MISO_Queue.xlsx", "MISO"),
    ("Processed Queues/PJM_Queue.xlsx", "PJM"),
    ("Processed Queues/SPP_Queue.xlsx", "SPP"),
    ("Processed Queues/NEISO_Queue.xlsx", "NEISO"),
    ("Processed Queues/NYISO_Queue.xlsx", "NYISO"),
]

# File paths for concatenation
file_paths = [
    "Processed Queues/CAISO_Queue.xlsx", "Processed Queues/ERCOT_Queue.xlsx", "Processed Queues/MISO_Queue.xlsx",
    "Processed Queues/PJM_Queue.xlsx", "Processed Queues/SPP_Queue.xlsx", "Processed Queues/NEISO_Queue.xlsx", "Processed Queues/NYISO_Queue.xlsx"
]

# Process each file, filling the first column with the ISO name
for file_path, iso_name in files_and_iso:
    add_column(file_path, iso_name)


output_file = "Combined_Queues.xlsx"
concatenate_excel_files(file_paths, output_file)
