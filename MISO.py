# Author: Selorm Kwami Dzakpasu

import pandas as pd

def process_miso_file(file_path):
    # Check the file extension
    file_extension = file_path.split('.')[-1].lower()

    if file_extension not in ['csv', 'xlsx']:
        raise ValueError("Unsupported file format. Please provide a .csv or .xlsx file.")

    # Load the file into a DataFrame
    if file_extension == 'csv':
        df = pd.read_csv(file_path, header=None)
    else:
        df = pd.read_excel(file_path, header=None)

    # Set the first row as column headers
    df.columns = df.iloc[0]
    df = df[1:]

    # Reset the index
    df.reset_index(drop=True, inplace=True)

    # Rename specific columns based on the provided mapping
    column_rename_map = {
        0: "Project ID",
        1: "Application Status",
        3: "Withdrawal Date",
        5: "Proposed In-Service/Initial Backfeed Date",
        6: "Transmission Owner/Developer",
        7: "Nearest Town or County",
        17: "Technology",
        19: "Projected COD"
    }

    for position, new_name in column_rename_map.items():
        if position < len(df.columns):
            df.columns.values[position] = new_name

    # Save the modified DataFrame back to a file
    output_file = f"Processed Queues/MISO_Queue.xlsx"
    
    # Save to Excel
    df.to_excel(output_file, index=False, engine='openpyxl')

    print(f"File processed and saved at {output_file}")

# Specify file name and run
file_path = "Queues/MISO GI Interactive Queue.csv"  # Replace with your file name
process_miso_file(file_path)