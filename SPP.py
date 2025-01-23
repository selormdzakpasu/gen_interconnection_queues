# Author: Selorm Kwami Dzakpasu

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from datetime import datetime

def process_spp_file(file_path):
    # Check the file extension
    file_extension = file_path.split('.')[-1].lower()

    if file_extension not in ['csv', 'xlsx']:
        raise ValueError("Unsupported file format. Please provide a .csv or .xlsx file.")

    # Load the file into a DataFrame
    if file_extension == 'csv':
        df = pd.read_csv(file_path, header=None)
    else:
        df = pd.read_excel(file_path, header=None)
   
    # Remove the first row
    df = df.iloc[1:, :]

    # Reset the index after dropping row
    df.reset_index(drop=True, inplace=True)
    
    # Set the second row as column headers
    df.columns = df.iloc[0]
    df = df[1:]

    # Reset the index again
    df.reset_index(drop=True, inplace=True)

    # Rename specific columns based on the provided mapping
    column_rename_map = {
        4: "Nearest Town or County",
        6: "Transmission Owner/Developer",
        7: "Proposed In-Service/Initial Backfeed Date",
        8: "Projected COD",
        11: "Capacity (MW)",
        15: "Technology",
        16: "Fuel",
        17: "POI Name",
        18: "Interconnection Request Receive Date",
        19: "Withdrawal Date",
        20: "Project Status"
    }

    for position, new_name in column_rename_map.items():
        if position < len(df.columns):
            df.columns.values[position] = new_name    

    # Save the modified DataFrame back to a file
    output_file = f"Processed Queues/SPP_Queue.xlsx"
    
    # Save to Excel
    df.to_excel(output_file, index=False, engine='openpyxl')

    # Load workbook
    wb = load_workbook(output_file)
    ws = wb.active

    # Create a date format style
    date_style = NamedStyle(name="short_date", number_format="MM/DD/YYYY")

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, datetime):  # Check if the cell contains a date
                cell.style = date_style  # Apply short date format

    wb.save(output_file)

    print(f"File processed and saved at {output_file}")

# Specify file name and run
file_path = "Queues/SPP GI_ActiveRequest.xlsx"  # Replace with your file name
process_spp_file(file_path)