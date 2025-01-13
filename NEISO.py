# Author: Selorm Kwami Dzakpasu

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, NamedStyle
from datetime import datetime

def process_neiso_file(file_path):
    # Check the file extension
    file_extension = file_path.split('.')[-1].lower()

    if file_extension not in ['csv', 'xlsx']:
        raise ValueError("Unsupported file format. Please provide a .csv or .xlsx file.")

    # Load the file into a DataFrame
    if file_extension == 'csv':
        df = pd.read_csv(file_path, header=None)
    else:
        df = pd.read_excel(file_path, header=None)

    # Remove the first four rows
    df = df.iloc[4:, :]

    # Reset the index after dropping rows
    df.reset_index(drop=True, inplace=True)

    # Set the fifth row as column headers
    df.columns = df.iloc[0]
    df = df[1:]

    # Reset the index again
    df.reset_index(drop=True, inplace=True)
    
    # Delete columns at specific indices
    df.drop(df.columns[[16, 19, 20, 21]], axis=1, inplace=True)
    
    # Reset the index again
    df.reset_index(drop=True, inplace=True)

    # Rename specific columns based on the provided mapping
    column_rename_map = {
        0: "Queue Position",
        1: "Date Updated",
        2: "Service Type",
        3: "Interconnection Request Receive Date",
        4: "Project Name",
        5: "Technology",
        6: "Fuel",
        7: "Net MWs to Grid",
        10: "Nearest Town or County",
        12: "Projected COD",
        13: "Proposed Sync Date",
        14: "Withdrawal Date",
        15: "POI Name",
        18: "Feasibility Study Status",
        19: "System Impact Study Status",
        20: "Optional Study",
        21: "Facilities Study Status",
        22: "Interconnection Agreement Status"
    }

    for position, new_name in column_rename_map.items():
        if position < len(df.columns):
            df.columns.values[position] = new_name

    # Save the modified DataFrame back to a file
    output_file = f"NEISO_Queue.xlsx"
    
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
            cell.alignment = Alignment(horizontal="left", vertical="center") # Apply alignment to cells

    wb.save(output_file)

    print(f"File processed and saved as {output_file}")

# Specify file name and run
file_path = "NEISO QueueReport_20250107130953.xlsx"  # Replace with your file name
process_neiso_file(file_path)