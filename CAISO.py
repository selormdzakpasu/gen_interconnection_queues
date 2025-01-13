# Author: Selorm Kwami Dzakpasu

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, NamedStyle
from datetime import datetime

def process_caiso_file(file_path):
    # Check the file extension
    file_extension = file_path.split('.')[-1].lower()

    if file_extension not in ['csv', 'xlsx']:
        raise ValueError("Unsupported file format. Please provide a .csv or .xlsx file.")

    # Load the file into a DataFrame
    if file_extension == 'csv':
        df = pd.read_csv(file_path, header=None)
    else:
        df = pd.read_excel(file_path, header=None)

    # Remove the first three rows
    df = df.iloc[3:, :]

    # Reset the index after dropping rows
    df.reset_index(drop=True, inplace=True)

    # Set the fourth row as column headers
    df.columns = df.iloc[0]
    df = df[1:]

    # Reset the index again
    df.reset_index(drop=True, inplace=True)

    # Check for two consecutive empty cells in the first column
    empty_cell_count = 0
    last_filled_row = len(df) - 1  # Default to the last row if no empty cells are found

    for i in range(len(df)):
        if pd.isnull(df.iloc[i, 0]):  # Check if the cell is empty (NaN)
            empty_cell_count += 1
        else:
            empty_cell_count = 0
        
        # If two consecutive empty cells are found, break the loop
        if empty_cell_count == 2:
            last_filled_row = i - 2  # Set to the row before the first empty cell
            break

    # Slice the DataFrame to keep only the rows up to the last filled row in the first column
    df = df.iloc[:last_filled_row + 1, :]

    # Insert new columns into the specified positions first
    new_columns = {
        6: "Technology",
        10: "Fuel",
        14: "Capacity (MW)",
        19: "Energy"
    }

    for position, header in new_columns.items():
        # Insert a new column at the specified position
        if position >= len(df.columns):
            # Extend the DataFrame to ensure the position exists
            while len(df.columns) <= position:
                df[len(df.columns)] = ''

        # Insert the new column header
        df.insert(position, header, '')

    # Rename specific columns based on the provided mapping
    column_rename_map = {
        2: "Interconnection Request Receive Date",
        5: "Study Process",
        7: "Technology-1",
        8: "Technology-2",
        9: "Technology-3",
        15: "Capacity-1",
        16: "Capacity-2",
        17: "Capacity-3",
        20: "Energy-1",
        21: "Energy-2",
        22: "Energy-3",
        27: "Nearest Town or County",
        31: "POI Name",
        32: "Proposed In-Service/Initial Backfeed Date",
        33: "Projected COD",
        35: "Feasibility Study Status",
        36: "System Impact Study Status",
        37: "Facilities Study Status",
        39: "Interconnection Agreement Status"
    }

    for position, new_name in column_rename_map.items():
        if position < len(df.columns):
            df.columns.values[position] = new_name

    # Save the modified DataFrame back to a file
    output_file = f"CAISO_Queue.xlsx"
    
    # Save to Excel
    df.to_excel(output_file, index=False, engine='openpyxl')

    # load workbook
    wb = load_workbook(output_file)
    ws = wb.active
    
    # Create a date format style
    date_style = NamedStyle(name="short_date", number_format="MM/DD/YYYY")

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, datetime):  # Check if the cell contains a date
                cell.style = date_style  # Apply short date format
            cell.alignment = Alignment(horizontal="left", vertical="center") # Apply cell alignment

    wb.save(output_file)

    print(f"File processed and saved as {output_file}")

# Specify file name and run
file_path = "CAISO Queue Report.xlsx"  # Replace with your file name
process_caiso_file(file_path)