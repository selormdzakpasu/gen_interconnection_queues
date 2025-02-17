# Author: Selorm Kwami Dzakpasu

import pandas as pd

def process_pjm_file(file_path):
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
    
    # Drop redundant columns
    df.drop(['MFO', 'MW In Service', 'Project (AC/DC)', 'Rights (MW)', 'Initial Study', 'Feasibility Study', 'System Impact Study', 'Facilities Study',
             'Interim/Interconnection Service/Generation Interconnection Agreement', 'Wholesale Market Participation Agreement', 'Construction Service Agreement',
             'Construction Service Agreement Status', 'Upgrade Construction Service Agreement', 'Upgrade Construction Service Agreement Status', 'Backfeed Date',
             'Long-Term Firm Service Start Date', 'Long-Term Firm Service End Date', 'Test Energy Date', 'Revised In Service Date', 'Attachment Type',
             'Alternate Project', 'Sliding Project'], axis=1, inplace=True)

    # Rename specific columns based on the provided mapping
    columns_to_rename = {
    "Name": "POI Name",
    "Commercial Name": "Project Name",
    "Status": "Project Status",
    "MW Capacity": "Capacity (MW)",
    "Capacity or Energy": "Capacity Status",
    "Project Type": "Service Type",
    "Fuel": "Technology",
    "Interim/Interconnection Service/Generation Interconnection Agreement Status": "Interconnection Agreement Status",
    "Submitted Date": "Queue Date",
    "Commercial Operation Milestone": "Projected COD",
    "Withdrawn Remarks": "Comment"
    }
    
    # Rename Columns
    df.rename(columns=columns_to_rename, inplace=True)

    # Save the modified DataFrame back to a file
    output_file = f"Processed Queues/PJM_Queue.xlsx"
    
    # Save to Excel
    df.to_excel(output_file, index=False, engine='openpyxl')

    print(f"File processed and saved at {output_file}")

# Specify file name and run
file_path = "Queues/PJM Queue.xlsx"  # Replace with your file name
process_pjm_file(file_path)