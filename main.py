import pandas as pd
import glob
import os
from datetime import datetime

# Specify the folder where Excel files are stored
folder_path = r""  # Change this to your folder path
output_folder = "output"  # Folder where the final Excel file will be saved

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Get all Excel files in the folder
excel_files = glob.glob(f"{folder_path}/*.xlsb")  # Use "*.xls" if needed

def detect_header(df):
    for i, row in df.iterrows():
        non_empty_cells = row.notna().sum()
        if non_empty_cells >= 2:  # Assuming headers usually have multiple non-empty cells
            return i
    return 0  # Default to first row if none found

# Read data from all Excel files
dataframes = []
header_included = False  # Flag to include header only once

for file in excel_files:
    try:
        temp_df = pd.read_excel(file, sheet_name='Data', header=None)
        if temp_df.empty:
            print(f"The 'Data' sheet in {file} is empty.")
            continue

        start_row = detect_header(temp_df)

        if not header_included:
            df = pd.read_excel(file, sheet_name='Data', header=start_row)
            header_included = True
            column_names = df.columns  # Store the header from the first file
        else:
            df = pd.read_excel(file, sheet_name='Data', header=None, skiprows=start_row)
            df.columns = column_names  # Assign the same header as the first file

        # Remove empty rows
        df.dropna(how='all', inplace=True)

        # Remove potential duplicate header rows
        df = df[df[column_names[0]] != column_names[0]]

        if not df.empty:
            dataframes.append(df)
            print(f"Successfully extracted data from {file}")
        else:
            print(f"No data found after processing {file}.")

    except ValueError as ve:
        print(f"Sheet 'Data' not found in {file}: {ve}")
    except Exception as e:
        print(f"Error collecting data from {file}: {e}")

# Combine all data into a single DataFrame and save it
if dataframes:
    final_df = pd.concat(dataframes, ignore_index=True)

    # Generate the output filename with today's date
    today_date = datetime.today().strftime("%Y-%m-%d")
    output_file = os.path.join(output_folder, f"{today_date}.xlsx")

    # Save to Excel
    final_df.to_excel(output_file, index=False)
    print(f"Data saved successfully to {output_file}")
else:
    print("No data found.")