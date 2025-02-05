import pandas as pd
import glob
import os
from datetime import datetime

# Specify the folder where Excel files are stored
folder_path = "files"  # Change this to your folder path
output_folder = "output"  # Folder where the final Excel file will be saved

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Get all Excel files in the folder
excel_files = glob.glob(f"{folder_path}/*.xlsx")  # Use "*.xls" if needed

def detect_header(df):
    for i, row in df.iterrows():
        for j, cell in enumerate(row):
            if pd.notna(cell):
                return i, j  # Return the first non-empty cell's row and column index
    return 0, 0  # Default to first cell if none found

# Read data from all Excel files
dataframes = []
header_included = False  # Flag to include header only once

for file in excel_files:
    try:
        temp_df = pd.read_excel(file, header=None)
        start_row, start_col = detect_header(temp_df)

        if not header_included:
            df = pd.read_excel(file, header=start_row, usecols=range(start_col, temp_df.shape[1]))
            header_included = True
            column_names = df.columns  # Store the header from the first file
        else:
            df = pd.read_excel(file, header=None, skiprows=start_row, usecols=range(start_col, temp_df.shape[1]))
            df.columns = column_names  # Assign the same header as the first file

        # Remove potential duplicate header rows
        df = df[df[column_names[0]] != column_names[0]]

        dataframes.append(df)
        print(f"Successfully extracted data from {file}")
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