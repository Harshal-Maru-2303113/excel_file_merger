import pandas as pd
import glob
import os
import shutil
from datetime import datetime

# Specify the folder where Excel files are stored
folder_path = r"C:/Users/harsh/Documents/Jan/"  # Change this to your folder path
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

# Generate the output filename with today's date
today_date = datetime.today().strftime("%Y-%m-%d")
output_file = os.path.join(output_folder, f"{today_date}.xlsx")

# Backup existing file
if os.path.exists(output_file):
    timestamp = datetime.now().strftime("%H-%M-%S")
    backup_file = os.path.join(output_folder, f"{today_date}_{timestamp}.xlsx")
    shutil.copy(output_file, backup_file)
    print(f"\U0001F4C2 Backup created: {backup_file}")

# Process each file and append directly to the output Excel file
for file in excel_files:
    print(f"\n🔍 Extracting data from: {file}")

    try:
        temp_df = pd.read_excel(file, sheet_name='Data', header=None)

        if temp_df.empty:
            print(f"⚠ The 'Data' sheet in {file} is empty. Skipping...")
            continue

        start_row = detect_header(temp_df)
        df = pd.read_excel(file, sheet_name='Data', header=start_row)
        df.dropna(how='all', inplace=True)  # Remove empty rows

        if os.path.exists(output_file):
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                existing_df = pd.read_excel(output_file, sheet_name='Sheet1')
                df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=len(existing_df) + 1)
        else:
            df.to_excel(output_file, index=False, engine='openpyxl')

        print(f"✅ Successfully appended data from {file}")

    except ValueError as ve:
        print(f"❌ Sheet 'Data' not found in {file}. Skipping... Error: {ve}")
    except Exception as e:
        print(f"❌ Error collecting data from {file}. Skipping... Error: {e}")

print(f"\n📁 Data successfully updated in {output_file}")
