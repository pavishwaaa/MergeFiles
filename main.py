import pandas as pd
import os

# Path to the folder containing the reports
folder_path = 'D:/Python projects/MergeFiles/InventoryDashboard'

# Get a list of all files in the folder
files = os.listdir(folder_path)

# Create an empty DataFrame to hold the combined data
combined_data = pd.DataFrame()

# Flag to keep track of header inclusion
header_added = False

# Iterate through each file in the folder
for file in files:
    file_path = os.path.join(folder_path, file)

    # Read the file skipping the first 10 rows (assuming they're in Excel format, modify for CSV if needed)
    data = pd.read_excel(file_path, skiprows=10)

    try:
     # Append data to the combined DataFrame
        if not header_added:
            combined_data = combined_data._append(data, ignore_index=True)  # Assign the first file's data as the initial combined data
            header_added = True
        else:
            combined_data = combined_data._append(data, ignore_index=True)
    except Exception as e:
        print(f"Error reading {file}: {e}")
        continue

# Write the combined data to a single Excel file
combined_data.to_excel(r'D:\Python projects\MergeFiles\Output\combined_reports1.xlsx', index=False)
