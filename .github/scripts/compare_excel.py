import os
import re
import pandas as pd
from datetime import datetime

# Directory containing Excel files
excel_dir = 'excel_files'  # Make sure this folder exists and contains your .xlsx files

# Regex pattern to extract date from filename (e.g., report_20240601.xlsx)
date_pattern = re.compile(r'_(\d{8})\.xlsx$')

# Group files by base name (excluding date)
file_groups = {}

for filename in os.listdir(excel_dir):
    match = date_pattern.search(filename)
    if match:
        date_str = match.group(1)
        base_name = filename[:match.start()]
        file_groups.setdefault(base_name, []).append((filename, datetime.strptime(date_str, '%Y%m%d')))

# Function to compare two dataframes
def compare_dataframes(df_old, df_new):
    differences = []
    max_rows = max(len(df_old), len(df_new))
    max_cols = max(len(df_old.columns), len(df_new.columns))

    for row in range(max_rows):
        for col in range(max_cols):
            val_old = df_old.iloc[row, col] if row < len(df_old) and col < len(df_old.columns) else None
            val_new = df_new.iloc[row, col] if row < len(df_new) and col < len(df_new.columns) else None
            if val_old != val_new:
                differences.append((row, col, val_old, val_new))
    return differences

# Compare latest and previous versions
for base_name, files in file_groups.items():
    if len(files) < 2:
        continue

    sorted_files = sorted(files, key=lambda x: x[1])
    old_file, new_file = sorted_files[-2], sorted_files[-1]

    old_path = os.path.join(excel_dir, old_file[0])
    new_path = os.path.join(excel_dir, new_file[0])

    print(f"\nComparing '{old_file[0]}' with '{new_file[0]}'")

    old_excel = pd.ExcelFile(old_path, engine='openpyxl')
    new_excel = pd.ExcelFile(new_path, engine='openpyxl')

    common_sheets = set(old_excel.sheet_names).intersection(new_excel.sheet_names)

    for sheet in common_sheets:
        df_old = old_excel.parse(sheet)
        df_new = new_excel.parse(sheet)
        differences = compare_dataframes(df_old, df_new)

        if differences:
            print(f"  Differences found in sheet '{sheet}':")
            for row, col, val_old, val_new in differences:
                print(f"    Cell ({row+1}, {col+1}): Old='{val_old}' | New='{val_new}'")
        else:
            print(f"  No differences in sheet '{sheet}'.")

    old_only_sheets = set(old_excel.sheet_names) - set(new_excel.sheet_names)
    new_only_sheets = set(new_excel.sheet_names) - set(old_excel.sheet_names)

    if old_only_sheets:
        print(f"  Sheets only in old file: {', '.join(old_only_sheets)}")
    if new_only_sheets:
        print(f"  Sheets only in new file: {', '.join(new_only_sheets)}")
