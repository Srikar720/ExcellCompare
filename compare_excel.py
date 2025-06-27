import os
import pandas as pd
import re
from datetime import datetime

# Folder with Excel files
folder = 'excel_files'
output_file = 'comparison_output.txt'

# Get all Excel files with date in filename
files = []
for f in os.listdir(folder):
    match = re.search(r'_(\d{8})\.xlsx$', f)
    if match:
        date = datetime.strptime(match.group(1), '%Y%m%d')
        files.append((f, date))

# Sort files by date
files.sort(key=lambda x: x[1])

# Only compare if we have at least two files
if len(files) >= 2:
    old_file = os.path.join(folder, files[-2][0])
    new_file = os.path.join(folder, files[-1][0])

    old_excel = pd.ExcelFile(old_file, engine='openpyxl')
    new_excel = pd.ExcelFile(new_file, engine='openpyxl')

    with open(output_file, 'w') as out:
        out.write(f"Comparing {files[-2][0]} with {files[-1][0]}\n")
        changes_found = False

        for sheet in old_excel.sheet_names:
            if sheet in new_excel.sheet_names:
                df_old = old_excel.parse(sheet)
                df_new = new_excel.parse(sheet)

                for i in range(max(len(df_old), len(df_new))):
                    for j in range(max(len(df_old.columns), len(df_new.columns))):
                        val_old = df_old.iloc[i, j] if i < len(df_old) and j < len(df_old.columns) else None
                        val_new = df_new.iloc[i, j] if i < len(df_new) and j < len(df_new.columns) else None
                        if val_old != val_new:
                            out.write(f"Sheet '{sheet}' Cell ({i+1},{j+1}): '{val_old}' -> '{val_new}'\n")
                            changes_found = True

        if not changes_found:
            out.write("No changes detected.\n")
else:
    with open(output_file, 'w') as out:
        out.write("Not enough files to compare.\n")
