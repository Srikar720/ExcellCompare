import os
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from tabulate import tabulate
 
# Directory containing Excel files
excel_dir = "excel_files"
 
# Ensure the directory exists
if not os.path.exists(excel_dir):
    os.makedirs(excel_dir)
    print(f"Directory '{excel_dir}' created. Please add Excel files and rerun the script.")
    exit()
 
# Extract version number from filename
def extract_version(filename):
    match = re.search(r'V(\d+)\.(\d+)\.(\d+)', filename)
    if match:
        return tuple(map(int, match.groups()))
    return (0, 0, 0)
 
# List and sort Excel files by version
excel_files = [f for f in os.listdir(excel_dir) if f.endswith('.xlsx')]
excel_files.sort(key=extract_version)
 
# Ensure at least two files are available
if len(excel_files) < 2:
    print("Not enough versioned Excel files found in the directory.")
    exit()
 
# Select the two most recent files
file_old = os.path.join(excel_dir, excel_files[-2])
file_new = os.path.join(excel_dir, excel_files[-1])
 
# ✅ Create output directory if it doesn't exist
output_dir = "Output"
os.makedirs(output_dir, exist_ok=True)
 
# ✅ Save with timestamp so artifact doesn't get overwritten
timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
output_file = os.path.join(output_dir, f"difference_{timestamp}.xlsx")
 
print(f"Old File: '{file_old}'")
print(f"New File: '{file_new}'")
 
# Load workbooks
wb_old = load_workbook(file_old)
wb_new = load_workbook(file_new)
 
# Highlight style
highlight = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='lightDown')
bold_font = Font(bold=True, color='FFFFFF')
 
# Track changes summary
summary_data = []
 
print("Processing sheets...")
 
for sheet_name in wb_old.sheetnames:
    if sheet_name not in wb_new.sheetnames:
        summary_data.append((sheet_name, 'Missing in new file'))
        continue
 
    ws_old = wb_old[sheet_name]
    ws_new = wb_new[sheet_name]
    max_row = max(ws_old.max_row, ws_new.max_row)
    max_col = max(ws_old.max_column, ws_new.max_column)
 
    changes = 0
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            val_old = ws_old.cell(row=row, column=col).value
            val_new = ws_new.cell(row=row, column=col).value
            if val_old != val_new:
                ws_new.cell(row=row, column=col).fill = highlight
                ws_new.cell(row=row, column=col).font = bold_font
                changes += 1
 
    if changes:
        summary_data.append((sheet_name, f"{changes} changes"))
    else:
        summary_data.append((sheet_name, "No changes"))
 
# Add summary sheet
ws_summary = wb_new.create_sheet("Summary_of_Changes")
ws_summary.append(["Sheet Name", "Change Summary"])
for entry in summary_data:
    ws_summary.append(entry)
 
# Print summary
print("Summary of Changes:")
print(tabulate(summary_data, headers=["Sheet Name", "Change Summary"], tablefmt="grid"))
 
#  Save output file
wb_new.save(output_file)
print(f"Saving output file to: {output_file}")
print("File saved successfully.")
 
