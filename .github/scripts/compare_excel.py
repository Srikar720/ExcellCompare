import pandas as pd
import os

def compare_excels(old_file, new_file):
    old_df = pd.read_excel(old_file, engine='openpyxl')
    new_df = pd.read_excel(new_file, engine='openpyxl')

    comparison_result = []

    for index, row in new_df.iterrows():
        if index >= len(old_df):
            comparison_result.append(f"Added row {index}: {row.to_dict()}")
        elif not row.equals(old_df.iloc[index]):
            comparison_result.append(f"Changed row {index}:\nOld: {old_df.iloc[index].to_dict()}\nNew: {row.to_dict()}")

    if len(old_df) > len(new_df):
        for index in range(len(new_df), len(old_df)):
            comparison_result.append(f"Removed row {index}: {old_df.iloc[index].to_dict()}")

    return comparison_result

if os.path.exists('old_data.xlsx') and os.path.exists('new_data.xlsx'):
    differences = compare_excels('old_data.xlsx', 'new_data.xlsx')
    with open('excel_diff_report.txt', 'w') as report:
        for line in differences:
            report.write(line + '\n')
    print("Comparison complete. Differences written to excel_diff_report.txt.")
else:
    print("One or both Excel files are missing.")
