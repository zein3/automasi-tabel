#!/usr/bin/env python3

import os
import sys
import glob
import pandas as pd

def get_excel_files(folder_path):
    # Define the pattern to match Excel files
    pattern_xlsx = os.path.join(folder_path, '*.xlsx')
    pattern_xls = os.path.join(folder_path, '*.xls')
    # Get a list of all Excel files in the given folder
    excel_files = glob.glob(pattern_xlsx) + glob.glob(pattern_xls)
    return excel_files

def merge_excel_files_into_sheets(folder_path, output_file):
    excel_files = get_excel_files(folder_path)
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for file in excel_files:
            excel_data = pd.ExcelFile(file)
            for sheet_name in excel_data.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name)
                # Create a unique sheet name if there are duplicates
                if sheet_name in writer.sheets:
                    sheet_name = f"{sheet_name}_{os.path.splitext(os.path.basename(file))[0]}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"All files have been merged into {output_file} with separate sheets.")

def main():
    if (len(sys.argv) < 3):
        print("Penggunaan: merge.py {folder} {file excel}")
        sys.exit(1)

    folder_path = sys.argv[1]
    output_file = sys.argv[2]
    merge_excel_files_into_sheets(folder_path, output_file)


if __name__ == '__main__':
    main()