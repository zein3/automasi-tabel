#!/usr/bin/env python3

import sys
import os
import pandas as pd

def is_majority_zeros_or_small(df, threshold=5):
    non_zero_values = df[df != 0].count().sum()
    # total_values = df.size
    return non_zero_values <= threshold


def main():
    if (len(sys.argv) < 3):
        print("Penggunaan: merge.py {nama folder berisi file-file excel} {nama folder output}")
        exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2]
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Loop melalui semua file di direktori sumber
    for filename in os.listdir(input_folder):
        if filename.endswith('.xlsx'):
            filepath = os.path.join(input_folder, filename)
            
            # Baca file Excel
            excel_data = pd.ExcelFile(filepath)
            
            # Cari sheet yang diinginkan
            target_sheet_name = None
            for sheet_name in excel_data.sheet_names:
                if sheet_name.endswith('_kec'):
                    target_sheet_name = sheet_name
                    break
            
            if target_sheet_name:
                # Baca sheet yang diinginkan
                df = excel_data.parse(target_sheet_name)
                
                # Filter baris dengan kolom kab yang bernilai 3276
                filtered_df = df[df['kab'] == 3276]
                
                # Periksa apakah mayoritas nilai dalam tabel adalah 0 atau kecil
                if not is_majority_zeros_or_small(filtered_df):
                    # Simpan file baru jika kondisi tidak terpenuhi
                    new_filename = os.path.join(output_folder, filename)
                    with pd.ExcelWriter(new_filename, engine='openpyxl') as writer:
                        filtered_df.to_excel(writer, sheet_name=target_sheet_name, index=False)

if __name__ == '__main__':
    main()