# -*- coding: utf-8 -*-
"""
Created on Tue Jan  6 15:52:33 2026

@author: basanel
"""

import os
import pandas as pd

folder_path = r"C:\Users\basanel\Downloads\QA real. test"
output_file = r"C:\Users\basanel\Downloads\Combined_Data.xlsx"

all_data = []

for filename in os.listdir(folder_path):
    if filename.endswith(".xls"):
        file_path = os.path.join(folder_path, filename)
        try:
            # Read all sheets
            xls = pd.read_excel(file_path, sheet_name=None, engine='xlrd')  # engine may need xlrd for .xls

            for sheet_name, df in xls.items():
                df.insert(0, 'SheetName', sheet_name)
                all_data.append(df)

        except Exception as e:
            print(f"Error reading {filename}: {e}")


if all_data:
    combined_df = pd.concat(all_data, ignore_index=True)
    combined_df.to_excel(output_file, index=False, sheet_name='Sheet1')
    print("✅ Combined file saved to:", output_file)
else:
    print("❌ No data combined. Check folder or file format.")