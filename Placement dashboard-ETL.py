#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Aug  7 23:22:53 2025

@author: Bhumika
"""

import pandas as pd
import os

# Input file path
input_file = "/Users/Bhumika/Desktop/Placement/IPRS_bi.xlsx"

# Output folder path
output_folder = "/Users/Bhumika/Desktop/Placement/output"
output_file = os.path.join(output_folder, "IPRS_Exported.xlsx")

# Ensure output folder exists
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
else:
    # Delete all files in the output folder
    for file_name in os.listdir(output_folder):
        file_path = os.path.join(output_folder, file_name)
        if os.path.isfile(file_path):
            os.remove(file_path)
            print(f"Deleted: {file_path}")

# Load the Excel file
a = pd.read_excel(input_file)

# Check for nulls and export
if a.isnull().sum().sum() == 0:
    a.to_excel(output_file, index=False) 
    print(f"Exported new file to: {output_file}")
else:
    print("Data contains null values. Export aborted.")

 

