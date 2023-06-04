# -*- coding: utf-8 -*-
"""
Created on Mon May 15 16:36:53 2023

@author: hakany
"""

import pandas as pd

# Set the PYTHONIOENCODING environment variable
import os
os.environ["PYTHONIOENCODING"] = "utf-8"

# Read the Excel file
excel_file = pd.ExcelFile('ExcelName.xlsx')

# Iterate over each sheet in the Excel file
for sheet_name in excel_file.sheet_names:
    # Read the sheet data
    data = excel_file.parse(sheet_name)

    # Get the column numbers (indices) based on your requirement
    start_column = 0  # Column number to start generating SQL from
    end_column = 21  # Column number to end generating SQL at (inclusive)

    # Generate SQL insert statements
    sql_statements = []
    for idx, row in data.iterrows():
        values = row.values[start_column:end_column+1].tolist()


        # Convert NEWID() to NULL if the value is NaN (missing value)
        values = [np.nan if value == "NEWID()" else value for value in values]


    
        values = [f"'{str(value)}'" if pd.notnull(value) else 'NEWID()' for value in values]
        sql = f"INSERT INTO table_name VALUES ({', '.join(values)});"
        sql_statements.append(sql)

    # Save the SQL statements to a file for the current sheet
    with open(f'output_{sheet_name}.sql', 'w', encoding='utf-8') as file:
        file.write('\n'.join(sql_statements))
