#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep 27 11:28:37 2023

@author: kakshay
"""
import pandas as pd
import os
import numpy as np

input_file = input("Enter the file name with extension: ")
print("Converting file to CSV and Excel. Please wait!")

#%%
# Specify the input .plt file and output CSV file paths
#input_file = "test_6.plt"
excel_output_file = os.path.splitext(input_file)[0] + ".xlsx"
csv_output_file = os.path.splitext(input_file)[0] + ".csv"
# Initialize variables to store data
headers = ["Timestamp(s)"]
body = []

# Read the .plt file
with open(input_file, "r") as plt_file:
    parsed_data = plt_file.readlines()

header_data = [] 
for item in parsed_data:
    try:
        header_data.append(float(item))
    except ValueError:
        header_data.append(item)

index = None
# Parse the header information
for i, value in enumerate(header_data):
    if value == 0.0:
        index = i
        break
    if i % 2 == 0 and i >= 4:
        # Even-indexed lines are headers
        header = value.strip()
        headers.append(header)
                 
# Parse the data section
for value in parsed_data[index:]:
    columns = value.strip().split()  # Split columns based on spaces or tabs  
    body.append(columns)
#%%
# Number of columns in your DataFrame
num_columns = (len(headers)-1)*6 + 1

# Initialize an empty list to store the rows of data
data_rows = []

# Iterate through the inner lists and fill the DataFrame
current_row = []
current_col = 0

for row in body:
    for item in row:
        current_row.append(item)
        current_col += 1

        if current_col >= num_columns:
            data_rows.append(current_row[:num_columns])  # Add a complete row
            current_row = current_row[num_columns:]  # Store the remaining items for the next row
            current_col = 0

# If there are remaining items in the last row, add it
if current_row:
    data_rows.append(current_row)

# Create a DataFrame from the data_rows
df = pd.concat([pd.Series(row) for row in data_rows], axis=1).T

# If the DataFrame doesn't fill the last row, fill it with NaN
if len(data_rows) > 0:
    last_row = data_rows[-1]
    last_row.extend([None] * (num_columns - len(last_row)))
    data_rows[-1] = last_row

# Create the final DataFrame
df = pd.concat([pd.Series(row) for row in data_rows], axis=1).T

# Convert to numeric
df = df.apply(pd.to_numeric, errors='coerce')
# Copy data frame for future processing
df1 = df.copy()

# List of level two headings
level_two = ['x', 'y', 'z', 'u', 'v', 'w']

# Assign the multi-index column to the DataFrame
df.columns = pd.MultiIndex.from_tuples([(headers[0], '')] + [(h, l) for h in headers[1:] for l in level_two])

#%%
df2 = pd.DataFrame()

# Iterate through the columns, len(level_two) columns at a time
for start in range(1, num_columns, len(level_two)):
    end = start + len(level_two)
    col_chunk = df1.iloc[:, start:end]

    # Extract the first column (timestamps)
    timestamp_col = df1.iloc[:, 0]
    
    # Add the timestamps to the col_chunk
    col_chunk = pd.concat([timestamp_col, col_chunk], axis=1)
    
    # To sort Max and Min vaues from df dataframe
    chunk_df = pd.DataFrame()
    
    # Iterate through each column to find the maximum value and its corresponding row
    for column in col_chunk.columns[1:]:
        min_value = col_chunk[column].min()
        max_value = col_chunk[column].max()
        min_row = col_chunk[col_chunk[column] == min_value].sample(n=1)  
        max_row = col_chunk[col_chunk[column] == max_value].sample(n=1)    
        # Concatenate the row with the maximum value to the temporary DataFrame
        chunk_df = pd.concat([chunk_df, max_row, min_row])
        
        # To calculate max vector sum and its corresponding values
        # Select the x,y,z columns (except timestamp column) and calculate the row vector sum
        row_vector_sum_1 = np.sqrt((chunk_df.iloc[:, 1:4] ** 2).sum(axis=1))
        max_vector_value_1 = row_vector_sum_1.max()
        max_vector_row_num_1 = row_vector_sum_1.idxmax()
        # Retrieve all values in the row with the maximum row_vector_sum
        max_vector_row_1 = col_chunk.loc[max_vector_row_num_1]
        
        # Select the u,v,w columns (except timestamp column) and calculate the row vector sum
        row_vector_sum_2 = np.sqrt((chunk_df.iloc[:, 4:7] ** 2).sum(axis=1))
        max_vector_value_2 = row_vector_sum_2.max()
        max_vector_row_num_2 = row_vector_sum_2.idxmax()
        # Retrieve all values in the row with the maximum row_vector_sum
        max_vector_row_2 = col_chunk.loc[max_vector_row_num_2]
        
    df2= pd.concat([df2, chunk_df, max_vector_row_1.to_frame().T, max_vector_row_2.to_frame().T], ignore_index=True)
    # Reset the index and drop rows with NaN values
    df2 = df2.reset_index(drop=True)
    
# Assign the multi-index column to the DataFrame
df2.columns = pd.MultiIndex.from_tuples([(headers[0], '')] + [(h, l) for h in headers[1:] for l in level_two])
    
#%%
#To adjust first level cell width
# Access the first row (header) of the DataFrame
first_row = df.iloc[0]

# Create a Pandas ExcelWriter object
with pd.ExcelWriter(excel_output_file, engine='xlsxwriter') as writer:
    # Write the DataFrame to the Excel file
    df.to_excel(writer, sheet_name='All Data', index=True)
    
    df2.to_excel(writer, sheet_name='Max_Min Data', index=True)
    
    # Get the xlsxwriter workbook and worksheet objects
    workbook  = writer.book
    worksheet_1 = writer.sheets['All Data']
    worksheet_2 = writer.sheets['Max_Min Data']

    # Get the level 0 column labels (header) for the first DataFrame
    header_1 = df.columns.get_level_values(0)
    
    # Set the column width for each column based on the content in the first row
    for idx, value in enumerate(header_1):
        col_width = 13
        worksheet_1.set_column(idx, idx, col_width)
        worksheet_2.set_column(idx, idx, col_width)
    
# Export the DataFrame to a CSV file
df.to_csv(csv_output_file, index=False)

print(f'file {input_file} converted to {csv_output_file} & {excel_output_file}\nPlease check the folder.')

















