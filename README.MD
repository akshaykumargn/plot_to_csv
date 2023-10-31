# Read to CSV

This Python script is designed to convert a specific type of data file format (.plt) into more usable formats, namely Excel (xlsx) and CSV (csv). It performs data parsing, cleaning, and transformation to ensure that the data is structured correctly for further analysis. This README file provides an overview of the script's functionality, how to use it, and any specific considerations.
##Table of Contents

    Prerequisites
    Usage
    Input
    Output
    File Structure
    Running the Script
    Issues and Limitations

## Prerequisites

    Python 3.x
    Required Python packages are specified at the beginning of the script and can be installed using pip:

```bash
pip install pandas openpyxl
```

## Usage

This script is intended for users who have data in a .plt file format and need to convert it to a more accessible format. It performs the following steps:

    Parses the input .plt file.
    Identifies header information and data.
    Cleans and structures the data.
    Writes the data to both Excel (.xlsx) and CSV (.csv) files.

## Input

    The primary input is the .plt file containing the data you want to convert. You can specify the input file path in the script's input_file variable.

## Output

    The script generates two output files:
        An Excel file (.xlsx) named based on the input file with columns formatted for readability.
        A CSV file (.csv) with the same data as the Excel file.

## File Structure

    readtocsv.py: The Python script for converting .plt files to Excel and CSV.
    readtocsv_mux.py: Python script with more detailed representation of plot file data with multi index columns along with max and min forces to enable smooth data visualization in the second sheet.
    test_6.plt: An example .plt file provided for testing the script.
    test_6.csv: Example output file.
    README.MD: This documentation file.

## Running the Script

    Ensure that you have installed the required Python packages (pandas and openpyxl).
    Place your .plt file in the same directory as the script.
    Open the script in a code editor and set the input_file variable to the name of your .plt file.
    Run the script.

After execution, you will find the converted Excel (.xlsx) and CSV (.csv) files in the same directory as the script. The Excel file will have improved column formatting and alignment for readability.
Issues and Limitations

    The script assumes a specific structure for the .plt file, where header information is followed by data. If your data format differs significantly, you may need to modify the script.
    The script performs basic data cleaning and formatting. Further data analysis or specific formatting may be required based on your use case.
    Be cautious when using this script with large datasets, as it may consume a significant amount of memory, leading to potential performance issues.
