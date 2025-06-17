# CMDP Sample Results

# Sample Data Transfer Tool

This tool is designed to transfer sample data from a feature class to an Excel template. The tool retrieves data from a specified feature class and populates an Excel file template with the retrieved data. It also performs data formatting and manipulation based on predefined rules.

## Requirements

- ArcPy
- openpyxl
- Python 3.x

## Usage

1. Ensure all required Python packages are installed.
2. Modify the script to specify the paths to the feature class and Excel template file.
3. Run the script.

## Features

- Retrieves data from a specified feature class.
- Populates an Excel template with the retrieved data.
- Performs data formatting and manipulation based on predefined rules.
- Handles date formatting and parsing.
- Generates unique sample IDs.
- Updates columns based on specific conditions.

## Example

```python
import arcpy
import openpyxl
import datetime

# Feature class location
feature_class = r'\\portalserver\Production Projects\GISDBSERVER22.sde\HUD_LGIM.DBO.State_Lab_Samples'
excel_file = r"\\VSERVER22\ForEveryone\Nolan\CMDP_Sample_Result_Template - Excel.xlsx" 

# Open Excel file
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Define fields mapping between Excel and feature class
fields_in_server = ['OBJECTID', 'DATE_REC_LAB', 'ROTATION', ...]
fields_in_excel = [...]  # Define fields from the Excel template

# Create a dictionary to map Excel fields to feature class fields
field_mapping = dict(zip(fields_in_excel, fields_in_server))

# Clear existing data in the Excel sheet
...

# Use SearchCursor to retrieve data from the feature class
...

# Save the final modified Excel file after deleting rows
...

# Format date columns in the Excel sheet
...

# Generate sample IDs
...

# Update 'A/P*f' column based on 'Comment'
...

# Save the modified Excel file
...
