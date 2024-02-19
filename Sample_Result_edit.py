
import arcpy
import openpyxl

# Feature class location
feature_class = r'C:\Users\Nolan\Documents\ArcGIS\ArcGIS Pro Projects\DateDelete\DateCleanUpDelete\DateCleanUpDelete.gdb\State_Lab_Samples'
excel_file = r"C:\Users\Nolan\Documents\ExcelSheets\CMDP_Sample_Result_Template - Copy3.xlsx"

fields_in_server = [ 'OBJECTID', 'ROTATION', 'INSTALLDATE', 'ELEVATION', 'NAME', 'SPDATE', 'SAMPLEVAL', 'THRESHHIT', 'DISINFRULE', 'ENABLED', 'ACTIVEFLAG', 'OWNEDBY',
                    'MAINTBY', 'LASTUPDATE', 'LASTEDITOR', 'LOCATION_COMMENTS', 'SECTION', 'SAMPLE_NUMBER', 'SAMPLE_DATE', 'SAMPLE_TIME', 'SAMPLE_FIXTURE', 'EMPLOYEE', 
                    'STATIONID', 'MAP_NUMBER', 'ADDRESS', 'SAMPLE_READING', 'SAMPLE_TYPE', 'TESTED', 'DATE_REC_LAB', 'TIME_REC_LAB', 'REC_LAB_BY', 'ANALYSIS_DATE', 
                    'ANALYSIS_TIME', 'ANALYST', 'TOTAL_COLIFORM_RESULTS']

fields_in_excel = [ 'Sample ID*', 'Sample Received Date f', 'WS ID*', 'Facility ID*', 'Sampling Point ID', 'Sampling Location', 'Collection Date*f', 'Collection Time (24H) f', 'Sample Type*f',
                   'Sample Volume (ML) f', 'Repeat Location', 'Original Sample ID +', 'Original Reporting Lab.ID', 'Original Collection Date', 'Comment', 'Sample Collector Name', 
                   'Analyte*f [Code - Name]', 'A/P*f', 'Count', 'Units +', 'Volume (ML) +', 'Interference', 'Volume Assayed (ML) f', 'Method f', 'Analysis Start Date f', 
                   'Analysis Start Time f', 'Analysis Completed Date', 'Analysis Completed Time', 'Analyst Name', 'Analyzing Lab ID', 'Source Type', 'Comment',
                   'Parameter* [Code - Name]', 'Result*', 'Result UOM*', 'Method', 'Analyst Name', 'Comment']

# Create a dictionary to map Excel fields to feature class fields
field_mapping = dict(zip(fields_in_excel, fields_in_server))

# Open Excel file
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Clear existing data in the Excel sheet
for row in sheet.iter_rows(min_row=9, max_row=sheet.max_row, max_col=sheet.max_column):
    for cell in row:
        cell.value = None

# Use SearchCursor to retrieve data from the feature class
with arcpy.da.SearchCursor(feature_class, fields_in_server) as cursor:
    # Find the first empty row in the Excel sheet
    row_index = 9  # Assuming the data starts from the ninth row in Excel
    while sheet.cell(row=row_index, column=1).value is not None:
        row_index += 1

    # Write data to Excel
    for row in cursor:
        # Create a dictionary to map feature class field names to values
        data_dict = dict(zip(fields_in_server, row))

        # Write data to the next available row in Excel, treating all values as text
        for excel_field, server_field in field_mapping.items():
            sheet.cell(row=row_index, column=fields_in_excel.index(excel_field) + 1, value=str(data_dict[server_field]))

        row_index += 1

# Save the modified Excel file
workbook.save(excel_file)