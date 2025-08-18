import openpyxl
from openpyxl import load_workbook
import io
import tempfile
import os
from datetime import datetime
from .debug_logger import debug_logger
import zipfile

def process_county_files(main_workbook_file, county_files):
    """
    Process county Excel files and update the main workbook's Raw sheet.
    
    Field Mapping (from analysis):
    - Column C (Medicare Enrollment) → Raw Column E (Medicare Beneficiaries)
    - Column E (Resident Deaths) → Raw Column G (Medicare Deaths)  
    - Column I (Patients Served) → Raw Column K (Hospice Unduplicated Beneficiaries)
    - Column G (Hospice Deaths) → Raw Column I (Hospice Deaths)
    """
    try:
        # Load main workbook
        main_wb = load_workbook(main_workbook_file, data_only=False)
        
        # Find the Raw sheet
        if 'Raw' not in main_wb.sheetnames:
            raise ValueError("Main workbook does not contain a 'Raw' sheet")
        
        raw_sheet = main_wb['Raw']
        
        # Find the next empty row in Raw sheet
        next_row = find_next_empty_row(raw_sheet)
        
        # Process each county file
        for county_file in county_files:
            county_data = extract_county_data(county_file)
            if county_data:
                # Add data to Raw sheet
                next_row = add_county_data_to_raw(raw_sheet, county_data, next_row)
        
        # Save to memory buffer
        output_buffer = io.BytesIO()
        main_wb.save(output_buffer)
        output_buffer.seek(0)
        
        return output_buffer
        
    except Exception as e:
        print(f"Error processing files: {e}")
        return None

def extract_county_data(county_file):
    """
    Extract data from County Trend sheet, rows 10+ (skipping header at row 9), columns B,C,E,G,I
    Stops at first empty row to extract only the PRIMARY table (first continuous data section)
    """
    debug_logger.reset_trace()
    
    try:
        # First, analyze the file
        file_info = debug_logger.log_file_info(county_file)
        
        # Check if file is a valid Excel file before attempting to load
        if not file_info.get("is_zip_file") and file_info.get("file_type") != "Old Excel format (XLS)":
            error_msg = f"File {county_file.filename} is not a valid Excel file: {file_info.get('file_type', 'Unknown format')}"
            print(error_msg)
            debug_logger.logger.error(error_msg)
            return None
        
        # Attempt to load the workbook
        debug_logger.logger.info(f"Loading workbook: {county_file.filename}")
        county_wb = load_workbook(county_file, data_only=True)
        
        # Log sheet detection
        sheet_names = county_wb.sheetnames
        sheet_info = debug_logger.log_sheet_detection(county_wb, sheet_names)
        
        # Find County Trend sheet
        trend_sheet = None
        for sheet_name in sheet_names:
            if 'trend' in sheet_name.lower() and 'county' in sheet_name.lower():
                trend_sheet = county_wb[sheet_name]
                debug_logger.logger.info(f"Selected sheet: {sheet_name}")
                break
        
        if not trend_sheet:
            error_msg = f"No 'County Trend' sheet found in {county_file.filename}"
            print(error_msg)
            debug_logger.logger.warning(error_msg)
            return None
        
        # Extract county name from filename
        county_name = county_file.filename.replace('.xlsx', '').replace('.xls', '')
        debug_logger.logger.info(f"Processing county: {county_name}")
        
        # Log sheet dimensions
        debug_logger.logger.info(f"Sheet dimensions - Max row: {trend_sheet.max_row}, Max column: {trend_sheet.max_column}")
        
        # Extract data from rows 10+ (row 9 is headers, data starts at row 10)
        extracted_data = []
        rows_checked = 0
        stop_reason = "Unknown"
        
        for row in range(10, trend_sheet.max_row + 1):
            rows_checked += 1
            
            year_cell = trend_sheet[f'B{row}']
            medicare_enrollment_cell = trend_sheet[f'C{row}']
            resident_deaths_cell = trend_sheet[f'E{row}']
            hospice_deaths_cell = trend_sheet[f'G{row}']
            patients_served_cell = trend_sheet[f'I{row}']
            
            # Log the raw values for debugging
            debug_logger.logger.debug(f"Row {row} raw values - Year: {year_cell.value}, Medicare: {medicare_enrollment_cell.value}")
            
            # Stop at first empty row (end of PRIMARY table)
            # Check if year or medicare enrollment is empty/None
            if (year_cell.value is None or year_cell.value == '' or 
                medicare_enrollment_cell.value is None or medicare_enrollment_cell.value == ''):
                stop_reason = f"Empty row detected at row {row}"
                debug_logger.log_row_extraction(row, year_cell.value, medicare_enrollment_cell.value, 
                                               "stop", "Empty year or medicare enrollment - end of PRIMARY table")
                break
            
            # Check if we have valid numeric data in this row
            if (isinstance(year_cell.value, (int, float)) and 
                isinstance(medicare_enrollment_cell.value, (int, float))):
                
                row_data = {
                    'county': county_name,
                    'state': 'NC',
                    'year': year_cell.value,
                    'medicare_enrollment': medicare_enrollment_cell.value,
                    'resident_deaths': resident_deaths_cell.value if resident_deaths_cell.value else 0,
                    'hospice_deaths': hospice_deaths_cell.value if hospice_deaths_cell.value else 0,
                    'patients_served': patients_served_cell.value if patients_served_cell.value else 0
                }
                extracted_data.append(row_data)
                debug_logger.log_row_extraction(row, year_cell.value, medicare_enrollment_cell.value,
                                               "extract", "Valid numeric data")
            else:
                debug_logger.log_row_extraction(row, year_cell.value, medicare_enrollment_cell.value,
                                               "skip", f"Non-numeric data - Year type: {type(year_cell.value)}, Medicare type: {type(medicare_enrollment_cell.value)}")
        
        # Log extraction summary
        if rows_checked >= (trend_sheet.max_row - 9):
            stop_reason = "Reached end of sheet"
            
        summary = debug_logger.log_extraction_summary(county_name, rows_checked, len(extracted_data), stop_reason)
        
        # Save trace for analysis
        if len(extracted_data) != 15:
            # If we didn't get exactly 15 rows, save trace for debugging
            trace_file = debug_logger.save_trace_to_file(f"{county_name}_unexpected_count")
            debug_logger.logger.warning(f"Unexpected row count ({len(extracted_data)} instead of 15). Trace saved to {trace_file}")
        
        return extracted_data
        
    except zipfile.BadZipFile as e:
        error_msg = f"Bad zip file error for {county_file.filename}: {e}"
        print(error_msg)
        debug_logger.logger.error(error_msg)
        debug_logger.save_trace_to_file(f"{county_file.filename}_bad_zip")
        return None
    except Exception as e:
        error_msg = f"Error extracting data from {county_file.filename}: {e}"
        print(error_msg)
        debug_logger.logger.error(error_msg, exc_info=True)
        debug_logger.save_trace_to_file(f"{county_file.filename}_error")
        return None

def find_next_empty_row(sheet):
    """
    Find the next empty row in the Raw sheet
    """
    # Start from row 2 (assuming row 1 is headers)
    row = 2
    while sheet[f'A{row}'].value is not None:
        row += 1
    return row

def add_county_data_to_raw(raw_sheet, county_data, start_row):
    """
    Add county data to Raw sheet with proper field mapping
    
    Field Mapping:
    - Column A: County
    - Column B: State 
    - Column C: Year
    - Column D: Key (will be auto-generated by CONCAT formula)
    - Column E: Medicare Beneficiaries (from Medicare Enrollment)
    - Column G: Medicare Deaths (from Resident Deaths)
    - Column I: Hospice Deaths (from Hospice Deaths)
    - Column K: Hospice Unduplicated Beneficiaries (from Patients Served)
    """
    current_row = start_row
    
    for data in county_data:
        # Map data to Raw sheet columns
        raw_sheet[f'A{current_row}'] = data['county']
        raw_sheet[f'B{current_row}'] = data['state']
        raw_sheet[f'C{current_row}'] = data['year']
        
        # Column D: Key field - copy formula from previous row if it exists
        if current_row > 2:  # If not the first data row
            prev_formula = raw_sheet[f'D{current_row-1}'].value
            if prev_formula and isinstance(prev_formula, str) and prev_formula.startswith('='):
                # Update row reference in formula
                updated_formula = prev_formula.replace(str(current_row-1), str(current_row))
                raw_sheet[f'D{current_row}'] = updated_formula
            else:
                # Create new CONCAT formula
                raw_sheet[f'D{current_row}'] = f'=CONCAT(A{current_row},"-",B{current_row},"-",C{current_row})'
        else:
            # First data row, create CONCAT formula
            raw_sheet[f'D{current_row}'] = f'=CONCAT(A{current_row},"-",B{current_row},"-",C{current_row})'
        
        # Field mapping based on analysis
        raw_sheet[f'E{current_row}'] = data['medicare_enrollment']  # Medicare Beneficiaries
        raw_sheet[f'G{current_row}'] = data['resident_deaths']      # Medicare Deaths
        raw_sheet[f'I{current_row}'] = data['hospice_deaths']       # Hospice Deaths
        raw_sheet[f'K{current_row}'] = data['patients_served']      # Hospice Unduplicated Beneficiaries
        
        current_row += 1
    
    return current_row
