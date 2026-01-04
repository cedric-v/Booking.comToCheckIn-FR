#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Booking.com to CheckInFR CSV Converter

This script converts Booking.com reservation exports (CSV, XLS, or XLSX format) 
into CheckInFR import format.
"""

import csv
import sys
from datetime import datetime
from pathlib import Path

# Try to import openpyxl for XLSX support (optional)
try:
    import openpyxl
    XLSX_SUPPORT = True
except ImportError:
    XLSX_SUPPORT = False

# Try to import xlrd for XLS support (optional)
try:
    import xlrd
    XLS_SUPPORT = True
except ImportError:
    XLS_SUPPORT = False

# Configuration: Update these paths to match your files
# Source file: Booking.com export (CSV, XLS, or XLSX format)
source_file = "booking_export.xls"
# Template file: CheckInFR template CSV (defines output columns)
template_file = "checkinfr_template.csv"
# Output file: Generated CSV for CheckInFR import
output_file = "output_importation.csv"

def convert_date(date_str):
    """
    Convert date from YYYY-MM-DD format to DD-MM-YYYY format.
    Removes HTML entities like &nbsp; from the date string.
    Also handles Excel date serial numbers if passed as float.
    
    Args:
        date_str: Date string in YYYY-MM-DD format, or float (Excel serial number)
        
    Returns:
        Formatted date string in DD-MM-YYYY format, or empty string if conversion fails
    """
    if not date_str:
        return ""
    
    try:
        # Handle Excel date serial numbers (float)
        if isinstance(date_str, (int, float)) and date_str > 0:
            # This is likely an Excel date serial number
            # xlrd should have already converted it, but just in case
            return ""
        
        # Convert to string and remove HTML entities (&nbsp; or others)
        clean_date = str(date_str).replace("&nbsp;", "").strip()
        
        # Try YYYY-MM-DD format (most common from Booking.com)
        if len(clean_date) >= 10:
            return datetime.strptime(clean_date[:10], "%Y-%m-%d").strftime("%d-%m-%Y")
        
        return ""
    except Exception:
        return ""

def read_csv_file(file_path):
    """Read CSV file and return list of dictionaries."""
    rows = []
    with open(file_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter=';')
        for row in reader:
            rows.append(row)
    return rows

def read_xlsx_file(file_path):
    """Read XLSX file and return list of dictionaries."""
    if not XLSX_SUPPORT:
        print("Error: XLSX support requires 'openpyxl' package.")
        print("Install it with: pip install openpyxl")
        sys.exit(1)
    
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    
    # Read header row
    headers = []
    for cell in sheet[1]:
        headers.append(cell.value if cell.value else "")
    
    # Read data rows
    rows = []
    for row in sheet.iter_rows(min_row=2, values_only=False):
        row_dict = {}
        for idx, cell in enumerate(row):
            header = headers[idx] if idx < len(headers) else f"Column_{idx+1}"
            row_dict[header] = str(cell.value) if cell.value is not None else ""
        rows.append(row_dict)
    
    return rows

def read_xls_file(file_path):
    """Read XLS file (old Excel format) and return list of dictionaries."""
    if not XLS_SUPPORT:
        print("Error: XLS support requires 'xlrd' package.")
        print("Install it with: pip install xlrd")
        sys.exit(1)
    
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)
    
    # Read header row (first row)
    headers = []
    for col_idx in range(sheet.ncols):
        cell_value = sheet.cell_value(0, col_idx)
        headers.append(str(cell_value) if cell_value else "")
    
    # Read data rows
    rows = []
    for row_idx in range(1, sheet.nrows):
        row_dict = {}
        for col_idx in range(sheet.ncols):
            header = headers[col_idx] if col_idx < len(headers) else f"Column_{col_idx+1}"
            cell_value = sheet.cell_value(row_idx, col_idx)
            # Convert dates to string format
            if isinstance(cell_value, float) and cell_value > 0:
                # Check if it might be a date (Excel date serial number)
                try:
                    date_tuple = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
                    if date_tuple[0] > 1900:  # Valid year
                        row_dict[header] = f"{date_tuple[0]}-{date_tuple[1]:02d}-{date_tuple[2]:02d}"
                    else:
                        row_dict[header] = str(cell_value)
                except:
                    row_dict[header] = str(cell_value)
            else:
                row_dict[header] = str(cell_value) if cell_value else ""
        rows.append(row_dict)
    
    return rows

def read_source_file(file_path):
    """Read source file (CSV, XLS, or XLSX) and return list of dictionaries."""
    file_ext = Path(file_path).suffix.lower()
    
    if file_ext == '.xlsx':
        if not XLSX_SUPPORT:
            print("Error: XLSX support requires 'openpyxl' package.")
            print("Install it with: pip install openpyxl")
            print("Or convert your XLSX file to CSV format first.")
            sys.exit(1)
        return read_xlsx_file(file_path)
    elif file_ext == '.xls':
        if not XLS_SUPPORT:
            print("Error: XLS support requires 'xlrd' package.")
            print("Install it with: pip install xlrd")
            print("Or convert your XLS file to CSV format first.")
            sys.exit(1)
        return read_xls_file(file_path)
    elif file_ext in ['.csv', '']:
        return read_csv_file(file_path)
    else:
        print(f"Error: Unsupported file format: {file_ext}")
        print("Supported formats: .csv, .xls, .xlsx")
        sys.exit(1)

def main():
    """Main conversion function."""
    # Use local variable for template file (may be updated if example template is used)
    template_file_path = template_file
    
    # Check if source file exists
    if not Path(source_file).exists():
        print(f"Error: Source file not found: {source_file}")
        print("Please update the 'source_file' variable in the script.")
        sys.exit(1)
    
    # Check if template file exists, try example template if not found
    if not Path(template_file_path).exists():
        # Try to use example template from project directory
        script_dir = Path(__file__).parent
        example_template = script_dir / "template_checkinfr_example.csv"
        if example_template.exists():
            print(f"⚠️  Template file not found: {template_file_path}")
            print(f"   Using example template: {example_template}")
            template_file_path = str(example_template)
        else:
            print(f"Error: Template file not found: {template_file_path}")
            print("Please update the 'template_file' variable in the script.")
            print("Or create a template CSV file with the required column headers.")
            sys.exit(1)
    
    # Detect file format
    file_ext = Path(source_file).suffix.lower()
    if file_ext == '.xlsx' and not XLSX_SUPPORT:
        print("⚠️  Warning: XLSX file detected but 'openpyxl' is not installed.")
        print("   Install it with: pip install openpyxl")
        print("   Or convert your XLSX file to CSV format first.")
        sys.exit(1)
    elif file_ext == '.xls' and not XLS_SUPPORT:
        print("⚠️  Warning: XLS file detected but 'xlrd' is not installed.")
        print("   Install it with: pip install xlrd")
        print("   Or convert your XLS file to CSV format first.")
        sys.exit(1)
    
    # Read source file and process reservations
    reservations = []
    try:
        source_rows = read_source_file(source_file)
        
        for row in source_rows:
            # Map Booking.com columns (English) to CheckInFR fields
            # Try both English and French column names for compatibility
            def get_value(key, alt_keys=None):
                """Get value from row, trying multiple key names."""
                value = row.get(key, "")
                if not value and alt_keys:
                    for alt_key in alt_keys:
                        value = row.get(alt_key, "")
                        if value:
                            break
                
                # Convert to string and clean
                if isinstance(value, float):
                    # Remove .0 from whole numbers
                    if value == int(value):
                        return str(int(value))
                    return str(value)
                elif isinstance(value, int):
                    return str(value)
                else:
                    # Handle string values that might be numbers
                    value_str = str(value).strip() if value else ""
                    # Check if string represents a whole number with .0 (e.g., "2.0")
                    try:
                        float_val = float(value_str)
                        if float_val == int(float_val):
                            return str(int(float_val))
                        return value_str
                    except (ValueError, TypeError):
                        return value_str
            
            book_number = get_value("Book number", ["Numéro de réservation"])
            check_in = get_value("Check-in", ["Arrivée&nbsp;", "Arrivée"])
            check_out = get_value("Check-out", ["Départ&nbsp;", "Départ"])
            guest_name = get_value("Guest name(s)", ["Nom du client"])
            phone = get_value("Phone number", ["Numéro de téléphone"])
            travel_purpose = get_value("Travel purpose", ["Motif du voyage"])
            address = get_value("Address", ["Adresse"])
            adults = get_value("Adults", ["Personnes"])
            children = get_value("Children", ["Enfants"])
            
            # Clean and structure data
            name_parts = guest_name.split() if guest_name else []
            
            reservations.append({
                "Numéro de référence externe": book_number,
                "Date d'arrivée": convert_date(check_in),
                "Date de départ": convert_date(check_out),
                "Nom": name_parts[0] if name_parts else "",
                "Prénom": " ".join(name_parts[1:]) if len(name_parts) > 1 else "",
                "Téléphone privé": phone,
                "Segmentation": travel_purpose,
                "Adresse (Rue)": address,
                "Nombre total d'adultes": adults,
                "Nombre total d'enfants": children
            })
    except Exception as e:
        print(f"Error reading source file: {e}")
        sys.exit(1)
    
    # Read template file to get column structure
    columns = []
    try:
        with open(template_file_path, "r", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=';')
            columns = next(reader)  # First line contains column names
    except Exception as e:
        print(f"Error reading template file: {e}")
        sys.exit(1)
    
    # Write output file
    try:
        with open(output_file, "w", encoding="utf-8", newline='') as f:
            writer = csv.DictWriter(f, fieldnames=columns, delimiter=';')
            writer.writeheader()
            for reservation in reservations:
                # Filter only fields that exist in 'columns'
                filtered_reservation = {key: reservation.get(key, "") for key in columns}
                writer.writerow(filtered_reservation)
        
        print(f"✓ Successfully converted {len(reservations)} reservations")
        print(f"✓ Output file generated: {output_file}")
    except Exception as e:
        print(f"Error writing output file: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()