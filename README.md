# Booking.com to CheckInFR Converter

A Python script to convert Booking.com reservation data (CSV, XLS, or XLSX format) into CheckInFR import format.

## Description

This script processes reservation data exported from Booking.com (in CSV, XLS, or XLSX format) and transforms it into the CSV format required by [CheckInFR](https://checkin-fr.ch/index.php), a Swiss hotel management system from Fribourg. It handles date format conversion, field mapping, and data cleaning.

## Features

- Converts Booking.com CSV exports to CheckInFR import format
- Handles date format conversion (YYYY-MM-DD to DD-MM-YYYY)
- Cleans HTML entities (e.g., `&nbsp;`) from date fields
- Maps Booking.com fields to CheckInFR required fields:
  - External reference number
  - Arrival and departure dates
  - Guest name (split into first and last name)
  - Phone number
  - Travel purpose (segmentation)
  - Address
  - Number of adults
  - Number of children
- Uses a template file to ensure correct column structure

## Requirements

- Python 3.x
- **For CSV files**: Standard library only (no external dependencies required)
- **For XLSX files**: `openpyxl` package (optional, install with `pip install openpyxl`)
- **For XLS files**: `xlrd` package (optional, install with `pip install xlrd`)

## Usage

1. **Export data from Booking.com:**
   
   To get your reservation data from Booking.com:
   - Log in to your Booking.com extranet
   - Go to **"Reservations"**
   - Select the desired date range
   - In **"Reservation status"**, select **"OK"**
   - Click the **"Show"** button
   - Once the list is displayed, click **"Download reservations statement"**
   - Download the generated XLS file
   
   The file will typically be named something like: `Check-in YYYY-MM-DD to YYYY-MM-DD.xls`

2. **Prepare your files:**
   - Have your Booking.com export file (CSV, XLS, or XLSX format)
   - Have a CheckInFR template CSV file ready (this defines the output column structure)
   
   **Note**: If using XLSX or XLS files, install the optional dependencies:
   ```bash
   pip install openpyxl xlrd
   ```
   - `openpyxl` for .xlsx files (Excel 2007+)
   - `xlrd` for .xls files (older Excel format, as exported by Booking.com)

3. **Update the script:**
   Edit the file paths in `booking_to_checkinfr_converter.py`:
   ```python
   source_file = "path/to/your/booking.com/export.xls"  # or .csv, .xlsx
   template_file = "path/to/checkinfr/template.csv"
   output_file = "path/to/output.csv"
   ```
   
   The script automatically detects the file format (CSV, XLS, or XLSX) based on the file extension.

4. **Run the script:**
   ```bash
   python booking_to_checkinfr_converter.py
   ```

5. **Import into CheckInFR:**
   
   To import the generated CSV file into CheckInFR:
   - Log in to your CheckInFR account
   - Go to **"Clients"**
   - Click on **"Importer"** (Import)
   - Upload/drop the generated CSV output file
   - Map the columns to match CheckInFR's expected fields
   - Follow the importation steps as indicated
   - Complete the import process
   
   The generated file will be in the location specified by `output_file` in the script (typically `output_importation.csv`).

## Field Mapping

| Booking.com Field | CheckInFR Field |
|-------------------|----------------|
| Book number | Numéro de référence externe |
| Check-in | Date d'arrivée |
| Check-out | Date de départ |
| Guest name(s) (first word) | Nom |
| Guest name(s) (remaining words) | Prénom |
| Phone number | Téléphone privé |
| Travel purpose | Segmentation |
| Address | Adresse (Rue) |
| Adults | Nombre total d'adultes |
| Children | Nombre total d'enfants |

**Note**: The script automatically maps both English column names (as exported by Booking.com) and French column names for compatibility.

## Date Format

The script automatically converts dates from `YYYY-MM-DD` format (Booking.com) to `DD-MM-YYYY` format (CheckInFR).

## Notes

- The script uses semicolon (`;`) as the CSV delimiter
- HTML entities in date fields are automatically cleaned
- Only fields present in the template file will be included in the output
- The script handles missing or empty fields gracefully

## Installation & Setup

1. Clone this repository:
   ```bash
   git clone https://github.com/cedric-v/Booking.comToCheckIn-FR.git
   cd Booking.comToCheckIn-FR
   ```

2. **Optional**: If you plan to use XLSX or XLS files, install the optional dependencies:
   ```bash
   pip install openpyxl xlrd
   ```
   
   For CSV files only, no additional dependencies are needed - uses Python standard library only.

## File Structure

```
Booking.comToCheckIn-FR/
├── booking_to_checkinfr_converter.py             # Main conversion script
├── README.md                                      # This file
├── .gitignore                                     # Git ignore rules
├── requirements.txt                               # Dependencies (none required)
├── setup_git.sh                                   # Git setup script
├── GIT_SETUP.md                                   # Git setup instructions
└── LISEZMOI - script décompte nuiteés.docx      # Original French documentation
```

## License

This project is licensed under the BSD 3-Clause License - see the [LICENSE](LICENSE) file for details.

## About CheckInFR

[CheckInFR](https://checkin-fr.ch/index.php) is a hotel management system from Fribourg, Switzerland (FR = Fribourg, not France).

## Author

Created for converting Booking.com reservation data to CheckInFR format.

