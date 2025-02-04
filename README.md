# DHL Report Generator

A web-based application for processing and formatting Excel files into DHL-compatible reports with additional features for email updates and Australian state code lookups.

## Features

### Main Functionality
- Upload and process Excel files with order data
- Display and edit data in a responsive table format
- Export data to CSV in DHL report format
- Delete individual rows as needed
- Auto-formatting of dates in DD-MM-YYYY format
- Automatic generation of output filename based on order date

### Australian State Lookup
- Automatically detects orders with country code 'AU'
- Looks up and inserts correct state codes based on postcodes
- Uses Australian postcode database for accurate state assignment

### Email Updates
- Secondary file upload for email address updates
- Matches order numbers to update corresponding email addresses
- Supports multiple email update files
- Shows success message with number of updates made
- Maintains all other order data while updating emails

### User Interface
- Clean, responsive Bootstrap-based design
- Compact table view with fixed action column
- Clear status messages with automatic fade-out
- Disabled buttons when no data is loaded
- Intuitive delete functionality for individual rows

## File Requirements

### Main Excel File
- Must contain order data in specified columns
- Column mapping:
  - A: Order Number
  - B: Date
  - C: To Name
  - D: Destination Building
  - E: Destination Street
  - F: Destination Suburb
  - G: Destination City
  - H: Destination Postcode
  - I: Destination State
  - K: Destination Email
  - L: Destination Phone
  - R: Reference
  - X: Country Code

### Email Update Files
- Must contain two columns:
  - Column A: Order Number (matching main file)
  - Column B: Email Address

### Australian Postcode File
- JSON file containing Australian postcode to state mappings
- Required for automatic state code lookup
- Format: `[{"postcode": number, "State": "string"}]`

## Setup

1. Clone the repository
2. Ensure all files are in the same directory:
   - index.html
   - script.js
   - style.css
   - AUSPoscodes.json

## Usage

1. Open index.html in a web browser
2. Upload main Excel file using the "Upload Main Excel File" button
3. Wait for processing to complete
4. (Optional) Upload email update file if needed
5. Review and edit data in the table as needed
6. Use the delete buttons to remove unwanted rows
7. Click "Export CSV" to download the formatted DHL report

## Output Format

The generated CSV file will:
- Be named using the date from the first order (DD-MM-YYYY_DHL Report.csv)
- Include all 51 required DHL report columns
- Contain mapped data from the input Excel
- Include fixed values for specific fields:
  - Carrier: "DHL"
  - Carrier Product Unit Type: "Parcel"
  - Declared Value Currency: "AUD"
  - Dangerous Goods: "0"
  - DDP: "N"
  - Qty: "1"

## Technical Details

### Dependencies
- Bootstrap 5.3.2
- Font Awesome 6.4.0
- SheetJS 0.18.5
- PapaParse 5.4.1

### Browser Compatibility
- Works with modern web browsers (Chrome, Firefox, Safari, Edge)
- Requires JavaScript enabled
- Supports Excel files (.xlsx, .xls)