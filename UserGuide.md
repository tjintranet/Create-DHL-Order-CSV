# DHL Report Generator - User Guide

## Table of Contents
1. [Getting Started](#getting-started)
2. [Processing Main Order File](#processing-main-order-file)
3. [Updating Email Addresses](#updating-email-addresses)
4. [Working with the Data Table](#working-with-the-data-table)
5. [Exporting the DHL Report](#exporting-the-dhl-report)
6. [Troubleshooting](#troubleshooting)

## Getting Started

### System Requirements
- Modern web browser (Chrome, Firefox, Safari, or Edge)
- JavaScript enabled
- Internet connection for loading required libraries

### Before You Begin
Ensure you have:
- Your main order Excel file
- Any email update files (if required)
- All column headers match the expected format

## Processing Main Order File

### Step 1: Initial Upload
1. Open the DHL Report Generator in your web browser
2. Locate the "Upload Main Excel File" section
3. Click the file input or drag and drop your Excel file
4. Wait for the "File processed successfully!" message

### Step 2: Verify Data
After upload, check that:
- All order details appear correctly in the table
- Dates are in the correct format (DD-MM-YYYY)
- Australian orders have correct state codes
- All required columns contain the expected data

### Notes About Australian Orders
- Orders with country code 'AU' will automatically have their state codes looked up
- The lookup is based on the postcode in column H
- State codes are automatically inserted in column I
- Verify that state codes are correct for all Australian orders

## Updating Email Addresses

### Step 1: Prepare Email Update File
Ensure your email update file:
- Contains two columns
- Column A has order numbers matching the main file
- Column B has the new email addresses

### Step 2: Upload Email Updates
1. Wait for the main file to finish processing
2. Look for "Upload Email Update File" to become enabled
3. Click to upload your email update file
4. Watch for the success message showing number of updates

### Step 3: Verify Updates
- Check that email addresses were updated correctly
- Verify that other order data remained unchanged
- Multiple email update files can be processed sequentially

## Working with the Data Table

### Viewing Data
- Use horizontal scroll to view all columns
- The action column (delete button) stays fixed on the left
- Smaller font size helps view more data at once

### Deleting Rows
1. Locate the row you want to remove
2. Click the trash can icon in the leftmost column
3. The row will be immediately removed
4. This action cannot be undone

### Data Display
Columns shown in the table:
- Order Number (A)
- Date (B)
- To Name (C)
- Destination Building (D)
- Destination Street (E)
- Destination Suburb (F)
- Destination City (G)
- Destination Postcode (H)
- Destination State (I)
- Destination Email (K)
- Destination Phone (L)
- Reference (R)
- Country Code (X)

## Exporting the DHL Report

### Step 1: Prepare for Export
Before exporting:
- Verify all data is correct
- Remove any unwanted rows
- Check Australian state codes
- Confirm email addresses are up to date

### Step 2: Export Process
1. Click the "Export CSV" button in the top right
2. The file will automatically download
3. Filename format: DD-MM-YYYY_DHL Report.csv
   - Date taken from first order in the file

### Step 3: Email to despatch
email csv export to despatch01@tjbooks.co.uk

## Troubleshooting

### Common Issues and Solutions

#### File Won't Upload
- Check file format (.xlsx or .xls only)
- Ensure file isn't corrupted
- Try saving as a new Excel file

#### Missing Email Updates
- Verify order numbers match exactly
- Check for extra spaces in order numbers
- Ensure email file has correct column structure

#### Incorrect State Codes
- Verify postcode is correct
- Check country code is 'AU'
- Ensure postcode exists in database

#### Export Problems
- Clear browser cache
- Try a different browser
- Check for special characters in data

### Getting Help
If you encounter persistent issues:
1. Clear your browser cache
2. Restart your browser
3. Try a different supported browser
4. Check file formats and data structure
5. Contact technical support if problems persist