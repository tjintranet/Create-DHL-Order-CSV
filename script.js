// Function to handle Weight value conversion
function processWeightValue(value) {
    // If value is already a number, return it directly
    if (typeof value === 'number') {
        return value;
    }
    
    // If value is undefined or empty, return default
    if (!value) {
        return 0.0;
    }
    
    // Try to parse the string value
    const parsed = parseFloat(value);
    
    // If parsed value is not a number, return default
    if (isNaN(parsed)) {
        return 0.0;
    }
    
    return parsed;
}
    // Global variables
let ausPostcodes = null;
let currentData = [];
let columnHeaders = {};

// Load Australian postcodes data
fetch('AUSPoscodes.json')
    .then(response => response.json())
    .then(data => {
        ausPostcodes = data;
        console.log('Loaded AUS postcodes:', ausPostcodes.length);
    })
    .catch(error => console.error('Error loading postcodes:', error));

// File input handlers
document.getElementById('excelFile').addEventListener('change', handleMainFileSelect);
document.getElementById('emailUpdateFile').addEventListener('change', handleEmailUpdateFile);

function handleMainFileSelect(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    document.getElementById('processingStatus').style.display = 'block';
    document.getElementById('processingStatus').textContent = 'Processing file...';

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array', raw: true});
        
        // Get first sheet
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Get headers first
        const range = XLSX.utils.decode_range(firstSheet['!ref']);
        const headers = {};
        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'R', 'X'];
        
        cols.forEach(col => {
            const cell = firstSheet[`${col}1`];
            if (cell) {
                headers[col] = cell.v;
            } else {
                headers[col] = `Column ${col}`;
            }
        });
        columnHeaders = headers;
        
        // Update table headers
        updateTableHeaders(headers);
        
        // Convert specific columns to array of objects
        const rawData = XLSX.utils.sheet_to_json(firstSheet, {header: 'A'});
        
        // Process the data to keep only specific columns
        const processedData = rawData.slice(1).map(row => ({
            A: row.A || '',
            B: row.B || '',
            C: row.C || '',
            D: row.D || '',
            E: row.E || '',
            F: row.F || '',
            G: row.G || '',
            H: row.H || '',
            I: row.I || '',
            J: row.J || '',
            K: row.K || '',
            L: row.L || '',
            P: row.P || '',
            R: row.R || '',
            X: row.X || ''
        }));

        // Debug log for Weight values
            console.log('Weight values from Excel:', processedData.map(row => row.P));

        processData(processedData);
        
        // Enable email update file input
        document.getElementById('emailUpdateFile').disabled = false;
    };

    reader.readAsArrayBuffer(file);
}

function handleEmailUpdateFile(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    document.getElementById('processingStatus').style.display = 'block';
    document.getElementById('processingStatus').textContent = 'Processing email updates...';
    document.getElementById('processingStatus').style.opacity = '1';

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array', raw: true});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const emailData = XLSX.utils.sheet_to_json(firstSheet, {header: ['orderNumber', 'email']});

        // Remove header row if present
        if (emailData.length > 0 && emailData[0].orderNumber === 'Order Number') {
            emailData.shift();
        }

        // Update emails in main data
        let updatesCount = 0;
        currentData.forEach(row => {
            const matchingEmail = emailData.find(emailRow => 
                emailRow.orderNumber && row.A && 
                emailRow.orderNumber.toString() === row.A.toString()
            );
            if (matchingEmail && matchingEmail.email) {
                row.K = matchingEmail.email;
                updatesCount++;
            }
        });

        // Display updated data
        displayData(currentData);

        // Show success message
        const statusElement = document.getElementById('processingStatus');
        statusElement.textContent = `Updated ${updatesCount} email addresses successfully!`;
        statusElement.style.opacity = '1';
        
        // Fade out message after 3 seconds
        setTimeout(() => {
            statusElement.style.opacity = '0';
            setTimeout(() => {
                statusElement.style.display = 'none';
            }, 500);
        }, 3000);
    };

    reader.readAsArrayBuffer(file);
}

function lookupAustralianState(postcode) {
    if (!ausPostcodes || !postcode) return '';
    
    // Convert postcode to number for comparison
    const postcodeNum = parseInt(postcode);
    if (isNaN(postcodeNum)) return '';
    
    // Find matching postcode
    const match = ausPostcodes.find(p => p.postcode === postcodeNum);
    if (match && match['State\r']) {
        // Remove carriage return and trim
        return match['State\r'].replace('\r', '').trim();
    }
    return '';
}

// Helper function to parse and format date consistently
function formatDateForFilename(dateValue) {
    // If it's already in DD-MM-YYYY format, return as is
    if (typeof dateValue === 'string') {
        const parts = dateValue.split(/[-/]/);
        if (parts.length === 3) {
            // If the first part is a valid day (1-31)
            const day = parseInt(parts[0]);
            if (day >= 1 && day <= 31) {
                return `${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}-${parts[2]}`;
            }
        }
    }
    
    // Try parsing as date object
    const date = new Date(dateValue);
    if (!isNaN(date.getTime())) {
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}-${month}-${year}`;
    }
    
    return null;
}

function processData(data) {
    currentData = data;
    
    // Process each row
    currentData.forEach(row => {
        // Check if country code is AU
        if (row.X === 'AU') {
            const postcode = row.H;
            const state = lookupAustralianState(postcode);
            if (state) {
                row.I = state; // Update the state in column I
                console.log(`Updated state for postcode ${postcode} to ${state}`);
            } else {
                console.log(`No state found for postcode ${postcode}`);
            }
        }
    });

    displayData(currentData);
    enableButtons();
}

function updateTableHeaders(headers) {
    const headerRow = document.querySelector('#previewTable thead tr');
    headerRow.innerHTML = '<th>Action</th>';
    
    // Add the column headers in order
    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'R', 'X'].forEach(col => {
        const th = document.createElement('th');
        th.textContent = headers[col];
        headerRow.appendChild(th);
    });
}

function displayData(data) {
    const tbody = document.getElementById('previewBody');
    tbody.innerHTML = '';

    data.forEach((row, index) => {
        const tr = document.createElement('tr');
        
        // Create the delete button cell first
        const deleteCell = document.createElement('td');
        deleteCell.innerHTML = `
            <button class="btn btn-danger btn-sm" onclick="deleteRow(${index})">
                <i class="bi bi-trash3"></i>
            </button>
        `;
        tr.appendChild(deleteCell);

        // Add all other cells
        ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'R', 'X'].forEach(col => {
            const td = document.createElement('td');
            td.textContent = row[col] || '';
            // Highlight cells where state was looked up
            if (col === 'I' && row.X === 'AU' && row[col]) {
                td.classList.add('bg-light', 'fw-bold');
            }
            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    const statusElement = document.getElementById('processingStatus');
    statusElement.textContent = 'File processed successfully!';
    statusElement.style.opacity = '1';
    
    // Fade out after 3 seconds
    setTimeout(() => {
        statusElement.style.opacity = '0';
        // Hide element completely after fade
        setTimeout(() => {
            statusElement.style.display = 'none';
        }, 500);
    }, 3000);
}

function deleteRow(index) {
    currentData.splice(index, 1);
    displayData(currentData);
}

function clearAll() {
    currentData = [];
    columnHeaders = {};
    document.getElementById('previewBody').innerHTML = '<tr><td></td><td colspan="14" class="text-center">No data loaded</td></tr>';
    document.getElementById('excelFile').value = '';
    document.getElementById('emailUpdateFile').value = '';  // Clear email update file input
    document.getElementById('emailUpdateFile').disabled = true;  // Disable email update file input
    document.getElementById('processingStatus').style.display = 'none';
    
    // Reset headers to default
    const headerRow = document.querySelector('#previewTable thead tr');
    headerRow.innerHTML = `
    <th>Action</th>
    <th>Order Number</th>
    <th>Order date</th>
    <th>Name</th>
    <th>Add. Line 1</th>
    <th>Add. Line 2</th>
    <th>Add. Line 3</th>
    <th>City</th>
    <th>Postcode</th>
    <th>Destination State</th>
    <th>Destination Country</th>
    <th>Email</th>
    <th>Phone</th>
    <th>Reference</th>
    <th>Country Code</th>
`;
    
    disableButtons();
}

function exportToCsv() {
    // Get date from column B of first row
    let filename = 'DHL Report.csv';
    if (currentData.length > 0 && currentData[0].B) {
        const formattedDate = formatDateForFilename(currentData[0].B);
        if (formattedDate) {
            filename = `${formattedDate}_DHL Report.csv`;
        }
    }

    // Define all columns in order
    const columns = [
        'Order Number', 'Date', 'To Name', 'Destination Building', 
        'Destination Street', 'Destination Suburb', 'Destination City',
        'Destination Postcode', 'Destination State', 'Destination Country',
        'Destination Email', 'Destination Phone', 'Item Name', 'Item Price',
        'Instructions', 'Weight', 'Shipping Method', 'Reference', 'SKU',
        'Qty', 'Company', 'Signature Required', 'ATL', 'Country Code',
        'Package Height', 'Package Width', 'Package Length', 'Carrier',
        'Carrier Product Code', 'Carrier Product Unit Type', 
        'Declared Value Currency', 'Code', 'Color', 'Size', 'Contents',
        'Dangerous Goods', 'Country of Manufacturer', 'DDP', 'ReceiverVAT',
        'ReceiverEORI', 'ShippingFreightValue', 'Brand', 'Usage',
        'Material', 'Model', 'MID Code', 'Receiver National ID',
        'Receiver Passport Number', 'Receiver USCI', 'Receiver CR',
        'Receiver Brazil CNP'
    ];

    // Prepare data with all columns
    const exportData = currentData.map(row => {
        const fullRow = {};
        
        // Set default values for all columns
        columns.forEach(col => {
            fullRow[col] = '';
        });

        // Map data from Excel columns to DHL report columns
        const columnMapping = {
            'A': 'Order Number',
            'B': 'Date',
            'C': 'To Name',
            'D': 'Destination Building',
            'E': 'Destination Street',
            'F': 'Destination Suburb',
            'G': 'Destination City',
            'H': 'Destination Postcode',
            'I': 'Destination State',
            'J': 'Destination Country',
            'K': 'Destination Email',
            'L': 'Destination Phone',
            'P': 'Weight',
            'R': 'Reference',
            'X': 'Country Code'
        };

        // Map the data using the column mapping
        Object.entries(columnMapping).forEach(([excelCol, dhlCol]) => {
            if (row[excelCol] !== undefined && row[excelCol] !== '') {
                // Special handling for date format
                if (dhlCol === 'Date' && row[excelCol]) {
                    // Check if the date is already in the correct format
                    if (row[excelCol].includes('-')) {
                        fullRow[dhlCol] = row[excelCol];
                    } else {
                        // Try to parse and format the date
                        try {
                            const date = new Date(row[excelCol]);
                            if (!isNaN(date.getTime())) {
                                const day = String(date.getDate()).padStart(2, '0');
                                const month = String(date.getMonth() + 1).padStart(2, '0');
                                const year = date.getFullYear();
                                fullRow[dhlCol] = `${day}-${month}-${year}`;
                            } else {
                                fullRow[dhlCol] = row[excelCol];
                            }
                        } catch (e) {
                            fullRow[dhlCol] = row[excelCol];
                        }
                    }
                } else {
                    fullRow[dhlCol] = row[excelCol];
                }
            }
        });

        // Set fixed values for specific columns
        fullRow['Item Name'] = 'Printed Books';
        fullRow['Item Price'] = '10';
        fullRow['Instructions'] = '';
        fullRow['Carrier'] = 'DHL';
        fullRow['Carrier Product Unit Type'] = 'Parcel';
        fullRow['Declared Value Currency'] = 'GBP';
        fullRow['Code'] = '49019900';
        fullRow['Contents'] = 'Printed Books';
        fullRow['Dangerous Goods'] = '';
        fullRow['DDP'] = 'Y';
        fullRow['Qty'] = '1';
        fullRow['Carrier Product Unit Type'] = 'Box';
        
        return fullRow;
    });

    // Convert to CSV with specific column order
    const csvContent = Papa.unparse({
        fields: columns,
        data: exportData
    });
    
    // Create and trigger download
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function enableButtons() {
    document.getElementById('clearBtn').disabled = false;
    document.getElementById('exportBtn').disabled = false;
}

function disableButtons() {
    document.getElementById('clearBtn').disabled = true;
    document.getElementById('exportBtn').disabled = true;
}