let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Load the first sheet
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        // Initially display the full sheet
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = ''; // Clear existing content

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    sheetData.forEach(rowData => {
        const row = document.createElement('tr');
        Object.values(rowData).forEach(cellData => {
            const td = document.createElement('td');
            td.textContent = cellData === null ? 'NULL' : cellData; // Display 'NULL' for null values
            row.appendChild(td);
        });
        table.appendChild(row);
    });

    sheetContentDiv.appendChild(table);
}

// Apply operation based on user input
document.getElementById('apply-operation').addEventListener('click', () => {
    const primaryCol = document.getElementById('primary-column').value.toUpperCase();
    const operationCols = document.getElementById('operation-columns').value.toUpperCase().split(',');

    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    const primaryIndex = getColumnIndex(primaryCol);
    const operationIndices = operationCols.map(getColumnIndex);

    filteredData = data.filter(row => {
        const primaryValue = row[primaryCol];
        const operationValues = operationIndices.map(index => row[Object.keys(data[0])[index]]);

        if (operationType === 'and') {
            return (operation === 'null') 
                ? operationValues.every(value => value === null) 
                : operationValues.every(value => value !== null);
        } else { // OR operation
            return (operation === 'null') 
                ? operationValues.some(value => value === null) 
                : operationValues.some(value => value !== null);
        }
    });

    displaySheet(filteredData); // Display filtered data
});

// Helper function to convert column letter to index
function getColumnIndex(col) {
    return col.charCodeAt(0) - 'A'.charCodeAt(0);
}

// Download button event
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});

// Close modal event
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Confirm download event
document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value;
    const format = document.getElementById('file-format').value;

    if (format === 'xlsx') {
        downloadExcel(filteredData, filename);
    } else if (format === 'csv') {
        downloadCSV(filteredData, filename);
    } else if (format === 'pdf') {
        downloadPDF(filteredData, filename);
    } else if (format === 'jpg' || format === 'jpeg') {
        downloadImage(filteredData, filename);
    }

    document.getElementById('download-modal').style.display = 'none';
});

// Download Excel file
function downloadExcel(data, filename) {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, `${filename}.xlsx`);
}

// Download CSV file
function downloadCSV(data, filename) {
    const csvContent = data.map(row => Object.values(row).join(",")).join("\n");
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}.csv`;
    a.click();
    URL.revokeObjectURL(url);
}

// Download PDF file
function downloadPDF(data, filename) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    let colNames = Object.keys(data[0]);
    let rows = data.map(row => Object.values(row));

    doc.autoTable({
        head: [colNames],
        body: rows,
    });

    doc.save(`${filename}.pdf`);
}

// Download Image (JPEG)
function downloadImage(data, filename) {
    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(data[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    data.forEach(rowData => {
        const row = document.createElement('tr');
        Object.values(rowData).forEach(cellData => {
            const td = document.createElement('td');
            td.textContent = cellData === null ? 'NULL' : cellData; // Display 'NULL' for null values
            row.appendChild(td);
        });
        table.appendChild(row);
    });

    document.body.appendChild(table);

    html2canvas(table).then(canvas => {
        const imgData = canvas.toDataURL('image/jpeg');
        const link = document.createElement('a');
        link.href = imgData;
        link.download = `${filename}.jpg`;
        link.click();
        document.body.removeChild(table); // Remove table after downloading
    });
}

// Load the Excel sheet on initial load
window.onload = () => {
    loadExcelSheet('path/to/your/excel-file.xlsx'); // Replace with your Excel file path
};
