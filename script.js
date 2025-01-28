let results = []; // Store results for all files

// Function to process both Excel and PDF files
// Function to process both the Excel and PDF files
function processFiles() {
    // Get uploaded files
const excelFile = document.getElementById('excelFile').files[0];
const pdfFiles = document.getElementById('pdfFiles').files;

@@ -10,10 +9,6 @@ function processFiles() {
return;
}

    // Show loading text
    document.getElementById('loading').style.display = 'block';
    document.getElementById('downloadButton').style.display = 'none';

// Read the Excel file
const reader = new FileReader();
reader.onload = function (event) {
@@ -25,51 +20,43 @@ function processFiles() {
const sheet = workbook.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Start processing the PDFs
let processedCount = 0;
const cycleTimes = [];

        // Process each PDF file one by one to avoid large memory load
        let currentPdfIndex = 0;
        processNextPdf(currentPdfIndex, pdfFiles, rows, cycleTimes, processedCount, workbook, sheetName);
    };
        // Loop through each PDF
        for (let i = 0; i < pdfFiles.length; i++) {
            const pdfFile = pdfFiles[i];
            const pdfReader = new FileReader();

    reader.readAsBinaryString(excelFile);
}
            pdfReader.onload = async function (event) {
                const pdfData = event.target.result;
                const pdfText = await extractTextFromPDF(pdfData);
                const cycleTime = extractCycleTime(pdfText);

// Function to handle PDF files one at a time
function processNextPdf(currentPdfIndex, pdfFiles, rows, cycleTimes, processedCount, workbook, sheetName) {
    if (currentPdfIndex >= pdfFiles.length) {
        // All PDFs processed, update results table and show download button
        updateResultsTable(cycleTimes);
        document.getElementById('loading').style.display = 'none';
        document.getElementById('downloadButton').style.display = 'inline-block';
        return;
    }
                // Add the cycle time to the results table
                cycleTimes.push({ filename: pdfFile.name, cycleTime: cycleTime });

    const pdfFile = pdfFiles[currentPdfIndex];
    const pdfReader = new FileReader();
                // Add the extracted cycle time to the corresponding row in the Excel data
                const rowIndex = i + 1; // Adjust row index based on your needs
                if (rows[rowIndex]) {
                    rows[rowIndex][2] = cycleTime; // Put cycle time in Column C (index 2)
                }

    pdfReader.onload = async function (event) {
        const pdfData = event.target.result;
        const pdfText = await extractTextFromPDF(pdfData);
        const cycleTime = extractCycleTime(pdfText);
                processedCount++;

        // Add the cycle time to the results table
        cycleTimes.push({ filename: pdfFile.name, cycleTime });
                // Once all PDFs are processed, update the table and show the download button
                if (processedCount === pdfFiles.length) {
                    updateResultsTable(cycleTimes);
                    document.getElementById('downloadButton').style.display = 'inline-block';
                }
            };

        // Find corresponding row in the Excel data
        const rowIndex = currentPdfIndex + 1; // Match PDF file with row in Excel sheet
        if (rows[rowIndex]) {
            rows[rowIndex][2] = cycleTime; // Add cycle time to Column C (index 2)
            pdfReader.readAsArrayBuffer(pdfFile); // Read PDF as ArrayBuffer
}

        processedCount++;

        // Recursively call the next PDF file
        processNextPdf(currentPdfIndex + 1, pdfFiles, rows, cycleTimes, processedCount, workbook, sheetName);
};

    pdfReader.readAsArrayBuffer(pdfFile);
    reader.readAsBinaryString(excelFile); // Read Excel file
}

// Extract text from PDF using PDF.js
@@ -114,7 +101,7 @@ function extractCycleTime(text) {
const regex = /(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)/i;
const match = line.match(regex);
if (match) {
                return match[0];
                return match[0]; // Return the matched cycle time
}
}
}
@@ -124,72 +111,36 @@ function extractCycleTime(text) {
// Update the results table with cycle times
function updateResultsTable(cycleTimes) {
const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    
    // Clear any previous results
    resultsTable.innerHTML = '';

cycleTimes.forEach((result) => {
const row = resultsTable.insertRow();
row.insertCell().textContent = result.filename;
row.insertCell().textContent = result.cycleTime;
});
}

// Download the updated Excel file (optimized)
function downloadUpdatedExcel() {
// Download the updated Excel file
function downloadExcel() {
const excelFile = document.getElementById('excelFile').files[0];
const reader = new FileReader();

reader.onload = function (event) {
const excelData = event.target.result;
const workbook = XLSX.read(excelData, { type: 'binary' });

        // Assume we are working with the first sheet
        // Assume we're working with the first sheet
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

        // Modify sheet (add the updated cycle times to Column C)
        // Convert the modified rows into a sheet
const updatedSheet = XLSX.utils.aoa_to_sheet(sheet);

// Create a new workbook with the updated sheet
const updatedWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(updatedWorkbook, updatedSheet, sheetName);

        // Trigger the file download using Blob
        const blob = XLSX.write(updatedWorkbook, {
            bookType: 'xlsx',
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });

        // Create a download link and click it to trigger the download
        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = 'updated_cycle_times.xlsx';

        // Trigger the download
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);  // Clean up after the download
        // Generate and download the updated Excel file
        XLSX.writeFile(updatedWorkbook, 'updated_cycle_times.xlsx');
};

reader.readAsBinaryString(excelFile);
}

// Reset Results function to clear the table and reset the form
function resetResults() {
    // Clear the results table
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    resultsTable.innerHTML = '';

    // Hide the download button and reset loading state
    document.getElementById('downloadButton').style.display = 'none';
    document.getElementById('loading').style.display = 'none';

    // Reset the file input fields
    document.getElementById('excelFile').value = '';
    document.getElementById('pdfFiles').value = '';

    // Clear any cycle times stored
    results = [];
}
