// Function to process both Excel and PDF files
function processFiles() {
    const excelFile = document.getElementById('excelFile').files[0];
    const pdfFiles = document.getElementById('pdfFiles').files;

    if (!excelFile || pdfFiles.length === 0) {
        alert("Please upload both an Excel file and PDF files.");
        return;
    }

    // Show loading text
    document.getElementById('loading').style.display = 'block';
    document.getElementById('downloadButton').style.display = 'none';

    // Read the Excel file
    const reader = new FileReader();
    reader.onload = function (event) {
        const excelData = event.target.result;
        const workbook = XLSX.read(excelData, { type: 'binary' });

        // Assume the first sheet is the one we want to work with
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let processedCount = 0;
        const cycleTimes = [];

        // Process each PDF file
        for (let i = 0; i < pdfFiles.length; i++) {
            const pdfFile = pdfFiles[i];
            const pdfReader = new FileReader();

            pdfReader.onload = async function (event) {
                const pdfData = event.target.result;
                const pdfText = await extractTextFromPDF(pdfData);
                const cycleTime = extractCycleTime(pdfText);

                // Add the cycle time to the results table
                cycleTimes.push({ filename: pdfFile.name, cycleTime });

                // Find corresponding row in the Excel data
                const rowIndex = i + 1; // Match PDF file with row in Excel sheet
                if (rows[rowIndex]) {
                    rows[rowIndex][2] = cycleTime; // Add cycle time to Column C (index 2)
                }

                processedCount++;

                // If all files are processed, update the table and show download button
                if (processedCount === pdfFiles.length) {
                    updateResultsTable(cycleTimes);
                    document.getElementById('loading').style.display = 'none';
                    document.getElementById('downloadButton').style.display = 'inline-block';
                }
            };

            pdfReader.readAsArrayBuffer(pdfFile);
        }
    };

    reader.readAsBinaryString(excelFile);
}

// Extract text from PDF using PDF.js
function extractTextFromPDF(pdfData) {
    return new Promise((resolve, reject) => {
        const pdfjsLib = window['pdfjs-dist/build/pdf'];
        const loadingTask = pdfjsLib.getDocument({ data: pdfData });

        loadingTask.promise.then((pdf) => {
            let text = '';
            const numPages = pdf.numPages;

            const fetchPage = (pageNum) => {
                return pdf.getPage(pageNum).then((page) => {
                    return page.getTextContent().then((textContent) => {
                        let pageText = '';
                        textContent.items.forEach((item) => {
                            pageText += item.str + ' ';
                        });
                        text += pageText + '\n'; // Add newline after each page
                    });
                });
            };

            const fetchAllPages = async () => {
                for (let i = 1; i <= numPages; i++) {
                    await fetchPage(i);
                }
                resolve(text);
            };

            fetchAllPages();
        }).catch(reject);
    });
}

// Extract cycle time from text
function extractCycleTime(text) {
    const lines = text.split('\n');
    for (const line of lines) {
        if (line.includes("TOTAL CYCLE TIME")) {
            const regex = /(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)/i;
            const match = line.match(regex);
            if (match) {
                return match[0];
            }
        }
    }
    return 'Not found';
}

// Update the results table with cycle times
function updateResultsTable(cycleTimes) {
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    cycleTimes.forEach((result) => {
        const row = resultsTable.insertRow();
        row.insertCell().textContent = result.filename;
        row.insertCell().textContent = result.cycleTime;
    });
}

// Download the updated Excel file
function downloadUpdatedExcel() {
    const excelFile = document.getElementById('excelFile').files[0];
    const reader = new FileReader();

    reader.onload = function (event) {
        const excelData = event.target.result;
        const workbook = XLSX.read(excelData, { type: 'binary' });

        // Assume we are working with the first sheet
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // Convert the modified rows into a sheet
        const updatedSheet = XLSX.utils.aoa_to_sheet(sheet);

        // Create a new workbook with the updated sheet
        const updatedWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(updatedWorkbook, updatedSheet, sheetName);

        // Trigger the file download using Blob
        const blob = XLSX.write(updatedWorkbook, { bookType: 'xlsx', type: 'blob' });

        // Create a download link and click it to trigger the download
        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = 'updated_cycle_times.xlsx';

        // Ensure the download link is properly triggered
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);  // Clean up after the download
    };

    reader.readAsBinaryString(excelFile);
}
