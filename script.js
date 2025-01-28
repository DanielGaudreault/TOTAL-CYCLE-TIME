let results = []; // Store results for files
let workbook = null; // To hold the Excel workbook

// Function to process both the Excel and PDF files
function processFiles() {
    const excelFile = document.getElementById('excelFile').files[0];
    const pdfFiles = document.getElementById('pdfFiles').files;

    if (!excelFile || pdfFiles.length === 0) {
        alert("Please upload both an Excel file and PDF files.");
        return;
    }

    // Clear the results table before inserting new results
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    resultsTable.innerHTML = ''; // Clear the table

    // Show loading indicator
    document.getElementById('loading').style.display = 'block';
    document.getElementById('downloadButton').style.display = 'none';

    // Read the Excel file
    const reader = new FileReader();
    reader.onload = function (event) {
        const excelData = event.target.result;
        workbook = XLSX.read(excelData, { type: 'binary' });

        // Assume the first sheet is the one we want to work with
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let processedCount = 0;
        const cycleTimes = [];

        // Loop through each PDF
        for (let i = 0; i < pdfFiles.length; i++) {
            const pdfFile = pdfFiles[i];
            const pdfReader = new FileReader();

            pdfReader.onload = async function (event) {
                const pdfData = event.target.result;
                const pdfText = await extractTextFromPDF(pdfData);
                const cycleTime = extractCycleTime(pdfText);

                // Add the cycle time to the results table
                cycleTimes.push({ filename: pdfFile.name, cycleTime: cycleTime });

                // Add the extracted cycle time and setup name to the corresponding row in the Excel data
                const rowIndex = i + 1; // Adjust row index based on your needs
                if (rows[rowIndex]) {
                    rows[rowIndex][2] = pdfFile.name;  // Column C: Setup name (PDF file name)
                    rows[rowIndex][3] = cycleTime;    // Column D: Cycle Time
                }

                processedCount++;

                // Once all PDFs are processed, update the table and show the download button
                if (processedCount === pdfFiles.length) {
                    updateResultsTable(cycleTimes);
                    document.getElementById('downloadButton').style.display = 'inline-block';
                    document.getElementById('loading').style.display = 'none';
                }
            };

            pdfReader.readAsArrayBuffer(pdfFile); // Read PDF as ArrayBuffer
        }
    };

    reader.readAsBinaryString(excelFile); // Read Excel file
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
                return match[0]; // Return the matched cycle time
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
        row.insertCell().textContent = result.filename;  // PDF file name
        row.insertCell().textContent = result.cycleTime;  // Cycle time extracted
    });
}

// Download the updated Excel file
function downloadExcel() {
    if (!workbook) {
        alert("No workbook to download.");
        return;
    }

    // Get the first sheet and update it
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert the modified rows into a sheet
    const updatedSheet = XLSX.utils.aoa_to_sheet(XLSX.utils.sheet_to_json(sheet, { header: 1 }));

    // Create a new workbook with the updated sheet
    const updatedWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(updatedWorkbook, updatedSheet, sheetName);

    // Generate and download the updated Excel file
    XLSX.writeFile(updatedWorkbook, 'updated_cycle_times.xlsx');
}

// Reset results function to clear the table and reset the form
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

    // Reset results array
    results = [];
}
