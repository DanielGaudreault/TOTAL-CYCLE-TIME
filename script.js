let workbook = null; // To hold the Excel workbook
let excelRows = []; // Store rows from the Excel sheet

// Function to process both the Excel and PDF files
async function processFiles() {
    const excelFile = document.getElementById('excelFile').files[0];
    const pdfFiles = document.getElementById('pdfFiles').files;

    if (!excelFile || pdfFiles.length === 0) {
        alert("Please upload both an Excel file and PDF files.");
        return;
    }

    // Show loading indicator
    document.getElementById('loading').style.display = 'block';
    document.getElementById('downloadButton').style.display = 'none';

    // Read the Excel file
    const reader = new FileReader();
    reader.onload = async function (event) {
        const excelData = event.target.result;
        workbook = XLSX.read(excelData, { type: 'binary' });

        // Read the first sheet and extract rows
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        excelRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Handle PDF extraction concurrently
        const cycleTimes = await Promise.all(Array.from(pdfFiles).map((pdfFile) => {
            return extractCycleDataFromPDF(pdfFile);
        }));

        // Now process extracted PDF data
        cycleTimes.forEach((cycleData) => {
            const { projectName, totalCycleTime, setupName } = cycleData;

            // Find matching project names in the Excel sheet and update columns C (setup name) and D (cycle time)
            for (let rowIndex = 1; rowIndex < excelRows.length; rowIndex++) {
                if (excelRows[rowIndex][0] && excelRows[rowIndex][0].toLowerCase() === projectName.toLowerCase()) {
                    excelRows[rowIndex][2] = setupName;  // Update Setup Name (Column C)
                    excelRows[rowIndex][3] = totalCycleTime;  // Update Total Cycle Time (Column D)
                    break; // Exit once matched project name is found
                }
            }
        });

        // Once processing is complete, update the results table and show the download button
        updateResultsTable(cycleTimes);
        document.getElementById('downloadButton').style.display = 'inline-block';
        document.getElementById('loading').style.display = 'none';
    };

    reader.readAsBinaryString(excelFile); // Read Excel file
}

// Extract cycle data (project name, cycle time, setup name) from PDF
async function extractCycleDataFromPDF(pdfFile) {
    const pdfData = await readPDFFile(pdfFile);
    const pdfText = await extractTextFromPDF(pdfData);

    const projectName = extractProjectName(pdfText);
    const totalCycleTime = extractCycleTime(pdfText);
    const setupName = extractSetupName(pdfText);

    return { projectName, totalCycleTime, setupName };
}

// Helper function to read PDF file as ArrayBuffer
function readPDFFile(pdfFile) {
    return new Promise((resolve, reject) => {
        const pdfReader = new FileReader();
        pdfReader.onload = function (event) {
            resolve(event.target.result);
        };
        pdfReader.onerror = function (error) {
            reject(error);
        };
        pdfReader.readAsArrayBuffer(pdfFile);
    });
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

// Extract the project name from PDF (e.g., CNT2301)
function extractProjectName(text) {
    const regex = /PROJECT NAME: (\w+)/i;
    const match = text.match(regex);
    if (match && match[1]) {
        return match[1].trim();
    }
    return 'Not found';
}

// Extract the total cycle time from PDF (e.g., 0 HOURS, 3 MINUTES, 8 SECONDS)
function extractCycleTime(text) {
    const regex = /TOTAL CYCLE TIME: (\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)/i; 
    const match = text.match(regex);
    if (match && match[1]) {
        return match[1];
    }
    return 'Not found';
}

// Extract the setup name from the PDF text
function extractSetupName(text) {
    const regex = /PROJECT NAME: (\w+ \w+)/i; // Example: PROJECT NAME: CNT2301 R1
    const match = text.match(regex);
    if (match && match[1]) {
        return match[1];
    }
    return 'Not found';
}

// Update the results table with the cycle times
function updateResultsTable(cycleTimes) {
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    cycleTimes.forEach((result) => {
        const row = resultsTable.insertRow();
        row.insertCell().textContent = result.projectName;
        row.insertCell().textContent = result.totalCycleTime;
    });
}

// Download the updated Excel file with modified data
function downloadExcel() {
    if (!workbook) {
        alert("No workbook to download.");
        return;
    }

    // Get the first sheet and update it
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert the modified rows into a sheet
    const updatedSheet = XLSX.utils.aoa_to_sheet(excelRows);

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
}
