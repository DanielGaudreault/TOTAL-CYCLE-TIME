let results = []; // Store results for files
let workbook = null; // To hold the Excel workbook

// Function to process both the Excel and PDF files
function processFiles() {
async function processFiles() {
const excelFile = document.getElementById('excelFile').files[0];
const pdfFiles = document.getElementById('pdfFiles').files;

@@ -21,7 +21,7 @@

// Read the Excel file
const reader = new FileReader();
    reader.onload = function (event) {
    reader.onload = async function (event) {
const excelData = event.target.result;
workbook = XLSX.read(excelData, { type: 'binary' });

@@ -30,52 +30,60 @@
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

                // Extract information from the PDF
                const projectName = extractProjectName(pdfText); // Extract project name (e.g. CNT2301)
                const totalCycleTime = extractCycleTime(pdfText); // Extract total cycle time
                const setupName = extractSetupName(pdfText); // Extract setup name (e.g. CNT2301 R1)

                // Add the cycle time to the results table
                cycleTimes.push({ filename: pdfFile.name, cycleTime: totalCycleTime });

                // Find matching project names and update the corresponding rows
                for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
                    if (rows[rowIndex][0] && rows[rowIndex][0].toLowerCase() === projectName.toLowerCase()) {
                        // Match found, update columns C and D
                        rows[rowIndex][2] = setupName;  // Setup name in Column C
                        rows[rowIndex][3] = totalCycleTime;  // Cycle time in Column D
                    }
        // Handle PDF extraction concurrently with Promise.all
        const cycleTimes = await Promise.all(Array.from(pdfFiles).map((pdfFile) => {
            return extractCycleDataFromPDF(pdfFile);
        }));

        // Process the extracted data and update Excel rows
        cycleTimes.forEach((cycleData) => {
            const { projectName, totalCycleTime, setupName } = cycleData;

            // Find matching project names and update the corresponding rows in Excel sheet
            for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
                if (rows[rowIndex][0] && rows[rowIndex][0].toLowerCase() === projectName.toLowerCase()) {
                    // Match found, update columns C and D
                    rows[rowIndex][2] = setupName;  // Setup name in Column C
                    rows[rowIndex][3] = totalCycleTime;  // Cycle time in Column D
}
            }
        });

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
        // Once all PDFs are processed, update the table and show the download button
        updateResultsTable(cycleTimes);
        document.getElementById('downloadButton').style.display = 'inline-block';
        document.getElementById('loading').style.display = 'none';
};

reader.readAsBinaryString(excelFile); // Read Excel file
}

// Extract cycle data from PDF (project name, cycle time, setup name)
async function extractCycleDataFromPDF(pdfFile) {
    const pdfData = await readPDFFile(pdfFile);
    const pdfText = await extractTextFromPDF(pdfData);

    const projectName = extractProjectName(pdfText); // Extract project name (e.g. CNT2301)
    const totalCycleTime = extractCycleTime(pdfText); // Extract total cycle time
    const setupName = extractSetupName(pdfText); // Extract setup name (e.g. CNT2301 R1)

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
@@ -145,8 +153,8 @@
const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
cycleTimes.forEach((result) => {
const row = resultsTable.insertRow();
        row.insertCell().textContent = result.filename;  // PDF file name
        row.insertCell().textContent = result.cycleTime;  // Cycle time extracted
        row.insertCell().textContent = result.projectName;  // Project name
        row.insertCell().textContent = result.totalCycleTime;  // Total cycle time
});
}
