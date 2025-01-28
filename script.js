function searchFile() {
    const fileInput = document.getElementById('fileInput');
    const results = document.getElementById('results');
    results.textContent = ''; // Clear previous results
// Function to process both the Excel and PDF files
function processFiles() {
    // Get uploaded files
    const excelFile = document.getElementById('excelFile').files[0];
    const pdfFiles = document.getElementById('pdfFiles').files;

    if (fileInput.files.length === 0) {
        alert('Please select a file.');
    if (!excelFile || pdfFiles.length === 0) {
        alert("Please upload both an Excel file and PDF files.");
return;
}

    const file = fileInput.files[0];
    // Read the Excel file
const reader = new FileReader();

reader.onload = function (event) {
        const content = event.target.result;

        if (file.type === 'application/pdf') {
            // Handle PDF files
            parsePDF(content).then(text => {
                const cycleTime = extractCycleTime(text);
                const setupName = extractSetupName(text);
                displayResults(cycleTime, setupName);
                sendDataToBackend(cycleTime, setupName, file.name);
            });
        } else {
            // Handle text files
            const cycleTime = extractCycleTime(content);
            const setupName = extractSetupName(content);
            displayResults(cycleTime, setupName);
            sendDataToBackend(cycleTime, setupName, file.name);
        }
    };
        const excelData = event.target.result;
        const workbook = XLSX.read(excelData, { type: 'binary' });

    if (file.type === 'application/pdf') {
        reader.readAsArrayBuffer(file); // Read PDF as ArrayBuffer
    } else {
        reader.readAsText(file); // Read text files as text
    }
}
        // Assume the first sheet is the one we want to work with
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

function extractCycleTime(text) {
    const lines = text.split('\n'); // Split text into lines
    for (const line of lines) {
        if (line.includes("TOTAL CYCLE TIME")) {
            // Use regex to extract the time part (e.g., "0 HOURS, 4 MINUTES, 16 SECONDS")
            const regex = /(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)/i;
            const match = line.match(regex);
            return match ? match[0] : null; // Return the matched time or null
        }
    }
    return null; // Return null if no match is found
}
        // Start processing the PDFs
        let processedCount = 0;
        const cycleTimes = [];

function extractSetupName(text) {
    const lines = text.split('\n'); // Split text into lines
    for (const line of lines) {
        if (line.includes("Setup Name:")) {
            // Extract the setup name after the colon
            const setupName = line.split("Setup Name:")[1].trim();
            return setupName || null; // Return the setup name or null
        }
    }
    return null; // Return null if no match is found
}
        // Loop through each PDF
        for (let i = 0; i < pdfFiles.length; i++) {
            const pdfFile = pdfFiles[i];
            const pdfReader = new FileReader();

function displayResults(cycleTime, setupName) {
    const results = document.getElementById('results');
    if (cycleTime && setupName) {
        results.textContent = `Cycle Time: ${cycleTime}, Setup Name: ${setupName}`;
    } else {
        results.textContent = 'No instances of "TOTAL CYCLE TIME" or "Setup Name" found.';
    }
}
            pdfReader.onload = async function (event) {
                const pdfData = event.target.result;
                const pdfText = await extractTextFromPDF(pdfData);
                const cycleTime = extractCycleTime(pdfText);

                // Add the cycle time to the results table
                cycleTimes.push({ filename: pdfFile.name, cycleTime: cycleTime });

                // Add the extracted cycle time to the corresponding row in the Excel data
                const rowIndex = i + 1; // Adjust row index based on your needs
                if (rows[rowIndex]) {
                    rows[rowIndex][2] = cycleTime; // Put cycle time in Column C (index 2)
                }

                processedCount++;

                // Once all PDFs are processed, update the table and show the download button
                if (processedCount === pdfFiles.length) {
                    updateResultsTable(cycleTimes);
                    document.getElementById('downloadButton').style.display = 'inline-block';
                }
            };

function sendDataToBackend(cycleTime, setupName, fileName) {
    const data = {
        cycleTime: cycleTime,
        setupName: setupName,
        fileName: fileName
            pdfReader.readAsArrayBuffer(pdfFile); // Read PDF as ArrayBuffer
        }
};

    fetch('/update-excel', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
    })
    .then(response => response.json())
    .then(result => {
        console.log(result.message);
    })
    .catch(error => {
        console.error('Error:', error);
    });
    reader.readAsBinaryString(excelFile); // Read Excel file
}

function parsePDF(data) {
// Extract text from PDF using PDF.js
function extractTextFromPDF(pdfData) {
return new Promise((resolve, reject) => {
const pdfjsLib = window['pdfjs-dist/build/pdf'];
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        const loadingTask = pdfjsLib.getDocument({ data: pdfData });

        const loadingTask = pdfjsLib.getDocument({ data });
        loadingTask.promise.then(pdf => {
        loadingTask.promise.then((pdf) => {
let text = '';
const numPages = pdf.numPages;

const fetchPage = (pageNum) => {
                return pdf.getPage(pageNum).then(page => {
                    return page.getTextContent().then(textContent => {
                return pdf.getPage(pageNum).then((page) => {
                    return page.getTextContent().then((textContent) => {
let pageText = '';
                        textContent.items.forEach(item => {
                        textContent.items.forEach((item) => {
pageText += item.str + ' ';
});
text += pageText + '\n'; // Add newline after each page
@@ -128,3 +92,55 @@ function parsePDF(data) {
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
        row.insertCell().textContent = result.filename;
        row.insertCell().textContent = result.cycleTime;
    });
}

// Download the updated Excel file
function downloadExcel() {
    const excelFile = document.getElementById('excelFile').files[0];
    const reader = new FileReader();

    reader.onload = function (event) {
        const excelData = event.target.result;
        const workbook = XLSX.read(excelData, { type: 'binary' });

        // Assume we're working with the first sheet
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // Convert the modified rows into a sheet
        const updatedSheet = XLSX.utils.aoa_to_sheet(sheet);

        // Create a new workbook with the updated sheet
        const updatedWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(updatedWorkbook, updatedSheet, sheetName);

        // Generate and download the updated Excel file
        XLSX.writeFile(updatedWorkbook, 'updated_cycle_times.xlsx');
    };

    reader.readAsBinaryString(excelFile);
}
