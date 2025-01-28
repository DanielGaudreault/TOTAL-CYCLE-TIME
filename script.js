let results = []; // Store results for all files

function processFiles() {
    const fileInput = document.getElementById('fileInput');
    const loading = document.getElementById('loading');
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    const downloadButton = document.getElementById('downloadButton');

    if (fileInput.files.length === 0) {
        alert('Please select at least one file.');
        return;
    }

    // Clear previous results
    results = [];
    resultsTable.innerHTML = '';
    downloadButton.style.display = 'none';
    loading.style.display = 'block';

    const files = Array.from(fileInput.files);

    const promises = files.map(file => {
        const reader = new FileReader();
        return new Promise((resolve, reject) => {
            reader.onload = async function (event) {
                const content = event.target.result;
                let cycleTime = null;

                if (file.type === 'application/pdf') {
                    // Handle PDF files
                    const text = await parsePDF(content);
                    cycleTime = extractCycleTime(text);
                } else {
                    // Handle text files
                    cycleTime = extractCycleTime(content);
                }

                // Add result to the table
                results.push({ fileName: file.name, cycleTime });
                const row = resultsTable.insertRow();
                row.insertCell().textContent = file.name;
                row.insertCell().textContent = cycleTime || 'Not Found';

                resolve(); // Signal that this file has been processed
            };

            reader.onerror = reject;

            if (file.type === 'application/pdf') {
                reader.readAsArrayBuffer(file); // Read PDF as ArrayBuffer
            } else {
                reader.readAsText(file); // Read text files as text
            }
        });
    });

    Promise.all(promises).then(() => {
        loading.style.display = 'none';
        downloadButton.style.display = 'block';
    }).catch(error => {
        console.error("File processing failed:", error);
        loading.style.display = 'none';
    });
}

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

function parsePDF(data) {
    return new Promise((resolve, reject) => {
        const pdfjsLib = window['pdfjs-dist/build/pdf'];
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

        const loadingTask = pdfjsLib.getDocument({ data });
        loadingTask.promise.then(pdf => {
            let text = '';
            const numPages = pdf.numPages;

            const fetchPage = (pageNum) => {
                return pdf.getPage(pageNum).then(page => {
                    return page.getTextContent().then(textContent => {
                        let pageText = '';
                        textContent.items.forEach(item => {
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

function resetResults() {
    // Clear the results array
    results = [];

    // Clear the table
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    resultsTable.innerHTML = '';

    // Hide the download button
    document.getElementById('downloadButton').style.display = 'none';

    // Clear the file input (this needs to be done for each file input)
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadExcelInput').value = '';
}

function downloadResults() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        // If no file has been uploaded, just download the results from PDFs
        const ws = XLSX.utils.json_to_sheet(results.map(result => ({
            'Item No.': result.fileName.split('Project Name -')[1]?.trim() || result.fileName,
            'Setup Number': 'Setup Number', 
            'Total Cycle Time': result.cycleTime || 'Not Found'
        })));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Results');
        XLSX.writeFile(wb, 'cycle_times.xlsx');
    } else {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const sheetName = workbook.SheetNames[0];
            let worksheet = workbook.Sheets[sheetName];
            let excelRows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

            // Update with results from PDFs
            results.forEach(result => {
                const matchIdentifier = result.fileName.split('Project Name -')[1]?.trim() || result.fileName;
                const matchingRowIndex = excelRows.findIndex(row => row[0]?.toString().trim() === matchIdentifier);

                if (matchingRowIndex !== -1) {
                    excelRows[matchingRowIndex] = {
                        ...excelRows[matchingRowIndex],
                        2: 'Setup Number', // Assuming 'Setup Number' should be in column C
                        3: result.cycleTime || 'Not Found' // Assuming 'Total Cycle Time' should be in column D
                    };
                } else {
                    // If no match, add new row
                    excelRows.push({
                        'Item No.': matchIdentifier,
                        '': '', // Assuming column B is blank or something else, adjust if necessary
                        'Setup Number': 'Setup Number',
                        'Total Cycle Time': result.cycleTime || 'Not Found'
                    });
                }
            });

            // Convert back to worksheet
            const updatedWS = XLSX.utils.json_to_sheet(excelRows);
            const updatedWB = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(updatedWB, updatedWS, sheetName);

            // Save the updated workbook
            XLSX.writeFile(updatedWB, 'cycle_times.xlsx');
        };
        reader.readAsArrayBuffer(file);
    }
}

function updateToExcel() {
    const fileInput = document.getElementById('uploadExcelInput');
    const file = fileInput.files[0];

    if (!file) {
        alert('Please select an Excel file to update.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});

        // Assuming data is in the first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        let excelRows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

        // Process each result from the PDFs
        results.forEach(result => {
            // Assuming the project name in the PDF is immediately followed by another identifier (like a number or code)
            // Let's assume this identifier is the part of the filename after "Project Name -"
            const matchIdentifier = result.fileName.split('Project Name -')[1]?.trim();
            
            if (!matchIdentifier) {
                console.error('Could not extract match identifier from PDF file name:', result.fileName);
                return; // Skip this result if we can't find a match identifier
            }

            // Try to find an existing row in the Excel sheet where the identifier matches the 'Item No.'
            const matchingRowIndex = excelRows.findIndex(row => row[0]?.toString().trim() === matchIdentifier);

            if (matchingRowIndex !== -1) {
                // Match found, update existing row
                excelRows[matchingRowIndex] = {
                    ...excelRows[matchingRowIndex],
                    2: 'Setup Number', // Update column C with 'Setup Number'
                    3: result.cycleTime || 'Not Found' // Update column D with total cycle time
                };
            } else {
                // No match found, add new row
                excelRows.push([
                    matchIdentifier, // Assuming 'Item No.' column A
                    '', // B might be something else, keeping it blank
                    'Setup Number',  // C - Setup Number
                    result.cycleTime || 'Not Found', // D - Total Cycle Time
                    ...Array(excelRows[0].length - 4).fill('') // Fill rest with blanks to match row length
                ]);
            }
        });

        // Convert back to a worksheet format
        const newWS = XLSX.utils.aoa_to_sheet(excelRows.map(row => Object.values(row)));
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, newWS, sheetName);

        // Save the new workbook
        XLSX.writeFile(newWB, 'updated_cycle_times.xlsx');

        // Optionally update the table on the page if needed
        const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
        resultsTable.innerHTML = '';
        results.forEach(result => {
            const newRow = resultsTable.insertRow();
            newRow.insertCell().textContent = result.fileName;
            newRow.insertCell().textContent = result.cycleTime || 'Not Found';
        });

        document.getElementById('downloadButton').style.display = 'block';
    };
    reader.readAsArrayBuffer(file);
}

// Event Listeners
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processButton').addEventListener('click', processFiles);
    document.getElementById('resetButton').addEventListener('click', resetResults);
    document.getElementById('downloadButton').addEventListener('click', downloadResults);
    document.getElementById('uploadExcelButton').addEventListener('click', updateToExcel);
});
