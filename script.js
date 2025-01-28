document.addEventListener('DOMContentLoaded', () => {
    // Event listener for "Process Files" button
    document.getElementById('processButton').addEventListener('click', processFiles);

    // Event listener for "Reset Results" button
    document.getElementById('resetButton').addEventListener('click', resetResults);

    // Event listener for "Upload to Excel" button
    document.getElementById('uploadExcelButton').addEventListener('click', updateToExcel);
});

let results = []; // Store results for all files

// Process selected files (PDFs or text)
function processFiles() {
    const fileInput = document.getElementById('fileInput');
    const loading = document.getElementById('loading');
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];

    if (fileInput.files.length === 0) {
        alert('Please select at least one file.');
        return;
    }

    // Clear previous results
    results = [];
    resultsTable.innerHTML = '';
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

// Parse PDF content using pdf.js
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

// Reset the results
function resetResults() {
    results = [];
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    resultsTable.innerHTML = '';
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadExcelInput').value = '';
}

// Update Excel with the summed cycle times
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

        // Create an object to store the total cycle time for each project
        let cycleTimeSums = {};

        // Process each result from the PDFs
        results.forEach(result => {
            // Extract the project name from the PDF filename, assuming it's something like "Project Name - Identifier"
            const matchIdentifier = result.fileName.match(/^(.+?) - \d+/)?.[1].trim(); // Adjusting regex

            if (!matchIdentifier) {
                console.error('Could not extract match identifier from PDF file name:', result.fileName);
                return; // Skip this result if we can't find a match identifier
            }

            // If the project already exists in the cycleTimeSums map, add the cycle time, otherwise set it
            if (cycleTimeSums[matchIdentifier]) {
                cycleTimeSums[matchIdentifier] += parseCycleTime(result.cycleTime);
            } else {
                cycleTimeSums[matchIdentifier] = parseCycleTime(result.cycleTime);
            }
        });

        // Now we go through the excelRows and match the project name (from the cycleTimeSums) with the 'Item No.' column
        excelRows.forEach((row, rowIndex) => {
            const itemNo = row[0]?.toString().trim(); // Assuming 'Item No.' is in the first column
            if (cycleTimeSums[itemNo]) {
                // Match found, update the cycle time sum in the 'Total Cycle Time' column (assuming column D is where it goes)
                row[3] = formatCycleTime(cycleTimeSums[itemNo]); // Assuming 'Total Cycle Time' is in column 4 (index 3)
            }
        });

        // Convert the updated data back to worksheet format
        const newWS = XLSX.utils.aoa_to_sheet(excelRows);
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, newWS, sheetName);

        // Save the new workbook
        XLSX.writeFile(newWB, 'updated_cycle_times.xlsx');
    };
    reader.readAsArrayBuffer(file);
}

// Helper function to parse cycle time (e.g., "1 HOURS, 30 MINUTES, 0 SECONDS" -> total seconds)
function parseCycleTime(cycleTimeString) {
    if (!cycleTimeString) return 0;

    const regex = /(\d+)\s*HOURS?/, hoursMatch = cycleTimeString.match(regex);
    const regexMinutes = /(\d+)\s*MINUTES?/, minutesMatch = cycleTimeString.match(regexMinutes);
    const regexSeconds = /(\d+)\s*SECONDS?/, secondsMatch = cycleTimeString.match(regexSeconds);

    const hours = hoursMatch ? parseInt(hoursMatch[1]) : 0;
    const minutes = minutesMatch ? parseInt(minutesMatch[1]) : 0;
    const seconds = secondsMatch ? parseInt(secondsMatch[1]) : 0;

    return (hours * 3600) + (minutes * 60) + seconds;
}

// Helper function to format cycle time back to a readable string (seconds -> "X HOURS, Y MINUTES, Z SECONDS")
function formatCycleTime(totalSeconds) {
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;

    return `${hours} HOURS, ${minutes} MINUTES, ${seconds} SECONDS`;
}
