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
async function processFiles() {
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
    loading.textContent = 'Processing files...';

    const files = Array.from(fileInput.files);

    try {
        for (const file of files) {
            const reader = new FileReader();
            reader.onload = async (event) => {
                const content = event.target.result;
                let cycleTime = null;
                let programName = null;

                if (file.type === 'application/pdf') {
                    // Handle PDF files
                    const text = await parsePDF(content);
                    ({ cycleTime, programName } = extractCycleTimeAndProgram(text));
                } else {
                    // Handle text files
                    ({ cycleTime, programName } = extractCycleTimeAndProgram(content));
                }

                // Add result to the table, now including program name
                results.push({ fileName: file.name, cycleTime, programName });
                const row = resultsTable.insertRow();
                row.insertCell().textContent = file.name;
                row.insertCell().textContent = programName || 'Not Found';
                row.insertCell().textContent = cycleTime || 'Not Found';
            };

            if (file.type === 'application/pdf') {
                reader.readAsArrayBuffer(file); // Read PDF as ArrayBuffer
            } else {
                reader.readAsText(file); // Read text files as text
            }
        }
        loading.textContent = 'Files processed!';
    } catch (error) {
        console.error("File processing failed:", error);
        alert('An error occurred while processing the files. Please try again.');
    } finally {
        loading.style.display = 'none';
    }
}

function extractCycleTimeAndProgram(text) {
    const lines = text.split('\n');
    let cycleTime = null;
    let programName = null;

    // Look for program name in the first few lines
    for (let i = 0; i < Math.min(5, lines.length); i++) {  // Check first 5 lines or less if the PDF has fewer lines
        if (lines[i].includes("Program Name") || lines[i].includes("Project Name")) { // Adjust based on what your PDFs might use
            programName = lines[i].replace(/Program Name:|Project Name:/i, '').trim();
            break;
        }
    }

    // Then look for the cycle time as before
    for (const line of lines) {
        if (line.includes("TOTAL CYCLE TIME") || line.includes("Subtotal")) {
            const regex = /(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)|(\$\d+(?:\.\d{2})?)/i;
            const match = line.match(regex);
            cycleTime = match ? match[0] : null;
            break; // Exit loop once cycle time is found
        }
    }

    return { cycleTime, programName };
}

// Parse PDF content using pdf.js
async function parsePDF(data) {
    const pdfjsLib = window['pdfjs-dist/build/pdf'];
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    const pdf = await pdfjsLib.getDocument({ data }).promise;
    let text = '';

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        text += textContent.items.map(item => item.str).join(' ') + '\n';
    }
    return text;
}

// Reset the results
function resetResults() {
    results = [];
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    resultsTable.innerHTML = '';
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadExcelInput').value = '';
}

// Update Excel with the summed cycle times and the net total cycle time
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

        // Create a new array for the updated rows
        let newExcelRows = [];

        // Calculate totals
        let cycleTimeSums = {};
        let totalCycleTime = 0;
        let subtotalSum = 0;

        results.forEach(result => {
            if (result.programName) {
                // Parse cycle time for each result
                let cycleTimeInSeconds = parseCycleTime(result.cycleTime) || 0; // Ensure we have a number
                totalCycleTime += cycleTimeInSeconds;

                // Handle subtotals
                const subtotalMatch = result.cycleTime.match(/\$\d+(?:\.\d{2})?/);
                if (subtotalMatch) {
                    subtotalSum += parseFloat(subtotalMatch[0].replace('$', ''));
                }

                // Sum up cycle time for each program
                cycleTimeSums[result.programName] = (cycleTimeSums[result.programName] || 0) + cycleTimeInSeconds;
            }
        });

        // Update rows with new data
        excelRows.forEach(row => {
            let newRow = [...row]; // Create a copy of the row
            const programName = row[1]?.toString().trim(); // Program Name in column B (index 1)

            if (cycleTimeSums[programName]) {
                newRow[3] = formatCycleTime(cycleTimeSums[programName]); // Total Cycle Time in column D (index 3)
            } else {
                // Keep original or set to empty string if there was no previous value
                newRow[3] = row[3] || '';
            }

            newExcelRows.push(newRow);
        });

        // Add summary rows
        newExcelRows.push([
            'Net Total Cycle Time', 
            '', 
            '', 
            formatCycleTime(totalCycleTime)
        ]);
        newExcelRows.push([
            'Sum of Subtotals', 
            '', 
            '', 
            `$${subtotalSum.toFixed(2)}`
        ]);

        // Create a new worksheet with the updated data
        const newWS = XLSX.utils.aoa_to_sheet(newExcelRows);
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, newWS, sheetName);

        // Save the new workbook with updated information
        XLSX.writeFile(newWB, 'new_updated_cycle_times.xlsx', {bookType:'xlsx', type: 'base64'});
    };

    reader.onerror = function(error) {
        console.error("Error reading file:", error);
        alert('An error occurred while reading the file. Please try again.');
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
