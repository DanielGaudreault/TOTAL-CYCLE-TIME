document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM fully loaded and parsed');

    // Test button click
    document.getElementById('processButton').addEventListener('click', () => {
        console.log('Process Button Clicked');
        processFiles();
    });

    document.getElementById('resetButton').addEventListener('click', () => {
        console.log('Reset Button Clicked');
        resetResults();
    });

    document.getElementById('uploadExcelButton').addEventListener('click', () => {
        console.log('Upload Excel Button Clicked');
        updateToExcel();
    });
});

let results = []; // Store results for all files

// Process selected files (PDFs)
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
                try {
                    const content = event.target.result;
                    const text = await parsePDF(content);
                    const cycleTime = extractCycleTime(text);

                    // Add result to the table
                    results.push({ fileName: file.name, cycleTime });
                    const row = resultsTable.insertRow();
                    row.insertCell().textContent = file.name;
                    row.insertCell().textContent = cycleTime || 'Not Found';

                    resolve();
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = (error) => {
                reject(error);
            };
            reader.readAsArrayBuffer(file); // Read PDF as ArrayBuffer
        });
    });

    Promise.all(promises)
        .then(() => {
            loading.style.display = 'none';
        })
        .catch(error => {
            console.error("File processing failed:", error);
            loading.style.display = 'none';
            alert('An error occurred while processing the files. Check the console for details.');
        });
}

// Extract cycle time from text
function extractCycleTime(text) {
    const lines = text.split('\n'); // Split text into lines
    for (const line of lines) {
        if (line.includes("TOTAL CYCLE TIME")) {
            const regex = /(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)/i;
            const match = line.match(regex);
            return match ? match[0] : null;
        }
    }
    return null;
}

// Parse PDF content using pdf.js
function parsePDF(data) {
    return new Promise((resolve, reject) => {
        const loadingTask = pdfjsLib.getDocument({ data });
        loadingTask.promise
            .then(pdf => {
                let text = '';
                const numPages = pdf.numPages;

                const fetchPage = async (pageNum) => {
                    const page = await pdf.getPage(pageNum);
                    const textContent = await page.getTextContent();
                    textContent.items.forEach(item => {
                        text += item.str + ' ';
                    });
                    text += '\n'; // Add newline after each page
                };

                const fetchAllPages = async () => {
                    for (let i = 1; i <= numPages; i++) {
                        await fetchPage(i);
                    }
                    resolve(text);
                };

                fetchAllPages();
            })
            .catch(reject);
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

// Update Excel with the summed cycle times and the net total cycle time
function updateToExcel() {
    const fileInput = document.getElementById('uploadExcelInput');
    const file = fileInput.files[0];

    if (!file) {
        alert('Please select an Excel file to update.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Assuming data is in the first sheet
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            let excelRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Create an object to store the total cycle time for each project
            let cycleTimeSums = {};
            let totalCycleTime = 0; // Variable to store the total cycle time for all PDFs

            // Process each result from the PDFs
            results.forEach(result => {
                const matchIdentifier = result.fileName.match(/^(.+?) - \d+/)?.[1].trim(); // Adjusting regex
                if (!matchIdentifier) {
                    console.error('Could not extract match identifier from PDF file name:', result.fileName);
                    return;
                }

                const cycleTimeInSeconds = parseCycleTime(result.cycleTime);
                totalCycleTime += cycleTimeInSeconds;

                if (cycleTimeSums[matchIdentifier]) {
                    cycleTimeSums[matchIdentifier] += cycleTimeInSeconds;
                } else {
                    cycleTimeSums[matchIdentifier] = cycleTimeInSeconds;
                }
            });

            // Update Excel rows with cycle time sums
            excelRows.forEach((row, rowIndex) => {
                const itemNo = row[0]?.toString().trim(); // Assuming 'Item No.' is in the first column
                if (cycleTimeSums[itemNo]) {
                    row[3] = formatCycleTime(cycleTimeSums[itemNo]); // Assuming 'Total Cycle Time' is in column 4 (index 3)
                }
            });

            // Add a row at the end for the net total cycle time
            excelRows.push([
                'Net Total Cycle Time', // First column
                '', // Leave second column empty
                '', // Leave third column empty
                formatCycleTime(totalCycleTime) // Total cycle time in the fourth column
            ]);

            // Convert the updated data back to worksheet format
            const newWS = XLSX.utils.aoa_to_sheet(excelRows);
            const newWB = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWB, newWS, sheetName);

            // Save the new workbook
            XLSX.writeFile(newWB, 'updated_cycle_times.xlsx');
        } catch (error) {
            console.error("Error updating Excel:", error);
            alert('An error occurred while updating the Excel file. Check the console for details.');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Helper function to parse cycle time (e.g., "1 HOURS, 30 MINUTES, 0 SECONDS" -> total seconds)
function parseCycleTime(cycleTimeString) {
    if (!cycleTimeString) return 0;

    const regexHours = /(\d+)\s*HOURS?/, hoursMatch = cycleTimeString.match(regexHours);
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
