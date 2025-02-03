document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processButton').addEventListener('click', processFiles);
    document.getElementById('resetButton').addEventListener('click', resetResults);
    document.getElementById('uploadExcelButton').addEventListener('click', updateToExcel);
});

let totalCycleTime = { hours: 0, minutes: 0, seconds: 0 };
let cycleTimesPerItem = {}; // To store total cycle time per item

async function processFiles() {
    const fileInput = document.getElementById('fileInput');
    const loading = document.getElementById('loading');
    const resultsTable = document.getElementById('resultsTable');
    const tbody = resultsTable.querySelector('tbody');

    if (fileInput.files.length === 0) {
        alert('Please select at least one file.');
        return;
    }

    results = [];
    totalCycleTime = { hours: 0, minutes: 0, seconds: 0 }; // Reset total cycle time
    cycleTimesPerItem = {}; // Reset cycle times per item map
    tbody.innerHTML = '';
    loading.style.display = 'block';
    loading.textContent = `Processing ${fileInput.files.length} files...`;

    try {
        const files = Array.from(fileInput.files);
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (file.type === 'application/pdf') {
                loading.textContent = `Processing file ${i + 1} of ${files.length}...`;
                
                const content = await readFile(file);
                const text = await parsePDF(content);
                const projectName = extractProjectNameLine(text);
                const cycleTime = extractCycleTime(text);
                if (projectName && cycleTime) {
                    let cleanProjectName = projectName.split(':')[1].trim().replace(/R\d+/g, '').trim();
                    
                    if (!cleanProjectName.includes(',')) {
                        results.push({ projectName: cleanProjectName, cycleTime, fileName: file.name });

                        const timeParts = cycleTime.split(' ');
                        const hours = parseInt(timeParts[0].replace('h', ''), 10);
                        const minutes = parseInt(timeParts[1].replace('m', ''), 10);
                        const seconds = parseInt(timeParts[2].replace('s', ''), 10);

                        console.log(`Parsed cycle time from ${file.name}: ${hours}h ${minutes}m ${seconds}s`);

                        if (!cycleTimesPerItem[cleanProjectName]) {
                            cycleTimesPerItem[cleanProjectName] = { hours: 0, minutes: 0, seconds: 0 };
                        }

                        cycleTimesPerItem[cleanProjectName].hours += hours;
                        cycleTimesPerItem[cleanProjectName].minutes += minutes;
                        cycleTimesPerItem[cleanProjectName].seconds += seconds;

                        // Handle overflow for seconds
                        if (cycleTimesPerItem[cleanProjectName].seconds >= 60) {
                            cycleTimesPerItem[cleanProjectName].minutes += Math.floor(cycleTimesPerItem[cleanProjectName].seconds / 60);
                            cycleTimesPerItem[cleanProjectName].seconds %= 60;
                        }

                        // Handle overflow for minutes
                        if (cycleTimesPerItem[cleanProjectName].minutes >= 60) {
                            cycleTimesPerItem[cleanProjectName].hours += Math.floor(cycleTimesPerItem[cleanProjectName].minutes / 60);
                            cycleTimesPerItem[cleanProjectName].minutes %= 60;
                        }

                        const row = tbody.insertRow();
                        row.insertCell().textContent = file.name;
                        row.insertCell().textContent = cleanProjectName;
                        row.insertCell().textContent = cycleTime;
                    } else {
                        console.log(`Skipping project name with comma: ${cleanProjectName}`);
                    }
                }
            }
        }

        console.log(`Total accumulated cycle time: ${totalCycleTime.hours}h ${totalCycleTime.minutes}m ${totalCycleTime.seconds}s`);
        loading.textContent = "Processing complete!";

    } catch (error) {
        console.error("Error processing files:", error);
        alert('An error occurred while processing the files.');
    } finally {
        loading.style.display = 'none';
    }
}

// The rest of your functions (readFile, parsePDF, extractProjectNameLine, extractCycleTime, resetResults, updateToExcel) remain the same.

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = event => resolve(event.target.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
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
                        text += pageText + '\n';
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

function extractProjectNameLine(text) {
    const regex = /PROJECT NAME:\s*(.*?)\s*DATE:/i;
    const lines = text.split('\n');
    for (let line of lines) {
        const match = line.match(regex);
        if (match && match[1].trim()) {
            return `PROJECT NAME: ${match[1].trim()}`;
        }
    }
    return null;
}

function extractCycleTime(text) {
    const regex = /TOTAL CYCLE TIME:\s*(\d+)\s*HOURS?,\s*(\d+)\s*MINUTES?,\s*(\d+)\s*SECONDS?/i;
    const lines = text.split('\n');
    for (const line of lines) {
        const match = line.match(regex);
        if (match) {
            return `${match[1]}h ${match[2]}m ${match[3]}s`;
        }
    }
    return null;
}

function resetResults() {
    results = [];
    const resultsTable = document.getElementById('resultsTable');
    if (resultsTable) {
        resultsTable.querySelector('tbody').innerHTML = '';
    }
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadExcelInput').value = '';
}

function updateToExcel() {
    const fileInput = document.getElementById('uploadExcelInput');
    const file = fileInput.files[0];

    if (!file) {
        alert('Please select an Excel file to update.');
        return;
    }

    console.log('Excel file selected:', file.name);

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            // Read the Excel file
            console.log('Reading Excel file...');
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            let excelRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            console.log('Excel rows before update:', excelRows);
            console.log('Results data (from PDFs):', results); // Check what we are comparing with

            // Loop through each row in the Excel sheet
            for (let i = 0; i < excelRows.length; i++) {
                const row = excelRows[i];
                const itemNo = (row[1] || '').toString().trim(); // Assuming Item No is in column B (index 1)
                console.log(`Processing row ${i + 1}: Item No. "${itemNo}"`);

                // Normalize Item No by removing "R" followed by digits (R0, R1, R2, etc.)
                const normalizedItemNo = itemNo.replace(/[^a-zA-Z0-9]/g, '').toLowerCase().replace(/r\d+/g, '');
                console.log(`Normalized Item No: "${normalizedItemNo}"`);

                let matchFound = false;

                // Loop through the results array (PDFs data)
                for (const projectName in cycleTimesPerItem) {
                    // Normalize project name to match
                    const normalizedProjectName = projectName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase().replace(/r\d+/g, '');
                    console.log(`Checking if "${normalizedItemNo}" matches with "${normalizedProjectName}"`);

                    if (normalizedItemNo === normalizedProjectName) {
                        console.log(`Match found! Updating Cycle Time: "${cycleTimesPerItem[projectName].hours}h ${cycleTimesPerItem[projectName].minutes}m ${cycleTimesPerItem[projectName].seconds}s"`);
                        row[3] = `${cycleTimesPerItem[projectName].hours}h ${cycleTimesPerItem[projectName].minutes}m ${cycleTimesPerItem[projectName].seconds}s`; // Update Column D
                        matchFound = true;
                    }
                }

                if (!matchFound) {
                    console.log(`No match found for Item No. "${itemNo}"`);
                }
            }

            // After processing all rows, create the updated worksheet
            const newWS = XLSX.utils.aoa_to_sheet(excelRows);
            const newWB = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWB, newWS, sheetName);

            // Generate the updated Excel file for download
            const wbout = XLSX.write(newWB, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/octet-stream' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            document.body.appendChild(a);
            a.href = url;
            a.download = 'updated_cycle_times.xlsx';
            a.click();
            setTimeout(() => URL.revokeObjectURL(url), 100);

            alert('Excel sheet has been updated. Download will start automatically.');

        } catch (error) {
            console.error('Error updating Excel:', error);
            alert('An error occurred while updating the Excel file. Check console for details.');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Function to add two time objects (hours, minutes, seconds)
function addCycleTimes(time1, time2) {
    const totalSeconds = (time1.hours + time2.hours) * 3600 + (time1.minutes + time2.minutes) * 60 + (time1.seconds + time2.seconds);
    console.log(`Adding times: ${time1.hours}h ${time1.minutes}m ${time1.seconds}s + ${time2.hours}h ${time2.minutes}m ${time2.seconds}s = ${totalSeconds} seconds`);
    return {
        hours: Math.floor(totalSeconds / 3600),
        minutes: Math.floor((totalSeconds % 3600) / 60),
        seconds: totalSeconds % 60
    };
}
