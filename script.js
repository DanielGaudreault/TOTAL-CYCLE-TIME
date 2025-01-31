document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processButton').addEventListener('click', processFiles);
    document.getElementById('resetButton').addEventListener('click', resetResults);
    document.getElementById('uploadExcelButton').addEventListener('click', updateToExcel);
});

let totalCycleTime = { hours: 0, minutes: 0, seconds: 0 };

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
    tbody.innerHTML = '';
    loading.style.display = 'block';

    try {
        const files = Array.from(fileInput.files);
        for (const file of files) {
            if (file.type === 'application/pdf') {
                const content = await readFile(file);
                const text = await parsePDF(content);
                const projectName = extractProjectNameLine(text);
                const cycleTime = extractCycleTime(text);
                if (projectName && cycleTime) {
                    // Clean the project name by removing all "R" followed by digits (e.g., "R1", "R2", "R3")
                    const cleanProjectName = projectName.split(':')[1].trim().replace(/R\d+/g, '').trim();
                    results.push({ projectName: cleanProjectName, cycleTime });

                    // Parse cycle time (hours, minutes, seconds)
                    const timeParts = cycleTime.split(' ');
                    const hours = parseInt(timeParts[0].replace('h', ''), 10);
                    const minutes = parseInt(timeParts[1].replace('m', ''), 10);
                    const seconds = parseInt(timeParts[2].replace('s', ''), 10);

                    console.log(`Processing cycle time from ${file.name}: ${hours}h ${minutes}m ${seconds}s`);

                    // Add to the total cycle time
                    totalCycleTime.hours += hours;
                    totalCycleTime.minutes += minutes;
                    totalCycleTime.seconds += seconds;

                    // Handle overflow for seconds
                    if (totalCycleTime.seconds >= 60) {
                        totalCycleTime.minutes += Math.floor(totalCycleTime.seconds / 60);
                        totalCycleTime.seconds = totalCycleTime.seconds % 60;
                    }

                    // Handle overflow for minutes
                    if (totalCycleTime.minutes >= 60) {
                        totalCycleTime.hours += Math.floor(totalCycleTime.minutes / 60);
                        totalCycleTime.minutes = totalCycleTime.minutes % 60;
                    }

                    // Add a row for each PDF processed
                    const row = tbody.insertRow();
                    row.insertCell().textContent = file.name;
                    row.insertCell().textContent = cleanProjectName;
                    row.insertCell().textContent = cycleTime;
                }
            }
        }

        console.log(`Total accumulated cycle time: ${totalCycleTime.hours}h ${totalCycleTime.minutes}m ${totalCycleTime.seconds}s`);

    } catch (error) {
        console.error("Error processing files:", error);
        alert('An error occurred while processing the files.');
    } finally {
        loading.style.display = 'none';
    }
}

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
                results.forEach((result) => {
                    // Clean the project name by removing "R0", "R1", "R2", etc.
                    const normalizedProjectName = result.projectName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase().replace(/r\d+/g, '');
                    console.log(`Checking if "${normalizedItemNo}" matches with "${normalizedProjectName}"`);

                    if (normalizedItemNo === normalizedProjectName) {
                        console.log(`Match found! Updating Cycle Time: "${result.cycleTime}"`);
                        row[3] = result.cycleTime; // Update Column D (index 3) with cycle time
                        matchFound = true;
                    }
                });

                if (!matchFound) {
                    console.log(`No match found for Item No. "${itemNo}"`);
                }
            }

            // After processing all rows, add the total cycle time in column D of the last row
            const totalCycleTimeString = `${totalCycleTime.hours}h ${totalCycleTime.minutes}m ${totalCycleTime.seconds}s`;

            // Append the total cycle time in the last row of column D
            excelRows.push(['', '', '', totalCycleTimeString]);  // Add the net total cycle time

            console.log('Excel rows after update:', excelRows);

            // Create the updated worksheet
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

function addCycleTimes(time1, time2) {
    const [h1, m1, s1] = time1.split('h ')[0].split('m ')[0].split('s').map(Number);
    const [h2, m2, s2] = time2.split('h ')[0].split('m ')[0].split('s').map(Number);
    const totalSeconds = (h1 + h2) * 3600 + (m1 + m2) * 60 + (s1 + s2);
    return `${Math.floor(totalSeconds / 3600)}h ${Math.floor((totalSeconds % 3600) / 60)}m ${totalSeconds % 60}s`;
}
