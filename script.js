document.addEventListener('DOMContentLoaded', () => {
    const processButton = document.getElementById('processButton');
    const resetButton = document.getElementById('resetButton');
    const uploadExcelButton = document.getElementById('uploadExcelButton');

    if (processButton) {
        processButton.addEventListener('click', processFiles);
        console.log('Process button event listener attached');
    } else {
        console.error('Process button not found in DOM');
    }

    if (resetButton) {
        resetButton.addEventListener('click', resetResults);
        console.log('Reset button event listener attached');
    } else {
        console.error('Reset button not found in DOM');
    }

    if (uploadExcelButton) {
        uploadExcelButton.addEventListener('click', updateToExcel);
        console.log('Upload Excel button event listener attached');
    } else {
        console.error('Upload Excel button not found in DOM');
    }
});

let results = [];

async function processFiles() {
    console.log('processFiles function called');
    const fileInput = document.getElementById('fileInput');
    const loading = document.getElementById('loading');
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];

    if (!fileInput || !loading || !resultsTable) {
        console.error('One or more required elements not found in DOM');
        return;
    }

    if (fileInput.files.length === 0) {
        alert('Please select at least one file.');
        return;
    }

    results = [];
    resultsTable.innerHTML = '';
    loading.style.display = 'block';

    try {
        const files = Array.from(fileInput.files);
        for (const file of files) {
            const content = await readFile(file);
            let cycleTime = null;

            if (file.type === 'application/pdf') {
                const text = await parsePDF(content);
                cycleTime = extractCycleTime(text);
            } else {
                cycleTime = extractCycleTime(content);
            }

            results.push({ fileName: file.name, cycleTime });
            const row = resultsTable.insertRow();
            row.insertCell().textContent = file.name;
            row.insertCell().textContent = cycleTime || 'Not Found';
            console.log(`Processed file: ${file.name}, Cycle Time: ${cycleTime}`);
        }
    } catch (error) {
        console.error("Error processing files:", error);
        alert('An error occurred while processing the files. Please try again or check the console for details.');
    } finally {
        loading.style.display = 'none';
        console.log("Results after processing:", results);
    }
}

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = event => resolve(event.target.result);
        reader.onerror = reject;
        file.type === 'application/pdf' ? reader.readAsArrayBuffer(file) : reader.readAsText(file);
    });
}

function extractCycleTime(text) {
    const lines = text.split('\n');
    for (const line of lines) {
        if (line.includes("TOTAL CYCLE TIME")) {
            const regex = /(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)/i;
            const match = line.match(regex);
            return match ? match[0] : null;
        }
    }
    return null;
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

function resetResults() {
    console.log('resetResults function called');
    results = [];
    const resultsTable = document.getElementById('resultsTable');
    if (resultsTable) {
        resultsTable.getElementsByTagName('tbody')[0].innerHTML = '';
    }
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadExcelInput').value = '';
}

function updateToExcel() {
    console.log('updateToExcel function called');
    const fileInput = document.getElementById('uploadExcelInput');
    const file = fileInput ? fileInput.files[0] : null;

    if (!file) {
        alert('Please select an Excel file to update.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        console.log('File read successfully');
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        console.log('Workbook parsed:', workbook);

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        let excelRows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

        console.log("Original Excel Rows:", excelRows);
        console.log('Results for updating:', results);

        let cycleTimeSums = {};
        let totalCycleTime = 0;

        results.forEach(result => {
            const matchIdentifier = result.fileName.match(/^(.+?) - \d+/)?.[1].trim();
            if (!matchIdentifier) {
                console.error('Could not extract project name from PDF file name:', result.fileName);
                return;
            }

            const cycleTimeInSeconds = parseCycleTime(result.cycleTime);
            if (!isNaN(cycleTimeInSeconds)) {
                totalCycleTime += cycleTimeInSeconds;
                cycleTimeSums[matchIdentifier] = (cycleTimeSums[matchIdentifier] || 0) + cycleTimeInSeconds;
                console.log(`Added ${cycleTimeInSeconds} seconds to ${matchIdentifier}.`);
            } else {
                console.error('Could not parse cycle time for:', result.fileName);
            }
        });

        console.log("Cycle Time Sums:", cycleTimeSums);
        console.log("Total Cycle Time:", totalCycleTime);

        // Adding subtotals for each project
        Object.keys(cycleTimeSums).forEach(project => {
            excelRows.push([
                `Subtotal for ${project}`,
                '', 
                '', 
                formatCycleTime(cycleTimeSums[project])
            ]);
            console.log(`Added subtotal for ${project}: ${formatCycleTime(cycleTimeSums[project])}`);
        });

        // Adding the total cycle time at the very end
        excelRows.push([
            'Net Total Cycle Time', 
            '', 
            '', 
            formatCycleTime(totalCycleTime)
        ]);
        console.log(`Added net total cycle time: ${formatCycleTime(totalCycleTime)}`);

        console.log('Updated Excel Rows:', excelRows);

        const newWS = XLSX.utils.aoa_to_sheet(excelRows);
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, newWS, sheetName);

        XLSX.writeFile(newWB, 'updated_cycle_times.xlsx');
        console.log('New Excel file created and saved');
    };
    reader.readAsArrayBuffer(file);
}

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

function formatCycleTime(totalSeconds) {
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;

    return `${hours} HOURS, ${minutes} MINUTES, ${seconds} SECONDS`;
}
