document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processButton').addEventListener('click', processFiles);
    document.getElementById('resetButton').addEventListener('click', resetResults);
    document.getElementById('uploadExcelButton').addEventListener('click', updateToExcel);
});

let results = []; // Store results for all files

async function processFiles() {
    const fileInput = document.getElementById('fileInput');
    const loading = document.getElementById('loading');
    const resultsTable = document.getElementById('resultsTable');
    const tbody = resultsTable ? resultsTable.querySelector('tbody') : null;

    if (!fileInput || !loading || !resultsTable || !tbody) {
        console.error('Required elements not found in the DOM');
        return;
    }

    if (fileInput.files.length === 0) {
        alert('Please select at least one file.');
        return;
    }

    results = [];
    tbody.innerHTML = '';
    loading.style.display = 'block';

    try {
        const files = Array.from(fileInput.files);
        for (const file of files) {
            if (file.type === 'application/pdf') {
                const content = await readFile(file);
                const text = await parsePDF(content);
                console.log('Full PDF Text:', text); // Log the entire text for debugging
                const projectNameLine = extractProjectNameLine(text);
                const cycleTime = extractCycleTime(text);
                console.log('Extracted Cycle Time:', cycleTime);
                if (projectNameLine || cycleTime) {
                    results.push({ fileName: file.name, projectNameLine, cycleTime });

                    const row = tbody.insertRow();
                    const cell1 = row.insertCell(0);
                    const cell2 = row.insertCell(1);
                    const cell3 = row.insertCell(2);
                    cell1.textContent = file.name;
                    cell2.textContent = projectNameLine || 'Not Found';
                    cell3.textContent = cycleTime || 'Not Found';
                } else {
                    console.log('Project name or cycle time not found for file:', file.name);
                    results.push({ fileName: file.name, projectNameLine: 'Not Found', cycleTime: 'Not Found' });

                    const row = tbody.insertRow();
                    const cell1 = row.insertCell(0);
                    const cell2 = row.insertCell(1);
                    const cell3 = row.insertCell(2);
                    cell1.textContent = file.name;
                    cell2.textContent = 'Not Found';
                    cell3.textContent = 'Not Found';
                }
            }
        }
    } catch (error) {
        console.error("Error processing files:", error);
        alert('An error occurred while processing the files.');
    } finally {
        loading.style.display = 'none';
        console.log("Processed results:", results);
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
    console.warn('Could not find "PROJECT NAME:" in the PDF');
    return null;
}

function extractCycleTime(text) {
    const lines = text.split('\n'); // Split text into lines
    for (const line of lines) {
        if (line.includes("TOTAL CYCLE TIME")) {
            // Use regex to extract the time part (e.g., "0 HOURS, 4 MINUTES, 16 SECONDS")
            const regex = /(\d+)\s*HOURS?,\s*(\d+)\s*MINUTES?,\s*(\d+)\s*SECONDS?/i;
            const match = line.match(regex);
            if (match) {
                return `${match[1]}h ${match[2]}m ${match[3]}s`;
            }
        }
    }
    return null; // Return null if no match is found
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

    console.log('File selected:', file.name, file.type);

    const reader = new FileReader();
    reader.onload = function(e) {
        console.log('File read completed');
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});

            console.log('Workbook parsed:', workbook);

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            let excelRows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

            console.log('Original Excel Rows:', excelRows);

            // Create a map of project names to cycle times
            const cycleTimeMap = new Map();
            results.forEach(result => {
                const projectName = result.projectNameLine ? result.projectNameLine.split(':')[1].trim() : null;
                if (!projectName) {
                    console.warn('Could not determine project name from:', result.fileName);
                    return;
                }

                const cycleTimeInSeconds = parseCycleTime(result.cycleTime);
                if (!isNaN(cycleTimeInSeconds)) {
                    cycleTimeMap.set(projectName, cycleTimeInSeconds);
                } else {
                    console.error('Could not parse cycle time for:', result.fileName);
                }
            });

            console.log("CycleTimeMap:", cycleTimeMap);

            // Update existing rows in Excel, adding cycle times to column D (index 3), matching with column B for 'Item No.'
            excelRows.forEach((row, rowIndex) => {
                const itemNo = row[1]?.toString().trim(); // 'Item No.' is in column B (index 1)
                if (cycleTimeMap.has(itemNo)) {
                    row[3] = formatCycleTime(cycleTimeMap.get(itemNo)); // Update cycle time in column D (index 3)
                    console.log(`Updated cycle time for ${itemNo}: ${row[3]}`);
                } else {
                    console.log(`No match found for Item No.: ${itemNo}`);
                }
            });

            // Add summary rows
            cycleTimeMap.forEach((cycleTime, project) => {
                excelRows.push([
                    '', // Empty for column A
                    `Subtotal for ${project}`, // Column B for 'Item No.'
                    '', 
                    formatCycleTime(cycleTime)
                ]);
            });

            // Add total cycle time at the end
            const totalCycleTime = Array.from(cycleTimeMap.values()).reduce((sum, time) => sum + time, 0);
            excelRows.push([
                '', // Empty for column A
                'Net Total Cycle Time', // Column B for 'Item No.'
                '', 
                formatCycleTime(totalCycleTime)
            ]);

            console.log("Updated Excel Rows:", excelRows);

            const newWS = XLSX.utils.aoa_to_sheet(excelRows);
            const newWB = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWB, newWS, sheetName);
            
            console.log('Attempting to save new Excel file');
            try {
                XLSX.writeFile(newWB, 'updated_cycle_times.xlsx');
                console.log('Excel sheet updated and saved.');
                alert('Excel file has been updated and saved as "updated_cycle_times.xlsx".');
            } catch (saveError) {
                console.error('Error saving Excel file:', saveError);
                alert('An error occurred while saving the Excel file. Check console for details.');
            }
        } catch (readError) {
            console.error('Error reading or parsing Excel file:', readError);
            alert('Failed to read or parse the Excel file. Please check if it\'s a valid Excel document.');
            return;
        }
    };
    reader.onerror = function(error) {
        console.error('Error reading file:', error);
        alert('There was an error reading the file. Please try again with a different file.');
    };
    reader.readAsArrayBuffer(file);
}

function parseCycleTime(cycleTimeString) {
    if (!cycleTimeString) return 0;

    const regex = /(\d+)h\s*(\d+)m\s*(\d+)s/;
    const match = cycleTimeString.match(regex);
    if (match) {
        const hours = parseInt(match[1]);
        const minutes = parseInt(match[2]);
        const seconds = parseInt(match[3]);
        return (hours * 3600) + (minutes * 60) + seconds;
    }
    return 0; // If format doesn't match, return 0
}

function formatCycleTime(totalSeconds) {
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;

    return `${hours}h ${minutes}m ${seconds}s`;
}
