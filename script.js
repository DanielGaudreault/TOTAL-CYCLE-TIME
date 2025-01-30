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
    const tbody = resultsTable.querySelector('tbody');

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
                const projectName = extractProjectNameLine(text);
                const cycleTime = extractCycleTime(text);
                if (projectName && cycleTime) {
                    results.push({ projectName: projectName.split(':')[1].trim(), cycleTime });
                    const row = tbody.insertRow();
                    row.insertCell().textContent = file.name;
                    row.insertCell().textContent = projectName;
                    row.insertCell().textContent = cycleTime;
                }
            }
        }
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

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            let excelRows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

            // Update existing rows in Excel, matching with column B for 'Item No.'
            results.forEach(result => {
                let updated = false;
                for (let row of excelRows) {
                    const itemNo = row[1]?.toString().trim(); // 'Item No.' is in column B (index 1)
                    if (itemNo === result.projectName) {
                        row[3] = result.cycleTime; // Update cycle time in column D (index 3)
                        console.log(`Updated cycle time for ${itemNo}: ${result.cycleTime}`);
                        updated = true;
                        break; // Only update the first match
                    }
                }
                if (!updated) {
                    console.log(`No match found for project name: ${result.projectName}`);
                }
            });

            const newWS = XLSX.utils.aoa_to_sheet(excelRows);
            const newWB = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWB, newWS, sheetName);
            XLSX.writeFile(newWB, 'updated_cycle_times.xlsx');
            alert('Excel sheet has been updated with new cycle times.');
        } catch (error) {
            console.error('Error updating Excel:', error);
            alert('An error occurred while updating the Excel file. Check console for details.');
        }
    };
    reader.readAsArrayBuffer(file);
}
