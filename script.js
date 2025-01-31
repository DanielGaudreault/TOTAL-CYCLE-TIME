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
                    const cleanProjectName = projectName.split(':')[1].trim();
                    results.push({ projectName: cleanProjectName, cycleTime });
                    const row = tbody.insertRow();
                    row.insertCell().textContent = file.name;
                    row.insertCell().textContent = cleanProjectName;
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
    // ... (previous code remains the same)

    reader.onload = function(e) {
        try {
            // ... (Excel reading and initial console logs)

            for (let i = 0; i < excelRows.length; i++) {
                const row = excelRows[i];
                let itemNo = (row[1] || '').toString().trim();
                let normalizedItemNo = itemNo.replace(/[\s\-\._]/g, '').toLowerCase();
                console.log(`Checking for match with normalized Item No from Excel: "${normalizedItemNo}"`);

                let match = results.find(result => {
                    let normalizedProjectName = result.projectName.replace(/[\s\-\._]/g, '').toLowerCase();
                    return normalizedProjectName === normalizedItemNo;
                });

                if (match) {
                    console.log(`Match found for Item No: ${itemNo}. Updating with cycle time: ${match.cycleTime}`);
                    row[3] = match.cycleTime;
                    console.log(`Updated row ${i + 1}:`, row);
                } else {
                    console.log(`No match found for Item No: ${itemNo}`);
                }
            }

            // ... (convert to worksheet)

            // Instead of writing directly to file:
            const wbout = XLSX.write(newWB, { bookType:'xlsx', type: 'array' });
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
