document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processButton').addEventListener('click', processFiles);
    document.getElementById('resetButton').addEventListener('click', resetResults);
    document.getElementById('uploadExcelButton').addEventListener('click', updateToExcel);
});

let results = [];

async function processFiles() {
    const fileInput = document.getElementById('fileInput');
    const loading = document.getElementById('loading');
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];

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
    results = [];
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    resultsTable.innerHTML = '';
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
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        let excel
