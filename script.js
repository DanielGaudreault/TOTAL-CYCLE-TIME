let results = []; // Store results for all files

function processFiles() {
    const fileInput = document.getElementById('fileInput');
    const loading = document.getElementById('loading');
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    const downloadButton = document.getElementById('downloadButton');

    if (fileInput.files.length === 0) {
        alert('Please select at least one file.');
        return;
    }

    // Clear previous results
    results = [];
    resultsTable.innerHTML = '';
    downloadButton.style.display = 'none';
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
        downloadButton.style.display = 'block';
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

function resetResults() {
    results = []; // Clear the results array
    const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
    resultsTable.innerHTML = ''; // Clear the table
    document.getElementById('downloadButton').style.display = 'none'; // Hide download button
}

function downloadResults() {
    // Convert results to Excel
    const ws = XLSX.utils.json_to_sheet(results);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Results');
    XLSX.writeFile(wb, 'cycle_times.xlsx');
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

        // Assuming data is in the first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});

        // Update existing results with new data from Excel
        jsonData.slice(1).forEach(row => {
            if (row.length > 0) {
                const existingResult = results.find(r => r.fileName === row[0]);
                if (existingResult) {
                    existingResult.cycleTime = row[1] || existingResult.cycleTime; // Update if new cycleTime is provided
                } else {
                    // If it's a new file, add to results
                    results.push({ fileName: row[0], cycleTime: row[1] || 'Not Found' });
                }
            }
        });

        // Update the table
        const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
        resultsTable.innerHTML = '';
        results.forEach(result => {
            const newRow = resultsTable.insertRow();
            newRow.insertCell().textContent = result.fileName;
            newRow.insertCell().textContent = result.cycleTime || 'Not Found';
        });

        document.getElementById('downloadButton').style.display = 'block';
    };
    reader.readAsArrayBuffer(file);
}

// Event Listeners
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processButton').addEventListener('click', processFiles);
    document.getElementById('resetButton').addEventListener('click', resetResults);
    document.getElementById('downloadButton').addEventListener('click', downloadResults);
    document.getElementById('uploadExcelButton').addEventListener('click', updateToExcel);
});
