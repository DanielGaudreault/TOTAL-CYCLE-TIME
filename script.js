document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processButton').addEventListener('click', processFiles);
    document.getElementById('resetButton').addEventListener('click', resetResults);
});

let results = [];

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
                if (projectNameLine) {
                    console.log('Found project name line:', projectNameLine);
                    results.push({ fileName: file.name, projectNameLine });

                    const row = tbody.insertRow();
                    const cell1 = row.insertCell(0);
                    const cell2 = row.insertCell(1);
                    cell1.textContent = file.name;
                    cell2.textContent = projectNameLine;
                } else {
                    console.log('Project name line not found for file:', file.name);
                    results.push({ fileName: file.name, projectNameLine: 'Not Found' });

                    const row = tbody.insertRow();
                    const cell1 = row.insertCell(0);
                    const cell2 = row.insertCell(1);
                    cell1.textContent = file.name;
                    cell2.textContent = 'Not Found';
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
    // Specific regex for your format
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

function resetResults() {
    results = [];
    const resultsTable = document.getElementById('resultsTable');
    if (resultsTable) {
        resultsTable.querySelector('tbody').innerHTML = '';
    }
    document.getElementById('fileInput').value = '';
}
