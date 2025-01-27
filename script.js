document.getElementById('upload-form').addEventListener('submit', function(event) {
    event.preventDefault();

    const excelFiles = document.getElementById('excel-files').files;
    const pdfFiles = document.getElementById('pdf-files').files;
    const formData = new FormData();

    for (let i = 0; i < excelFiles.length; i++) {
        formData.append('excel-files', excelFiles[i]);
    }

    for (let i = 0; i < pdfFiles.length; i++) {
        formData.append('pdf-files', pdfFiles[i]);
    }

    fetch('/process-files', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        document.getElementById('results').textContent = 'Files processed successfully!';
        console.log(data);
    })
    .catch(error => {
        document.getElementById('results').textContent = 'Error processing files.';
        console.error('Error:', error);
    });
});

function extractCycleTime(text) {
    const lines = text.split('\n'); // Split text into lines
    for (const line of lines) {
        if (line.includes("TOTAL CYCLE TIME")) {
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

function searchFile() {
    const fileInput = document.getElementById('pdf-files');
    const results = document.getElementById('results');
    results.textContent = ''; // Clear previous results

    if (fileInput.files.length === 0) {
        alert('Please select a file.');
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function (event) {
        const content = event.target.result;

        if (file.type === 'application/pdf') {
            parsePDF(content).then(text => {
                const cycleTime = extractCycleTime(text);
                displayResults(cycleTime);
            });
        } else {
            const cycleTime = extractCycleTime(content);
            displayResults(cycleTime);
        }
    };

    if (file.type === 'application/pdf') {
        reader.readAsArrayBuffer(file); // Read PDF as ArrayBuffer
    } else {
        reader.readAsText(file); // Read text files as text
    }
}

function displayResults(cycleTime) {
    const results = document.getElementById('results');
    if (cycleTime) {
        results.textContent = `Cycle Time: ${cycleTime}`;
    } else {
        results.textContent = 'No instances of "TOTAL CYCLE TIME" found.';
    }
}
