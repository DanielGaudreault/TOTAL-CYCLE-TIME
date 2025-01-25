function processFiles() {
    const excelFile = document.getElementById('excel').files[0];
    const pdfFiles = document.getElementById('pdf').files;
    const results = document.getElementById('results');

    if (!excelFile || pdfFiles.length === 0) {
        results.innerHTML = '<p>Please upload both Excel and PDF files.</p>';
        return;
    }

    // Read Excel file
    const excelReader = new FileReader();
    excelReader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert Excel sheet to JSON
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Add TOTAL CYCLE TIME as a new column
        json[0].push('TOTAL CYCLE TIME'); // Add header

        // Process each PDF file
        let processedCount = 0;
        for (let i = 0; i < pdfFiles.length; i++) {
            const pdfFile = pdfFiles[i];
            const pdfReader = new FileReader();
            pdfReader.onload = function (event) {
                const pdfData = new Uint8Array(event.target.result);
                parsePDF(pdfData).then(pdfText => {
                    // Extract TOTAL CYCLE TIME from PDF text
                    const cycleTime = extractCycleTime(pdfText);

                    if (cycleTime) {
                        // Add cycle time to each row
                        for (let j = 1; j < json.length; j++) {
                            json[j].push(cycleTime);
                        }
                    }

                    processedCount++;
                    if (processedCount === pdfFiles.length) {
                        // Convert JSON back to Excel
                        const updatedWorksheet = XLSX.utils.aoa_to_sheet(json);
                        const updatedWorkbook = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(updatedWorkbook, updatedWorksheet, 'Sheet1');

                        // Download the updated Excel file
                        XLSX.writeFile(updatedWorkbook, 'updated_excel.xlsx');
                        results.innerHTML = `<p>Excel file updated successfully! "TOTAL CYCLE TIME" added from ${pdfFiles.length} PDF(s).</p>`;
                    }
                });
            };
            pdfReader.readAsArrayBuffer(pdfFile);
        }
    };
    excelReader.readAsArrayBuffer(excelFile);
}

function extractCycleTime(text) {
    // Use regex to find "TOTAL CYCLE TIME" and extract the time
    const regex = /TOTAL CYCLE TIME\s*:\s*(\d+\s*HOURS?,\s*\d+\s*MINUTES?,\s*\d+\s*SECONDS?)/i;
    const match = text.match(regex);
    return match ? match[1] : null; // Return the matched time or null
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
