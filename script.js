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

        // Find the index of the "Item Number" column
        const itemNumberColumnIndex = json[0].indexOf('Item Number');
        if (itemNumberColumnIndex === -1) {
            results.innerHTML = '<p>Excel file must have an "Item Number" column.</p>';
            return;
        }

        // Add "TOTAL CYCLE TIME" as a new column if it doesn't exist
        const cycleTimeColumnIndex = json[0].indexOf('TOTAL CYCLE TIME');
        if (cycleTimeColumnIndex === -1) {
            json[0].push('TOTAL CYCLE TIME'); // Add header
        }

        // Process each PDF file
        let processedCount = 0;
        results.innerHTML = `<p>Processing ${pdfFiles.length} files... Please wait.</p>`;

        const processNextPdf = (index) => {
            if (index >= pdfFiles.length) {
                // All PDFs processed, update Excel
                const updatedWorksheet = XLSX.utils.aoa_to_sheet(json);
                const updatedWorkbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(updatedWorkbook, updatedWorksheet, 'Sheet1');

                // Download the updated Excel file
                XLSX.writeFile(updatedWorkbook, 'updated_excel.xlsx');
                results.innerHTML = `<p>Excel file updated successfully! "TOTAL CYCLE TIME" added for ${processedCount} PDF(s).</p>`;
                return;
            }

            const pdfFile = pdfFiles[index];
            if (!pdfFile.name.endsWith('.pdf')) {
                // Skip non-PDF files
                processedCount++;
                processNextPdf(index + 1);
                return;
            }

            const pdfReader = new FileReader();
            pdfReader.onload = function (event) {
                const pdfData = new Uint8Array(event.target.result);
                parsePDF(pdfData).then(pdfText => {
                    // Extract Item Number and TOTAL CYCLE TIME from PDF text
                    const itemNumber = extractItemNumber(pdfText);
                    const cycleTime = extractCycleTime(pdfText);

                    if (itemNumber && cycleTime) {
                        // Find the row in the Excel file that matches the item number
                        let rowUpdated = false;

                        for (let i = 1; i < json.length; i++) {
                            if (json[i][itemNumberColumnIndex] === itemNumber) {
                                // Update the "TOTAL CYCLE TIME" column for this row
                                if (cycleTimeColumnIndex === -1) {
                                    json[i].push(cycleTime); // Add to new column
                                } else {
                                    json[i][cycleTimeColumnIndex] = cycleTime; // Update existing column
                                }
                                rowUpdated = true;
                                break;
                            }
                        }

                        if (!rowUpdated) {
                            console.warn(`No matching item number found for PDF file: ${pdfFile.name}`);
                        }
                    } else {
                        console.warn(`Could not extract item number or cycle time from PDF file: ${pdfFile.name}`);
                    }

                    processedCount++;
                    results.innerHTML = `<p>Processed ${processedCount} of ${pdfFiles.length} files...</p>`;
                    processNextPdf(index + 1); // Process the next PDF
                });
            };
            pdfReader.readAsArrayBuffer(pdfFile);
        };

        processNextPdf(0); // Start processing the first PDF
    };
    excelReader.readAsArrayBuffer(excelFile);
}

function extractItemNumber(text) {
    // Use regex to find the item number in the PDF text
    const regex = /Item Number\s*:\s*([\w\s-]+)/i;
    const match = text.match(regex);
    return match ? match[1].trim() : null; // Return the matched item number or null
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
