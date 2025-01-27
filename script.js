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

        // Column B is index 1 (zero-based)
        const itemNumberColumnIndex = 1;

        // Add "File Name" and "TOTAL CYCLE TIME" as new columns if they don't exist
        const fileNameColumnIndex = json[0].indexOf('File Name');
        if (fileNameColumnIndex === -1) {
            json[0].push('File Name'); // Add header
        }

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
                results.innerHTML = `<p>Excel file updated successfully! "File Name" and "TOTAL CYCLE TIME" added for ${processedCount} PDF(s).</p>`;
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
                    // Extract Item Number from the file name
                    const itemNumber = extractItemNumberFromFileName(pdfFile.name);

                    // Extract TOTAL CYCLE TIME from PDF text
                    const cycleTime = extractCycleTime(pdfText);

                    if (itemNumber) {
                        // Find the row in the Excel file that matches the item number in column B
                        let rowUpdated = false;

                        for (let i = 1; i < json.length; i++) {
                            if (json[i][itemNumberColumnIndex] == itemNumber) { // Use == for loose comparison
                                // Update the "File Name" and "TOTAL CYCLE TIME" columns for this row
                                if (fileNameColumnIndex === -1) {
                                    json[i].push(pdfFile.name); // Add to new column
                                } else {
                                    json[i][fileNameColumnIndex] = pdfFile.name; // Update existing column
                                }

                                if (cycleTimeColumnIndex === -1) {
                                    json[i].push(cycleTime || ''); // Add to new column
                                } else {
                                    json[i][cycleTimeColumnIndex] = cycleTime || ''; // Update existing column
                                }

                                rowUpdated = true;
                                console.log(`Updated row ${i} with File Name: ${pdfFile.name}, TOTAL CYCLE TIME: ${cycleTime}`);
                                break;
                            }
                        }

                        if (!rowUpdated) {
                            console.warn(`No matching item number found for PDF file: ${pdfFile.name}`);
                        }
                    } else {
                        console.warn(`Could not extract item number from PDF file name: ${pdfFile.name}`);
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

function extractItemNumberFromFileName(fileName) {
    // Use regex to extract the item number from the file name
    // Example: "Item12345.pdf" -> "12345"
    const regex = /Item\s*(\d+)/i;
    const match = fileName.match(regex);
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
