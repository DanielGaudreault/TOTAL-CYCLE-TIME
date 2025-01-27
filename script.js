function searchFile() {
    const fileInput = document.getElementById('fileInput');
    const excelInput = document.getElementById('excelInput');
    const results = document.getElementById('results');
    results.textContent = ''; // Clear previous results

    if (fileInput.files.length === 0) {
        alert('Please select at least one file.');
        return;
    }

    if (!excelInput.files.length) {
        alert('Please select an Excel file.');
        return;
    }

    const files = fileInput.files;
    const resultsArray = [];
    let excelData = [];

    // Parse the Excel file
    const parseExcel = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function (event) {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0]; // Assuming the first sheet
                    const sheet = workbook.Sheets[sheetName];
                    excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                    resolve(workbook);
                } catch (error) {
                    console.error('Error parsing Excel:', error);
                    reject(error);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    };

    const processFile = (file, index) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = function (event) {
                try {
                    const content = event.target.result;
                    if (file.type === 'application/pdf') {
                        parsePDF(content).then(text => {
                            console.log(`Text extracted from ${file.name}:`, text);
                            const cycleTime = extractCycleTime(text);
                            resultsArray[index] = { fileName: file.name, cycleTime };
                            resolve();
                        }).catch(reject);
                    } else if (file.type === 'text/plain') {
                        const cycleTime = extractCycleTime(content);
                        resultsArray[index] = { fileName: file.name, cycleTime };
                        resolve();
                    } else {
                        reject(new Error('Unsupported file type'));
                    }
                } catch (error) {
                    console.error(`Error processing file ${file.name}:`, error);
                    reject(error);
                }
            };

            if (file.type === 'application/pdf') {
                reader.readAsArrayBuffer(file);
            } else if (file.type === 'text/plain') {
                reader.readAsText(file);
            } else {
                reject(new Error('Unsupported file type'));
            }
        });
    };

    const processAllFiles = async () => {
        results.textContent = 'Processing files...';
        let workbook;
        try {
            workbook = await parseExcel(excelInput.files[0]);
        } catch (error) {
            results.textContent = `Error parsing Excel file: ${error.message}`;
            return;
        }

        console.log('Excel Data Before Update:', excelData);

        // Process each file
        for (let i = 0; i < files.length; i++) {
            try {
                await processFile(files[i], i);
            } catch (error) {
                console.error('Error:', error);
                results.textContent += `\nError processing ${files[i].name}: ${error.message}`;
            }
        }

        // Update or add new rows to Excel based on filenames
        let newRowNumber = excelData.length + 1; // Start numbering from where existing data ends
        resultsArray.forEach(result => {
            const rowIndex = excelData.findIndex(row => row[0] && row[0].toString() === result.fileName);
            if (rowIndex === -1) {
                // If no match, add a new row
                const newRow = [result.fileName, result.cycleTime || 'No instances of "TOTAL CYCLE TIME" found.'];
                excelData.push(newRow);
                console.log(`Added new row for ${result.fileName} with number ${newRowNumber}`);
                newRowNumber++;
            } else {
                // Update existing row
                excelData[rowIndex][1] = result.cycleTime || 'No instances of "TOTAL CYCLE TIME" found.';
            }
        });

        console.log('Excel Data After Update:', excelData);

        // Update the sheet with modified data
        const updatedSheet = XLSX.utils.aoa_to_sheet(excelData);
        workbook.Sheets[workbook.SheetNames[0]] = updatedSheet;

        // Export the updated workbook
        const updatedExcelFile = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

        // Trigger the download of the updated Excel file
        const blob = new Blob([updatedExcelFile], { type: 'application/octet-stream' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'updated_file.xlsx';
        link.click();

        results.textContent = 'Processing complete. Download started.';
    };

    processAllFiles().catch(error => {
        console.error('Error processing files:', error);
        results.textContent = 'An error occurred while processing the files: ' + error.message;
    });
}

function extractCycleTime(text) {
    const lines = text ? text.split('\n') : []; // Split text into lines
    for (const line of lines) {
        if (line.includes("TOTAL CYCLE TIME")) {
            console.log('Line with cycle time:', line);
            const regex = /(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)/i;
            const match = line.match(regex);
            console.log('Match:', match);
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
            let pagesRead = 0;

            const fetchPage = (pageNum) => {
                return pdf.getPage(pageNum).then(page => {
                    return page.getTextContent().then(textContent => {
                        let pageText = '';
                        textContent.items.forEach(item => {
                            pageText += item.str + ' ';
                        });
                        text += pageText + '\n'; // Add newline after each page
                        pagesRead++;
                        if (pagesRead === numPages) {
                            resolve(text);
                        }
                    });
                });
            };

            const fetchAllPages = async () => {
                for (let i = 1; i <= numPages; i++) {
                    try {
                        await fetchPage(i);
                    } catch (error) {
                        console.error(`Error fetching page ${i}:`, error);
                    }
                }
            };

            fetchAllPages();
        }).catch(reject);
    });
}
