function processFiles() {
    const excelFile = document.getElementById('excel').files[0];
    const pdfFile = document.getElementById('pdf').files[0];
    const results = document.getElementById('results');

    if (!excelFile || !pdfFile) {
        results.innerHTML = '<p>Please upload both Excel and PDF files.</p>';
        return;
    }

    // Read PDF file
    const pdfReader = new FileReader();
    pdfReader.onload = function (event) {
        const pdfData = new Uint8Array(event.target.result);
        parsePDF(pdfData).then(pdfText => {
            // Read Excel file
            const excelReader = new FileReader();
            excelReader.onload = function (event) {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                // Convert Excel sheet to JSON
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                // Add PDF text as a new column
                json[0].push('PDF_Text'); // Add header
                for (let i = 1; i < json.length; i++) {
                    json[i].push(pdfText); // Add PDF text to each row
                }

                // Convert JSON back to Excel
                const updatedWorksheet = XLSX.utils.aoa_to_sheet(json);
                const updatedWorkbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(updatedWorkbook, updatedWorksheet, 'Sheet1');

                // Download the updated Excel file
                XLSX.writeFile(updatedWorkbook, 'updated_excel.xlsx');
                results.innerHTML = '<p>Excel file updated successfully! Download started.</p>';
            };
            excelReader.readAsArrayBuffer(excelFile);
        });
    };
    pdfReader.readAsArrayBuffer(pdfFile);
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
