// ... (previous code remains the same)

// New function to handle Excel upload
function updateFromExcel() {
    const fileInput = document.getElementById('uploadExcelInput');
    const file = fileInput.files[0];

    if (!file) {
        alert('Please select an Excel file.');
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

        // Clear previous results
        results = [];
        const resultsTable = document.getElementById('resultsTable').getElementsByTagName('tbody')[0];
        resultsTable.innerHTML = '';

        // Update results and table from Excel data
        jsonData.slice(1).forEach(row => {
            if (row.length >= 2) { // Check if we have at least file name and cycle time
                results.push({ fileName: row[0], cycleTime: row[1] });
                const newRow = resultsTable.insertRow();
                newRow.insertCell().textContent = row[0];
                newRow.insertCell().textContent = row[1] || 'Not Found';
            }
        });

        document.getElementById('downloadButton').style.display = 'block';
    };
    reader.readAsArrayBuffer(file);
}

// Update the event listeners section
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processButton').addEventListener('click', processFiles);
    document.getElementById('resetButton').addEventListener('click', resetResults);
    document.getElementById('downloadButton').addEventListener('click', downloadResults);
    document.getElementById('uploadExcelButton').addEventListener('click', updateFromExcel);
});
