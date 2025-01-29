document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('processButton').addEventListener('click', processFiles);
    document.getElementById('resetButton').addEventListener('click', resetResults);
    document.getElementById('uploadExcelButton').addEventListener('click', updateToExcel);
});

let results = [];

async function processFiles() {
    // ... (keep the existing code for processFiles)
}

function readFile(file) {
    // ... (keep the existing code for readFile)
}

function extractCycleTime(text) {
    // ... (keep the existing code for extractCycleTime)
}

function parsePDF(data) {
    // ... (keep the existing code for parsePDF)
}

function resetResults() {
    // ... (keep the existing code for resetResults)
}

function parseCycleTime(cycleTimeString) {
    // ... (keep the existing code for parseCycleTime)
}

function formatCycleTime(totalSeconds) {
    // ... (keep the existing code for formatCycleTime)
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
        let excelRows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

        let cycleTimeSums = {};
        let totalCycleTime = 0;

        results.forEach(result => {
            const matchIdentifier = result.fileName.match(/^(.+?) - \d+/)?.[1].trim();
            if (!matchIdentifier) {
                console.error('Could not extract project name from PDF file name:', result.fileName);
                return;
            }

            const cycleTimeInSeconds = parseCycleTime(result.cycleTime);
            if (!isNaN(cycleTimeInSeconds)) {
                totalCycleTime += cycleTimeInSeconds;
                cycleTimeSums[matchIdentifier] = (cycleTimeSums[matchIdentifier] || 0) + cycleTimeInSeconds;
            } else {
                console.error('Could not parse cycle time for:', result.fileName);
            }
        });

        console.log("Cycle Time Sums:", cycleTimeSums);
        console.log("Total Cycle Time:", totalCycleTime);

        // Add subtotals for each project
        Object.keys(cycleTimeSums).forEach(project => {
            excelRows.push([
                `Subtotal for ${project}`,
                '', 
                '', 
                formatCycleTime(cycleTimeSums[project])
            ]);
        });

        // Add the total cycle time at the very end
        excelRows.push([
            'Net Total Cycle Time', 
            '', 
            '', 
            formatCycleTime(totalCycleTime)
        ]);

        const newWS = XLSX.utils.aoa_to_sheet(excelRows);
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, newWS, sheetName);

        XLSX.writeFile(newWB, 'updated_cycle_times.xlsx');
        console.log("Updated Excel sheet has been saved.");
    };
    reader.readAsArrayBuffer(file);
}
