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
        let excelRows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

        // Create an object to store the total cycle time for each project
        let cycleTimeSums = {};

        // Process each result from the PDFs
        results.forEach(result => {
            // Extract the project name from the PDF filename, assuming it's something like "Project Name - Identifier"
            const matchIdentifier = result.fileName.match(/Project Name - (.+)/)?.[1].trim();

            if (!matchIdentifier) {
                console.error('Could not extract match identifier from PDF file name:', result.fileName);
                return; // Skip this result if we can't find a match identifier
            }

            // If the project already exists in the cycleTimeSums map, add the cycle time, otherwise set it
            if (cycleTimeSums[matchIdentifier]) {
                cycleTimeSums[matchIdentifier] += parseCycleTime(result.cycleTime);
            } else {
                cycleTimeSums[matchIdentifier] = parseCycleTime(result.cycleTime);
            }
        });

        // Find matching rows in Excel and update them
        excelRows.forEach((row, rowIndex) => {
            const matchIdentifier = row[0]?.toString().trim(); // Assuming 'Item No.' is in the first column
            if (cycleTimeSums[matchIdentifier]) {
                // Match found, update the cycle time sum in the 'Total Cycle Time' column (assuming column D is where it goes)
                row[3] = formatCycleTime(cycleTimeSums[matchIdentifier]);
            }
        });

        // Convert the updated data back to worksheet format
        const newWS = XLSX.utils.aoa_to_sheet(excelRows);
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, newWS, sheetName);

        // Save the new workbook
        XLSX.writeFile(newWB, 'updated_cycle_times.xlsx');
    };
    reader.readAsArrayBuffer(file);
}

// Helper function to parse cycle time (e.g., "0 HOURS, 4 MINUTES, 16 SECONDS" -> total seconds)
function parseCycleTime(cycleTimeString) {
    if (!cycleTimeString) return 0;

    const regex = /(\d+)\s*HOURS?/, hoursMatch = cycleTimeString.match(regex);
    const regexMinutes = /(\d+)\s*MINUTES?/, minutesMatch = cycleTimeString.match(regexMinutes);
    const regexSeconds = /(\d+)\s*SECONDS?/, secondsMatch = cycleTimeString.match(regexSeconds);

    const hours = hoursMatch ? parseInt(hoursMatch[1]) : 0;
    const minutes = minutesMatch ? parseInt(minutesMatch[1]) : 0;
    const seconds = secondsMatch ? parseInt(secondsMatch[1]) : 0;

    return (hours * 3600) + (minutes * 60) + seconds;
}

// Helper function to format cycle time back to a readable string (seconds -> "X HOURS, Y MINUTES, Z SECONDS")
function formatCycleTime(totalSeconds) {
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;

    return `${hours} HOURS, ${minutes} MINUTES, ${seconds} SECONDS`;
}
