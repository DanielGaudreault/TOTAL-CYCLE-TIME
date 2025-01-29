function updateToExcel() {
    console.log('updateToExcel function called');
    const fileInput = document.getElementById('uploadExcelInput');
    const file = fileInput ? fileInput.files[0] : null;

    if (!file) {
        alert('Please select an Excel file to update.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        console.log('File read successfully');
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        console.log('Workbook parsed:', workbook);

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        let excelRows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

        console.log("Original Excel Rows:", excelRows);
        console.log('Results for updating:', results);

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
                console.log(`Added ${cycleTimeInSeconds} seconds to ${matchIdentifier}.`);
            } else {
                console.error('Could not parse cycle time for:', result.fileName);
            }
        });

        console.log("Cycle Time Sums:", cycleTimeSums);
        console.log("Total Cycle Time:", totalCycleTime);

        // Adding subtotals for each project
        Object.keys(cycleTimeSums).forEach(project => {
            excelRows.push([
                `Subtotal for ${project}`,
                '', 
                '', 
                formatCycleTime(cycleTimeSums[project])
            ]);
            console.log(`Added subtotal for ${project}: ${formatCycleTime(cycleTimeSums[project])}`);
        });

        // Adding the total cycle time at the very end
        excelRows.push([
            'Net Total Cycle Time', 
            '', 
            '', 
            formatCycleTime(totalCycleTime)
        ]);
        console.log(`Added net total cycle time: ${formatCycleTime(totalCycleTime)}`);

        console.log('Updated Excel Rows:', excelRows);

        const newWS = XLSX.utils.aoa_to_sheet(excelRows);
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, newWS, sheetName);

        XLSX.writeFile(newWB, 'updated_cycle_times.xlsx');
        console.log('New Excel file created and saved');
    };
    reader.readAsArrayBuffer(file);
}
