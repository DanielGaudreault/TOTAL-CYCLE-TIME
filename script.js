document.getElementById('processBtn').addEventListener('click', processPDFs);

async function processPDFs() {
    const files = document.getElementById('pdfFiles').files;
    if (files.length === 0) {
        alert('Please select at least one PDF file.');
        return;
    }

    let results = [];

    for (let file of files) {
        const reader = new FileReader();
        reader.onload = async function(e) {
            const pdfData = new Uint8Array(e.target.result);
            const pdf = await pdfjsLib.getDocument({data: pdfData}).promise;
            let text = '';

            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const content = await page.getTextContent();
                text += content.items.map(item => item.str).join(' ');
            }

            // Extract project name and cycle time (this is a simple example)
            const projectNameMatch = text.match(/Project Name: (\w+)/);
            const cycleTimeMatch = text.match(/Cycle Time: (\d+)/);

            if (projectNameMatch && cycleTimeMatch) {
                results.push({
                    projectName: projectNameMatch[1],
                    cycleTime: parseInt(cycleTimeMatch[1], 10)
                });
            }

            if (results.length === files.length) {
                exportToExcel(results);
            }
        };
        reader.readAsArrayBuffer(file);
    }
}

function exportToExcel(data) {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Cycle Times");
    XLSX.writeFile(wb, "Cycle_Times.xlsx");
    document.getElementById('output').innerText = 'Data exported to Cycle_Times.xlsx';
}
