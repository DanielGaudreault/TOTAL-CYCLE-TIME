document.getElementById('upload-form').addEventListener('submit', function(event) {
    event.preventDefault();

    const excelFiles = document.getElementById('excel-files').files;
    const pdfFiles = document.getElementById('pdf-files').files;
    const formData = new FormData();

    for (let i = 0; i < excelFiles.length; i++) {
        formData.append('excel-files', excelFiles[i]);
        console.log(`Appending Excel file: ${excelFiles[i].name}`);
    }

    for (let i = 0; i < pdfFiles.length; i++) {
        formData.append('pdf-files', pdfFiles[i]);
        console.log(`Appending PDF file: ${pdfFiles[i].name}`);
    }

    fetch('/process-files', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        console.log('Response status:', response.status);
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        return response.json();
    })
    .then(data => {
        document.getElementById('results').textContent = 'Files processed successfully!';
        console.log('Response data:', data);
    })
    .catch(error => {
        document.getElementById('results').textContent = 'Error processing files.';
        console.error('Error:', error);
    });
});
