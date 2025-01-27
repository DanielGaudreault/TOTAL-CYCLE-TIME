document.getElementById('upload-form').addEventListener('submit', function(event) {
    event.preventDefault();
    
    const excelFiles = document.getElementById('excel-files').files;
    const pdfFiles = document.getElementById('pdf-files').files;
    const formData = new FormData();

    for (let i = 0; i < excelFiles.length; i++) {
        formData.append('excel-files', excelFiles[i]);
    }

    for (let i = 0; i < pdfFiles.length; i++) {
        formData.append('pdf-files', pdfFiles[i]);
    }

    fetch('/process-files', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        document.getElementById('output').textContent = 'Files processed successfully!';
        console.log(data);
    })
    .catch(error => {
        document.getElementById('output').textContent = 'Error processing files.';
        console.error('Error:', error);
    });
});
