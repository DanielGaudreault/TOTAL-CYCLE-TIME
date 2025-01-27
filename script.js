document.addEventListener('DOMContentLoaded', (event) => {
    document.getElementById('fileForm').addEventListener('submit', function(e) {
        e.preventDefault(); // Prevent form from submitting normally
    });
});

function processFiles(event) {
    event.preventDefault();
    const fileInput = document.querySelector('input[name="files"]');
    const loading = document.getElementById('loading');
    const message = document.getElementById('message');
    const downloadButton = document.getElementById('downloadButton');

    if (fileInput.files.length === 0) {
        alert('Please select at least one file.');
        return;
    }

    loading.style.display = 'block';
    message.textContent = '';
    downloadButton.style.display = 'none';

    const formData = new FormData();
    for (let i = 0; i < fileInput.files.length; i++) {
        formData.append('files', fileInput.files[i]);
    }

    fetch('/process-files', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        loading.style.display = 'none';
        message.textContent = data.message;
        if (data.output_excel) {
            downloadButton.style.display = 'block';
            downloadButton.setAttribute('data-excel-path', data.output_excel);
        }
    })
    .catch(error => {
        loading.style.display = 'none';
        message.textContent = 'An error occurred while processing the files.';
        console.error('Error:', error);
    });
}

function downloadResults() {
    const excelPath = document.getElementById('downloadButton').getAttribute('data-excel-path');
    if (excelPath) {
        const link = document.createElement('a');
        link.href = excelPath;
        link.download = 'updated_excel.xlsx';
        link.click();
    }
}
