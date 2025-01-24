function uploadFiles() {
    const fileInput = document.getElementById('fileInput');
    const results = document.getElementById('results');
    const downloadLink = document.getElementById('downloadLink');
    results.textContent = ''; // Clear previous results

    if (fileInput.files.length === 0) {
        alert('Please select at least one file.');
        return;
    }

    const formData = new FormData();
    for (const file of fileInput.files) {
        formData.append('files', file);
    }

    fetch('/process', {
        method: 'POST',
        body: formData,
    })
    .then(response => response.json())
    .then(data => {
        if (data.error) {
            results.textContent = `Error: ${data.error}`;
        } else {
            results.textContent = JSON.stringify(data.results, null, 2);
            if (data.outputFile) {
                downloadLink.href = `/download/${data.outputFile}`;
                downloadLink.style.display = 'block';
                downloadLink.download = data.outputFile;
            }
        }
    })
    .catch(error => {
        results.textContent = `Error: ${error.message}`;
    });
}
