document.getElementById('upload-form').addEventListener('submit', function(event) {
    event.preventDefault();

    const files = document.getElementById('files').files;
    const formData = new FormData();

    for (let i = 0; i < files.length; i++) {
        formData.append('files', files[i]);
        console.log(`Appending file: ${files[i].name}`);
    }

    console.log('Sending form data to server...');
    
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
