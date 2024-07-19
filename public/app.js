document.getElementById('uploadForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent default form submission

    const formData = new FormData();
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    formData.append('file', file);

    fetch('/convert', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Failed to convert file');
        }
        return response.blob();
    })
    .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'output.csv';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.getElementById('status').innerText = 'File converted successfully!';
    })
    .catch(error => {
        console.error('Error:', error);
        document.getElementById('status').innerText = 'Failed to convert file';
    });
});
