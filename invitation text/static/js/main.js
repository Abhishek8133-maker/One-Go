function sendEmail() {
    const invitation = document.querySelector('p').textContent;
    const filename = document.querySelector('input[type="file"]').files[0]?.name || '';  // Get the file name

    const formData = new FormData();
    formData.append('invitation', invitation);
    formData.append('filename', filename);

    document.getElementById('loading-spinner').style.display = 'block';

    fetch('/send_email', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        const notification = document.getElementById('notification');
        const spinner = document.getElementById('loading-spinner');

        spinner.style.display = 'none';

        if (data.status === 'success') {
            notification.textContent = data.message;
            notification.style.backgroundColor = '#d4edda';
            notification.style.color = '#155724';
        } else {
            notification.textContent = `Error: ${data.message}`;
            notification.style.backgroundColor = '#f8d7da';
            notification.style.color = '#721c24';
        }
        notification.style.display = 'block';
    })
    .catch(error => {
        document.getElementById('loading-spinner').style.display = 'none';
        console.error('Error:', error);
    });
}
 