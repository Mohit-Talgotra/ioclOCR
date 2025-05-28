let currentJobId = null;
let statusInterval = null;

document.getElementById('fileInput').addEventListener('change', function(e) {
    handleFileSelection(e.target.files[0]);
});

const uploadArea = document.querySelector('.upload-area');

uploadArea.addEventListener('dragover', function(e) {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', function(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', function(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFileSelection(files[0]);
    }
});

function handleFileSelection(file) {
    if (file && file.type === 'application/pdf') {
        const fileSize = (file.size / (1024 * 1024)).toFixed(2);
        document.getElementById('fileName').textContent = `${file.name} (${fileSize} MB)`;
        document.getElementById('selectedFile').style.display = 'block';
        document.getElementById('uploadBtn').disabled = false;
    } else {
        showAlert('Please select a valid PDF file.', 'warning');
    }
}

function clearFile() {
    document.getElementById('fileInput').value = '';
    document.getElementById('selectedFile').style.display = 'none';
    document.getElementById('uploadBtn').disabled = true;
}

function resetForm() {
    clearFile();
    document.getElementById('progressContainer').style.display = 'none';
    document.getElementById('resultContainer').style.display = 'none';
    document.getElementById('errorContainer').style.display = 'none';
    if (statusInterval) {
        clearInterval(statusInterval);
    }
}

function showAlert(message, type) {
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type} alert-dismissible fade show position-fixed`;
    alertDiv.style.cssText = 'top: 20px; right: 20px; z-index: 9999; min-width: 300px;';
    alertDiv.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    document.body.appendChild(alertDiv);
    
    setTimeout(() => {
        if (alertDiv.parentNode) {
            alertDiv.parentNode.removeChild(alertDiv);
        }
    }, 5000);
}

document.getElementById('uploadForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    const formData = new FormData();
    const fileInput = document.getElementById('fileInput');
    
    if (!fileInput.files[0]) {
        showAlert('Please select a file first.', 'warning');
        return;
    }

    formData.append('file', fileInput.files[0]);

    try {
        document.getElementById('progressContainer').style.display = 'block';
        document.getElementById('uploadBtn').disabled = true;

        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (response.ok) {
            currentJobId = result.job_id;
            startStatusPolling();
        } else {
            showError(result.error || 'Upload failed');
        }
    } catch (error) {
        showError('Network error: ' + error.message);
    }
});

function startStatusPolling() {
    statusInterval = setInterval(async () => {
        try {
            const response = await fetch(`/status/${currentJobId}`);
            const status = await response.json();

            if (response.ok) {
                updateProgress(status);
                
                if (status.status === 'completed') {
                    clearInterval(statusInterval);
                    showSuccess(status.download_url);
                } else if (status.status === 'error') {
                    clearInterval(statusInterval);
                    showError(status.error || 'Processing failed');
                }
            } else {
                clearInterval(statusInterval);
                showError(status.error || 'Status check failed');
            }
        } catch (error) {
            clearInterval(statusInterval);
            showError('Network error: ' + error.message);
        }
    }, 2000);
}

function updateProgress(status) {
    const progressBar = document.getElementById('progressBar');
    const progressMessage = document.getElementById('progressMessage');
    
    progressBar.style.width = status.progress + '%';
    progressMessage.textContent = status.message;
}

function showSuccess(downloadUrl) {
    document.getElementById('progressContainer').style.display = 'none';
    document.getElementById('resultContainer').style.display = 'block';
    document.getElementById('downloadBtn').href = downloadUrl;
    document.getElementById('uploadBtn').disabled = false;
}

function showError(message) {
    document.getElementById('progressContainer').style.display = 'none';
    document.getElementById('errorContainer').style.display = 'block';
    document.getElementById('errorMessage').textContent = message;
    document.getElementById('uploadBtn').disabled = false;
}

document.addEventListener('DOMContentLoaded', function() {
    const elements = document.querySelectorAll('.animate-fade-in');
    elements.forEach((el, index) => {
        el.style.animationDelay = `${index * 0.2}s`;
    });
});