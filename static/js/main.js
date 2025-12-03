// ==================== DOM Elements ====================
const uploadForm = document.getElementById('uploadForm');
const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const downloadSection = document.getElementById('downloadSection');

const zipFileInput = document.getElementById('zipFile');
const csvFileInput = document.getElementById('csvFile');
const zipFileName = document.getElementById('zipFileName');
const csvFileName = document.getElementById('csvFileName');

const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const progressPercentage = document.getElementById('progressPercentage');
const resultsBody = document.getElementById('resultsBody');

const downloadCsvBtn = document.getElementById('downloadCsvBtn');
const downloadExcelBtn = document.getElementById('downloadExcelBtn');
const resetBtn = document.getElementById('resetBtn');

const feedbackModal = document.getElementById('feedbackModal');
const modalClose = document.getElementById('modalClose');
const modalFeedbackText = document.getElementById('modalFeedbackText');

let currentSessionId = null;

// ==================== File Upload UI ====================
zipFileInput.addEventListener('change', function() {
    const fileUpload = this.parentElement;
    if (this.files.length > 0) {
        zipFileName.textContent = this.files[0].name;
        fileUpload.classList.add('has-file');
    } else {
        zipFileName.textContent = 'Drop ZIP file here or click to browse';
        fileUpload.classList.remove('has-file');
    }
});

csvFileInput.addEventListener('change', function() {
    const fileUpload = this.parentElement;
    if (this.files.length > 0) {
        csvFileName.textContent = this.files[0].name;
        fileUpload.classList.add('has-file');
    } else {
        csvFileName.textContent = 'Drop CSV file here or click to browse';
        fileUpload.classList.remove('has-file');
    }
});

// ==================== Form Submission ====================
uploadForm.addEventListener('submit', async function(e) {
    e.preventDefault();

    const formData = new FormData(this);

    try {
        const response = await fetch('/process', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const error = await response.json();
            alert('Error: ' + (error.error || 'Upload failed'));
            return;
        }

        const data = await response.json();
        currentSessionId = data.session_id;

        uploadSection.classList.add('hidden');
        processingSection.classList.remove('hidden');
        processingSection.classList.add('fade-in');

        startEventStream(currentSessionId, data.task_type);

    } catch (error) {
        console.error('Error:', error);
        alert('Error uploading files: ' + error.message);
    }
});

// ==================== Server-Sent Events ====================
function startEventStream(sessionId, taskType) {
    const eventSource = new EventSource('/stream/' + sessionId + '?task_type=' + taskType);

    eventSource.onmessage = function(event) {
        const data = JSON.parse(event.data);

        switch (data.type) {
            case 'progress':
                handleProgress(data);
                break;
            case 'complete':
                handleComplete(data);
                eventSource.close();
                break;
            case 'error':
                handleError(data);
                break;
            case 'fatal_error':
                handleFatalError(data);
                eventSource.close();
                break;
        }
    };

    eventSource.onerror = function(error) {
        console.error('EventSource error:', error);
        eventSource.close();
        alert('Connection error. Please refresh and try again.');
    };
}

// ==================== Progress Handling ====================
function handleProgress(data) {
    progressFill.style.width = data.percentage + '%';
    progressText.textContent = 'Processing ' + data.current + ' of ' + data.total + ' files...';
    progressPercentage.textContent = data.percentage + '%';

    addResultRow(data.result);
}

function addResultRow(result) {
    const row = document.createElement('tr');

    let matchBadge = '';
    if (result.match_status === 'success') {
        matchBadge = '<span class="status-badge status-success">✓ Matched</span>';
    } else {
        matchBadge = '<span class="status-badge status-danger">✗ No match</span>';
    }

    let matchPercentage = result.match_percentage > 0
        ? result.match_percentage + '%'
        : '—';

    const escapedComment = escapeHtml(result.comment || '');
    
    row.innerHTML = 
        '<td><small>' + escapeHtml(result.file_name) + '</small></td>' +
        '<td>' + escapeHtml(result.student_name) + '</td>' +
        '<td>' + (result.matched_name ? escapeHtml(result.matched_name) : '—') + '</td>' +
        '<td>' + matchPercentage + '</td>' +
        '<td><strong>' + (result.score || '—') + '</strong></td>' +
        '<td><span class="feedback-preview" data-feedback="' + escapedComment.replace(/"/g, '&quot;') + '">' +
        escapeHtml(result.comment_preview || 'No feedback') + '</span></td>';

    resultsBody.appendChild(row);

    row.scrollIntoView({ behavior: 'smooth', block: 'end' });
}

// ==================== Completion Handling ====================
function handleComplete(data) {
    progressFill.style.width = '100%';
    progressText.textContent = 'Completed! Processed ' + data.total + ' files.';
    progressPercentage.textContent = '100%';

    setTimeout(function() {
        processingSection.classList.add('hidden');
        downloadSection.classList.remove('hidden');
        downloadSection.classList.add('fade-in');

        downloadCsvBtn.href = '/download/' + currentSessionId + '/csv';
        downloadExcelBtn.href = '/download/' + currentSessionId + '/excel';
    }, 1000);
}

// ==================== Error Handling ====================
function handleError(data) {
    console.error('Processing error:', data);
    const row = document.createElement('tr');
    row.innerHTML =
        '<td colspan="6" style="color: var(--danger-color);">Error processing ' +
        escapeHtml(data.file) + ': ' + escapeHtml(data.message) + '</td>';
    resultsBody.appendChild(row);
}

function handleFatalError(data) {
    alert('Fatal error: ' + data.message);
    resetApp();
}

// ==================== Modal for Full Feedback ====================
resultsBody.addEventListener('click', function(e) {
    if (e.target.classList.contains('feedback-preview')) {
        const fullText = e.target.getAttribute('data-feedback');
        showFeedback(fullText);
    }
});

function showFeedback(fullText) {
    modalFeedbackText.textContent = fullText;
    feedbackModal.classList.remove('hidden');
}

modalClose.addEventListener('click', function() {
    feedbackModal.classList.add('hidden');
});

feedbackModal.addEventListener('click', function(e) {
    if (e.target === feedbackModal) {
        feedbackModal.classList.add('hidden');
    }
});

// ==================== Reset ====================
resetBtn.addEventListener('click', function() {
    resetApp();
});

function resetApp() {
    uploadForm.reset();
    zipFileName.textContent = 'Drop ZIP file here or click to browse';
    csvFileName.textContent = 'Drop CSV file here or click to browse';
    zipFileInput.parentElement.classList.remove('has-file');
    csvFileInput.parentElement.classList.remove('has-file');

    resultsBody.innerHTML = '';

    progressFill.style.width = '0%';
    progressText.textContent = 'Starting...';
    progressPercentage.textContent = '0%';

    uploadSection.classList.remove('hidden');
    processingSection.classList.add('hidden');
    downloadSection.classList.add('hidden');

    currentSessionId = null;
}

// ==================== Utility Functions ====================
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ==================== Drag and Drop ====================
document.querySelectorAll('.file-upload').forEach(function(fileUpload) {
    const input = fileUpload.querySelector('input[type="file"]');

    fileUpload.addEventListener('dragover', function(e) {
        e.preventDefault();
        this.querySelector('.file-upload-label').style.borderColor = 'var(--primary-color)';
    });

    fileUpload.addEventListener('dragleave', function(e) {
        e.preventDefault();
        this.querySelector('.file-upload-label').style.borderColor = 'var(--border-color)';
    });

    fileUpload.addEventListener('drop', function(e) {
        e.preventDefault();
        this.querySelector('.file-upload-label').style.borderColor = 'var(--border-color)';

        if (e.dataTransfer.files.length > 0) {
            input.files = e.dataTransfer.files;
            input.dispatchEvent(new Event('change'));
        }
    });
});
