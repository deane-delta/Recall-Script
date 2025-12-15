let selectedFile = null;

// File input change handler
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        selectFile(file);
    }
});

// Drag and drop functionality
const uploadArea = document.getElementById('uploadArea');

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
        const file = files[0];
        if (isValidFileType(file)) {
            selectFile(file);
        } else {
            showError('Please select a valid file type (CSV, XLS, or XLSX)');
        }
    }
});

// Removed upload area click handler to prevent double file selection dialogs
// Users can now only select files via the "Select File" button or drag & drop

function isValidFileType(file) {
    const validTypes = ['.csv', '.xls', '.xlsx'];
    const fileName = file.name.toLowerCase();
    return validTypes.some(type => fileName.endsWith(type));
}

function selectFile(file) {
    if (!isValidFileType(file)) {
        showError('Please select a valid file type (CSV, XLS, or XLSX)');
        return;
    }

    if (file.size > 10 * 1024 * 1024) { // 10MB limit
        showError('File size must be less than 10MB');
        return;
    }

    selectedFile = file;
    displayFileInfo(file);
    loadColumnOptions(file);
    document.getElementById('processBtn').disabled = false;
    hideError();
}

function displayFileInfo(file) {
    const fileName = document.getElementById('fileName');
    const fileSize = document.getElementById('fileSize');
    const fileInfo = document.getElementById('fileInfo');

    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileInfo.style.display = 'block';
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function removeFile() {
    selectedFile = null;
    document.getElementById('fileInput').value = '';
    document.getElementById('fileInfo').style.display = 'none';
    hideColumnSelection();
    document.getElementById('processBtn').disabled = true;
    hideResults();
    hideError();
}

function showColumnSelection() {
    document.getElementById('columnSelection').style.display = 'block';
}

function hideColumnSelection() {
    document.getElementById('columnSelection').style.display = 'none';
}

function getSelectedColumn() {
    // Always use auto-detection now
    return 'auto';
}

function loadColumnOptions(file) {
    // Show the detection message immediately
    showColumnSelection();
    updateDetectionMessage('Looking for VIN numbers in "SERIAL NO" column...');
    
    const formData = new FormData();
    formData.append('excelFile', file);

    fetch('/get-columns', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            // Check if "SERIAL NO" column exists
            const hasSerialNo = data.columns.some(col => col.name === 'SERIAL NO');
            if (hasSerialNo) {
                updateDetectionMessage('Found "SERIAL NO" column. Ready for VIN detection.');
            } else {
                updateDetectionMessage('Warning: "SERIAL NO" column not found. Will search other columns.');
            }
        } else {
            console.error('Error loading columns:', data.error);
            updateDetectionMessage('Error reading file structure.', 'error');
        }
    })
    .catch(error => {
        console.error('Error loading columns:', error);
        updateDetectionMessage('Error reading file structure.', 'error');
    });
}

function updateDetectionMessage(message, type = '') {
    const messageElement = document.getElementById('detectionMessage');
    messageElement.textContent = message;
    messageElement.className = type ? `detection-message ${type}` : 'detection-message';
}

function processFile() {
    if (!selectedFile) {
        showError('Please select a file first');
        return;
    }

    const selectedColumn = getSelectedColumn();
    const sessionId = Date.now().toString();
    
    const formData = new FormData();
    formData.append('excelFile', selectedFile);
    formData.append('vinColumn', selectedColumn);
    formData.append('sessionId', sessionId);

    // Show progress and update detection message
    showProgress();
    updateDetectionMessage('Processing file and extracting VINs from "SERIAL NO" column...');
    hideError();
    hideResults();

    // Disable process button
    document.getElementById('processBtn').disabled = true;

    // Start listening for SSE updates
    const eventSource = new EventSource(`/progress/${sessionId}`);
    eventSource.onmessage = function(event) {
        const data = JSON.parse(event.data);
        
        if (data.type === 'progress') {
            updateProgressBar(data.progress, data.message);
        } else if (data.type === 'error') {
            // Handle error message
            showError(data.message);
            hideProgress();
            document.getElementById('processBtn').disabled = false;
        } else if (data.type === 'complete') {
            eventSource.close();
            hideProgress();
            
            showResults(data.data);
            if (data.data.detectedColumn) {
                // Calculate total recall/safety numbers found (excluding those with "NONE" EA numbers)
                let totalRecallSafetyCount = 0;
                if (data.data.scrapedData && data.data.scrapedData.length > 0) {
                    data.data.scrapedData.forEach(item => {
                        if (item.fordData && 
                            item.fordData.success && 
                            item.fordData.recallData && 
                            item.fordData.recallData.recalls) {
                            const validRecalls = item.fordData.recallData.recalls.filter(recall => {
                                // Filter out invalid recall numbers
                                if (!recall.recallNumber || 
                                    recall.recallNumber === 'No recall information' || 
                                    recall.recallNumber === 'No recall information available') {
                                    return false;
                                }
                                
                                // Check if EA number is "NONE" and exclude if so
                                let eaNumber = null;
                                if (item.docsearchDataByRecall && typeof item.docsearchDataByRecall === 'object') {
                                    // Handle as plain object (converted from Map for JSON serialization)
                                    const docsearchData = item.docsearchDataByRecall[recall.recallNumber];
                                    if (docsearchData) {
                                        eaNumber = docsearchData.eaNumber;
                                    }
                                } else if (item.docsearchData && item.docsearchData.eaNumber) {
                                    // Fallback to old format
                                    eaNumber = item.docsearchData.eaNumber;
                                }
                                
                                if (eaNumber === 'NONE') {
                                    return false;
                                }
                                
                                return true;
                            });
                            totalRecallSafetyCount += validRecalls.length;
                        }
                    });
                }
                
                const vinCount = data.data.vinCount || 0;
                updateDetectionMessage(`Successfully processed ${vinCount} VIN numbers and found ${totalRecallSafetyCount} Safety and Recall numbers.`);
            }
            document.getElementById('processBtn').disabled = false;
        }
    };
    
    eventSource.onerror = function(error) {
        console.error('SSE error:', error);
        eventSource.close();
    };

    fetch('/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (!data.success) {
            hideProgress();
            showError(data.error || 'An error occurred while processing the file');
            updateDetectionMessage('Failed to extract VINs from "SERIAL NO" column.', 'error');
            document.getElementById('processBtn').disabled = false;
        }
    })
    .catch(error => {
        hideProgress();
        showError('Network error: ' + error.message);
        console.error('Error:', error);
        eventSource.close();
        document.getElementById('processBtn').disabled = false;
    });
}

function showProgress() {
    const progressSection = document.getElementById('progressSection');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');

    progressSection.style.display = 'block';
    progressFill.style.width = '0%';
    progressText.textContent = 'Processing file...';
}

function updateProgressBar(progress, message) {
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    
    progressFill.style.width = progress + '%';
    progressText.textContent = message;
}

function hideProgress() {
    const progressSection = document.getElementById('progressSection');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');

    // Complete progress bar
    progressFill.style.width = '100%';
    progressText.textContent = 'Complete!';

    // Hide after a short delay
    setTimeout(() => {
        progressSection.style.display = 'none';
    }, 1000);
}

function showResults(data) {
    const resultsSection = document.getElementById('resultsSection');
    const resultsContent = document.getElementById('resultsContent');
    const downloadSection = document.getElementById('downloadSection');

    // Create results HTML
    let html = `
        <div class="result-item">
            <span class="result-label">File Name:</span>
            <span class="result-value">${data.downloadFile || data.fileName}</span>
        </div>
        <div class="result-item">
            <span class="result-label">Total Rows:</span>
            <span class="result-value">${data.totalRows}</span>
        </div>
        <div class="result-item">
            <span class="result-label">VIN Numbers Found:</span>
            <span class="result-value">${data.vinCount}</span>
        </div>
    `;


    if (data.availableColumns && data.availableColumns.length > 0) {
        html += `
            <div class="result-item">
                <span class="result-label">Available Columns:</span>
                <span class="result-value">${data.availableColumns.map(col => `${col.letter} (${col.name})`).join(', ')}</span>
            </div>
        `;
    }

    if (data.message) {
        html += `
            <div class="result-item">
                <span class="result-label">Status:</span>
                <span class="result-value">${data.message}</span>
            </div>
        `;
    }

    resultsContent.innerHTML = html;
    resultsSection.style.display = 'block';

    // Show download section if there's a downloadable file
    if (data.downloadFile) {
        downloadSection.style.display = 'block';
        document.getElementById('downloadBtn').dataset.filename = data.downloadFile;
        
        // Store scraped data for comparison
        window.currentScrapedData = data.scrapedData;
        window.currentOutputFile = data.downloadFile;
        
        // Show comparison section
        document.getElementById('comparisonSection').style.display = 'block';
    }
}

function hideResults() {
    document.getElementById('resultsSection').style.display = 'none';
    document.getElementById('downloadSection').style.display = 'none';
}

function showError(message) {
    const errorSection = document.getElementById('errorSection');
    const errorMessage = document.getElementById('errorMessage');
    
    errorMessage.textContent = message;
    errorSection.style.display = 'block';
}

function hideError() {
    document.getElementById('errorSection').style.display = 'none';
}

function downloadResults() {
    const downloadBtn = document.getElementById('downloadBtn');
    const filename = downloadBtn.dataset.filename;
    
    if (filename) {
        window.open(`/download/${filename}`, '_blank');
    } else {
        showError('No file available for download');
    }
}

// Comparison file handling
let comparisonFile = null;

function handleComparisonFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        comparisonFile = file;
        const fileInfo = document.getElementById('comparisonFileInfo');
        const fileName = document.getElementById('comparisonFileName');
        
        fileName.textContent = file.name;
        fileInfo.style.display = 'block';
        document.getElementById('compareBtn').disabled = false;
    }
}

function removeComparisonFile() {
    comparisonFile = null;
    document.getElementById('comparisonFileInput').value = '';
    document.getElementById('comparisonFileInfo').style.display = 'none';
    document.getElementById('compareBtn').disabled = true;
    document.getElementById('comparisonDownloadSection').style.display = 'none';
}

function compareFiles() {
    if (!comparisonFile) {
        showError('Please select a reference file first');
        return;
    }
    
    if (!window.currentScrapedData || !window.currentOutputFile) {
        showError('No processed data available for comparison');
        return;
    }
    
    const formData = new FormData();
    formData.append('comparisonFile', comparisonFile);
    formData.append('outputFile', window.currentOutputFile);
    
    // Disable compare button
    const compareBtn = document.getElementById('compareBtn');
    compareBtn.disabled = true;
    compareBtn.textContent = 'Comparing...';
    
    fetch('/compare', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            if (data.missingFile) {
                document.getElementById('comparisonDownloadSection').style.display = 'block';
                document.getElementById('comparisonDownloadBtn').dataset.filename = data.missingFile;
                hideError();
                alert(`Comparison complete! Found ${data.missingCount} missing recall/safety numbers.`);
            } else {
                hideError();
                alert('All recall/safety numbers were found in the reference file.');
            }
        } else {
            showError(data.error || 'Error comparing files');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        showError('Network error: ' + error.message);
    })
    .finally(() => {
        compareBtn.disabled = false;
        compareBtn.textContent = 'Compare Files';
    });
}

function downloadComparisonResults() {
    const downloadBtn = document.getElementById('comparisonDownloadBtn');
    const filename = downloadBtn.dataset.filename;
    
    if (filename) {
        window.open(`/download/${filename}`, '_blank');
    } else {
        showError('No comparison file available for download');
    }
}

// Drag and drop functionality for comparison file
document.addEventListener('DOMContentLoaded', function() {
    console.log('VIN Recall Scraper initialized');
    
    const comparisonUploadArea = document.getElementById('comparisonUploadArea');
    const comparisonFileInput = document.getElementById('comparisonFileInput');
    
    if (comparisonUploadArea && comparisonFileInput) {
        comparisonUploadArea.addEventListener('dragover', function(e) {
            e.preventDefault();
            comparisonUploadArea.classList.add('dragover');
        });
        
        comparisonUploadArea.addEventListener('dragleave', function(e) {
            e.preventDefault();
            comparisonUploadArea.classList.remove('dragover');
        });
        
        comparisonUploadArea.addEventListener('drop', function(e) {
            e.preventDefault();
            comparisonUploadArea.classList.remove('dragover');
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                const file = files[0];
                const allowedTypes = ['.csv', '.xls', '.xlsx'];
                const fileExt = '.' + file.name.split('.').pop().toLowerCase();
                
                if (allowedTypes.includes(fileExt)) {
                    comparisonFile = file;
                    const dataTransfer = new DataTransfer();
                    dataTransfer.items.add(file);
                    comparisonFileInput.files = dataTransfer.files;
                    handleComparisonFileSelect({ target: comparisonFileInput });
                } else {
                    showError('Please upload a valid Excel file (.xlsx, .xls, or .csv)');
                }
            }
        });
    }
});
