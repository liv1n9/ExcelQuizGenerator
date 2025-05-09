document.addEventListener('DOMContentLoaded', function() {
    // Get DOM elements
    const uploadForm = document.getElementById('uploadForm');
    const excelFileInput = document.getElementById('excelFile');
    const numQuestionsInput = document.getElementById('numQuestions');
    const numVersionsInput = document.getElementById('numVersions');
    const generateBtn = document.getElementById('generateBtn');
    const alertArea = document.getElementById('alertArea');
    const downloadArea = document.getElementById('downloadArea');
    const regularDownloadBtn = document.getElementById('regularDownloadBtn');
    const highlightedDownloadBtn = document.getElementById('highlightedDownloadBtn');
    const loadingOverlay = document.getElementById('loadingOverlay');
    
    // Form validation
    uploadForm.addEventListener('submit', function(event) {
        event.preventDefault();
        
        // Clear previous alerts
        alertArea.innerHTML = '';
        downloadArea.classList.add('d-none');
        
        // Validate form
        if (!uploadForm.checkValidity()) {
            event.stopPropagation();
            uploadForm.classList.add('was-validated');
            return;
        }
        
        // Validate file type
        const file = excelFileInput.files[0];
        if (!file) {
            showAlert('Please select an Excel file.', 'danger');
            return;
        }
        
        const fileExt = file.name.split('.').pop().toLowerCase();
        if (fileExt !== 'xlsx' && fileExt !== 'xls') {
            showAlert('Invalid file type. Please upload an Excel file (.xlsx or .xls).', 'danger');
            return;
        }
        
        // Validate number inputs
        const numQuestions = parseInt(numQuestionsInput.value);
        const numVersions = parseInt(numVersionsInput.value);
        
        if (isNaN(numQuestions) || numQuestions <= 0) {
            showAlert('Number of questions must be a positive number.', 'danger');
            return;
        }
        
        if (isNaN(numVersions) || numVersions <= 0) {
            showAlert('Number of versions must be a positive number.', 'danger');
            return;
        }
        
        // Show loading overlay
        loadingOverlay.classList.remove('d-none');
        
        // Submit form data
        const formData = new FormData(uploadForm);
        
        fetch('/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            // Hide loading overlay
            loadingOverlay.classList.add('d-none');
            
            if (data.error) {
                // Show error message
                showAlert(data.error, 'danger');
            } else {
                // Show success message
                showAlert('Files generated successfully! Click the download buttons below.', 'success');
                
                // Set download links
                regularDownloadBtn.href = `/download/${data.regular_zip}`;
                highlightedDownloadBtn.href = `/download/${data.highlighted_zip}`;
                
                // Show download area
                downloadArea.classList.remove('d-none');
            }
        })
        .catch(error => {
            // Hide loading overlay
            loadingOverlay.classList.add('d-none');
            
            // Show error message
            showAlert('An error occurred: ' + error.message, 'danger');
            console.error('Error:', error);
        });
    });
    
    // Function to show alerts
    function showAlert(message, type) {
        alertArea.innerHTML = `
            <div class="alert alert-${type} alert-dismissible fade show" role="alert">
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
            </div>
        `;
    }
    
    // File input change event for better UX
    excelFileInput.addEventListener('change', function() {
        if (this.files.length > 0) {
            const fileName = this.files[0].name;
            const fileSize = (this.files[0].size / 1024).toFixed(2);
            
            const fileLabel = this.nextElementSibling;
            if (fileLabel) {
                fileLabel.textContent = `${fileName} (${fileSize} KB)`;
            }
        }
    });
    
    // Add input validation for number fields
    numQuestionsInput.addEventListener('input', function() {
        if (this.value < 1) {
            this.value = 1;
        }
    });
    
    numVersionsInput.addEventListener('input', function() {
        if (this.value < 1) {
            this.value = 1;
        }
    });
    
    // Add example values for better UX
    numQuestionsInput.placeholder = "e.g., 10";
    numVersionsInput.placeholder = "e.g., 2";
});
