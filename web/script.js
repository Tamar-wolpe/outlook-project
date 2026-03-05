// Update file name display when a file is selected
document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('myForm');
    const fileInput = document.getElementById('file');
    const fileNameDisplay = document.getElementById('fileName');
    const fileLabel = document.getElementById('fileLabel');

    // Show selected file name
    fileInput.addEventListener('change', function(e) {
        if (this.files && this.files.length > 0) {
            const file = this.files[0];
            fileNameDisplay.textContent = file.name;
            
            // Add a small preview for images
            if (file.type.startsWith('image/')) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    let imgPreview = fileLabel.querySelector('img');
                    if (!imgPreview) {
                        imgPreview = document.createElement('img');
                        imgPreview.style.maxWidth = '100%';
                        imgPreview.style.maxHeight = '150px';
                        imgPreview.style.marginTop = '10px';
                        imgPreview.style.borderRadius = '4px';
                        fileLabel.appendChild(imgPreview);
                    }
                    imgPreview.src = e.target.result;
                };
                reader.readAsDataURL(file);
            }
        } else {
            fileNameDisplay.textContent = 'לא נבחר קובץ';
            const imgPreview = fileLabel.querySelector('img');
            if (imgPreview) {
                imgPreview.remove();
            }
        }
    });

    // Handle form submission
    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        const submitBtn = form.querySelector('button[type="submit"]');
        submitBtn.disabled = true;
        submitBtn.textContent = 'שולח...';
        
        // Create FormData from the form
        const formData = new FormData(form);
        
        try {
            const response = await fetch('http://localhost:3000/create-draft', {
                method: 'POST',
                body: formData
            });
            
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new Error(errorData.message || 'שגיאה בשרת');
            }
            
            const result = await response.json();
            if (result.success) {
                alert('הטיוטה נוצרה בהצלחה! נפתח חלון Outlook עם הטיוטה.');
                form.reset();
                fileNameDisplay.textContent = 'לא נבחר קובץ';
                const imgPreview = fileLabel.querySelector('img');
                if (imgPreview) imgPreview.remove();
            } else {
                throw new Error(result.message || 'אירעה שגיאה ביצירת הטיוטה');
            }
        } catch (error) {
            console.error('Error:', error);
            alert('שגיאה: ' + (error.message || 'לא ניתן לשלוח את הטופס כרגע'));
        } finally {
            submitBtn.disabled = false;
            submitBtn.textContent = 'צור טיוטה ב-Outlook';
        }
    });

    // Add drag and drop support
    const dropArea = fileLabel;
    
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });

    function highlight() {
        dropArea.classList.add('highlight');
    }

    function unhighlight() {
        dropArea.classList.remove('highlight');
    }

    // Handle dropped files
    dropArea.addEventListener('drop', handleDrop, false);

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length) {
            fileInput.files = files;
            const event = new Event('change');
            fileInput.dispatchEvent(event);
        }
    }
});
