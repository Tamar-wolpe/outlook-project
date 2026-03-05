const express = require('express');
const multer = require('multer');
const path = require('path');
const { exec } = require('child_process');
const cors = require('cors');
const fs = require('fs');

const app = express();

// Enable CORS
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Create temp directory if it doesn't exist
const tempDir = path.join(__dirname, 'temp');
if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
}

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, tempDir);
    },
    filename: function (req, file, cb) {
        // Create a unique filename with timestamp and original extension
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
    }
});

const upload = multer({
    storage: storage,
    limits: { 
        fileSize: 10 * 1024 * 1024, // 10MB limit
        files: 1 // Limit to 1 file per request
    },
    fileFilter: function (req, file, cb) {
        // Allow specific file types
        const filetypes = /jpeg|jpg|png|pdf|doc|docx|xls|xlsx/;
        const extname = filetypes.test(path.extname(file.originalname).toLowerCase());
        const mimetype = filetypes.test(file.mimetype);

        if (extname && mimetype) {
            return cb(null, true);
        } else {
            cb(new Error('סוג הקובץ לא נתמך. אנא העלה קובץ מסוג: PDF, Word, Excel או תמונה'));
        }
    }
});

// Serve static files from the web directory
app.use(express.static(path.join(__dirname, '../web')));

// Root route - serve the index.html
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../web/index.html'));
});

// Handle form submission
app.post('/create-draft', upload.single('file'), (req, res, next) => {
    // Log the request (without the file content)
    console.log('Received draft request:', {
        subject: req.body.subject,
        recipients: req.body.recipients,
        body: req.body.body ? req.body.body.substring(0, 50) + '...' : 'Empty body',
        file: req.file ? req.file.originalname : 'No file',
        fileSize: req.file ? (req.file.size / 1024).toFixed(2) + 'KB' : 'N/A'
    });

    const { subject, recipients, body } = req.body;
    const filePath = req.file ? req.file.path : '';

    console.log('--- POST הגיע ---');
    console.log('Subject:', subject);
    console.log('Recipients:', recipients);
    console.log('Body:', body);
    console.log('File Path:', filePath);

    // Verify the file exists if it was uploaded
    if (filePath && !fs.existsSync(filePath)) {
        console.error('File not found after upload:', filePath);
        return res.status(400).json({
            success: false,
            message: 'הקובץ לא נמצא לאחר ההעלאה. נסה שנית.'
        });
    }

    const vbsPath = path.join(__dirname, 'createDraft.vbs');
    
    // Helper function to escape command line arguments
    function escapeParam(param) {
        if (param === undefined || param === null) return '';
        return String(param)
            .replace(/\\/g, '\\\\')  // Escape backslashes
            .replace(/"/g, '\\"')      // Escape double quotes
            .replace(/[\r\n]+/g, ' ')  // Replace newlines with spaces
            .trim();                    // Remove any leading/trailing whitespace
    }

    // Build the command with proper escaping
    const command = `cscript //nologo "${vbsPath}" "${escapeParam(subject)}" "${escapeParam(recipients)}" "${escapeParam(body)}" "${escapeParam(filePath)}"`;
    
    console.log('Executing command:', command);
    console.log('Working directory:', __dirname);

    // Execute the VBScript
    exec(command, { 
        cwd: __dirname, 
        maxBuffer: 10 * 1024 * 1024,
        windowsHide: true  // Hide the command prompt window on Windows
    }, (error, stdout, stderr) => {
        // Log the command output
        console.log('Command output:', { 
            stdout: stdout || '(empty)', 
            stderr: stderr || '(empty)',
            error: error ? error.message : 'No error'
        });
        
        // Clean up the uploaded file if it exists
        const cleanupFile = () => {
            if (filePath && fs.existsSync(filePath)) {
                fs.unlink(filePath, (err) => {
                    if (err) {
                        console.error('Error deleting temp file:', err);
                    } else {
                        console.log('Temporary file deleted:', filePath);
                    }
                });
            }
        };
        
        // Always clean up the file, regardless of success/failure
        cleanupFile();

        // Handle command execution results
        if (error) {
            console.error('Error executing command:', error);
            return res.status(500).json({
                success: false,
                message: 'שגיאה בביצוע הפקודה',
                error: error.message,
                details: stderr || 'אין פרטים נוספים'
            });
        }
        
        // Check for errors in stderr
        if (stderr && stderr.trim() !== '') {
            console.error('Command stderr:', stderr);
            // Don't fail on stderr if the script still completed successfully
            if (!stdout || !stdout.includes('SUCCESS')) {
                return res.status(500).json({
                    success: false,
                    message: 'שגיאה בביצוע הסקריפט',
                    error: stderr.trim()
                });
            }
        }
        
        // Check for success message in stdout
        if (stdout && stdout.includes('SUCCESS')) {
            console.log('Draft created successfully');
            return res.json({
                success: true,
                message: 'הטיוטה נוצרה בהצלחה!'
            });
        } else {
            console.error('Failed to create draft. Output:', stdout || '(empty)');
            return res.status(500).json({
                success: false,
                message: 'נכשל ביצירת הטיוטה',
                details: (stdout || 'אין פלט מהסקריפט').substring(0, 500) // Limit response size
            });
        }
    });
});

// Error handling middleware for multer file upload errors
app.use((err, req, res, next) => {
    console.error('Server error:', err);
    
    // Handle multer errors (like file size limit exceeded)
    if (err instanceof multer.MulterError) {
        if (err.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({
                success: false,
                message: 'גודל הקובץ גדול מדי. הגבלה: 10MB'
            });
        }
        if (err.code === 'LIMIT_FILE_COUNT') {
            return res.status(400).json({
                success: false,
                message: 'ניתן להעלות קובץ אחד בלבד בכל פעם'
            });
        }
    }
    
    // Handle other errors
    res.status(500).json({
        success: false,
        message: 'אירעה שגיאה בשרת',
        error: process.env.NODE_ENV === 'development' ? err.message : undefined
    });
});

const PORT = process.env.PORT || 3000;
const server = app.listen(PORT, '0.0.0.0', () => {
    const address = server.address();
    console.log(`Server is running on http://localhost:${address.port}`);
    console.log(`Serving files from: ${path.join(__dirname, '../web')}`);
    console.log(`Temporary upload directory: ${tempDir}`);
    console.log(`Node.js version: ${process.version}`);
    console.log(`Platform: ${process.platform} ${process.arch}`);
});

// Handle server errors
server.on('error', (error) => {
    if (error.code === 'EADDRINUSE') {
        console.error(`Port ${PORT} is already in use. Please close the other application or use a different port.`);
    } else {
        console.error('Server error:', error);
    }
    process.exit(1);
});

// Handle process termination
process.on('SIGINT', () => {
    console.log('Shutting down server...');
    process.exit(0);
});
