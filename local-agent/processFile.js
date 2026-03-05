const fs = require('fs');

function checkFile(filePath) {
    if (!filePath || !fs.existsSync(filePath)) {
        throw new Error('File does not exist: ' + filePath);
    }
    return filePath;
}

module.exports = { checkFile };
