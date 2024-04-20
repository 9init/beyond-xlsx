const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Define source and destination folders
const sourceFolder = 'folder/contains/files/to/move/';
const destinationFolder = path.join(sourceFolder, "new");

// Create destination folder if it doesn't exist
if (!fs.existsSync(destinationFolder)) {
    fs.mkdirSync(destinationFolder);
}

// Read Excel file
const workbook = xlsx.readFile('/path/to/excel/file.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Extract file names from first column
const records = xlsx.utils.sheet_to_json(worksheet, { header: 'A' });

// Move files from source to destination folder
records.forEach(({ A }) => {
    const sourceFilePath = path.join(sourceFolder, A);
    const destinationFilePath = path.join(destinationFolder, A);

    if (fs.existsSync(sourceFilePath)) {
        if (!fs.existsSync(destinationFolder)) {
            fs.mkdirSync(destinationFolder, { recursive: true });
        }
        fs.renameSync(sourceFilePath, destinationFilePath);
    }
});