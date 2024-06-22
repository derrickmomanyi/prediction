const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

function readExcelFile(filePath) {   
    if (!fs.existsSync(filePath)) {
        console.error('File not found:', filePath);
        return;
    }
    
    const workbook = xlsx.readFile(filePath);
   
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    return jsonData;
}


const excelFilePath = path.join(__dirname, 'DATA.xlsx');

// Read and display the Excel file data
const excelData = readExcelFile(excelFilePath);
console.log(excelData);