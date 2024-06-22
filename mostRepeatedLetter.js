const xlsx = require('xlsx');
const path = require('path');

function getMostRepeatedLetterInFirstRow(filePath) {
    try {
      
        const workbook = xlsx.readFile(filePath);
       
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
       
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        // Get the first row
        const firstRow = jsonData[0];

        // Count occurrences of A, B, and C
        const letterCounts = { A: 0, B: 0, C: 0 };
        let totalCount = 0;

        firstRow.forEach(cell => {
            if (cell === 'A' || cell === 'B' || cell === 'C') {
                letterCounts[cell]++;
            }           
            totalCount++;
        });
       
        let mostRepeatedLetter = null;
        let maxCount = 0;

        for (const [letter, count] of Object.entries(letterCounts)) {
            if (count > maxCount) {
                mostRepeatedLetter = letter;
                maxCount = count;
            }
        }

        return { letter: mostRepeatedLetter, count: maxCount, totalCount: totalCount };
    } catch (error) {
        console.error('Error processing the Excel file:', error);
        return null;
    }
}

const excelFilePath = path.join(__dirname, 'DATA.xlsx');

// Get the most repeated letter and its count in the first row
const result = getMostRepeatedLetterInFirstRow(excelFilePath);
if (result) {
    console.log(`The most repeated letter in the first row is '${result.letter}' with a count of ${result.count}.`);
    console.log(`Total count of items in the first row: ${result.totalCount}.`);
} else {
    console.log('Failed to determine the most repeated letter.');
}
