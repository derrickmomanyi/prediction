const xlsx = require('xlsx');
const path = require('path');

// Function to read an Excel file and predict the next letter based on the 78 inputs
function predictNextLetter(filePath, data) {
    try {
        // Read the file
        const workbook = xlsx.readFile(filePath);

        // Get the first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert the sheet to JSON
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        // Get the first row
        const firstRow = jsonData[0];

        // Initialize a map to store frequencies of each sequence
        const sequenceFrequencies = {};

        // Build frequency table for each sequence
        for (let i = 0; i < firstRow.length - 1; i++) {
            const sequence = firstRow.slice(i, i + data.sequenceLength).join('');
            const nextLetter = firstRow[i + data.sequenceLength];
            
            if (sequence.length === data.sequenceLength && data.validLetters.includes(nextLetter)) {
                if (!sequenceFrequencies[sequence]) {
                    sequenceFrequencies[sequence] = {};
                }
                if (!sequenceFrequencies[sequence][nextLetter]) {
                    sequenceFrequencies[sequence][nextLetter] = 0;
                }
                sequenceFrequencies[sequence][nextLetter]++;
            }
        }

        // Predict the next letter for the last sequence in the row
        const lastSequence = firstRow.slice(-data.sequenceLength).join('');
        const possibleNextLetters = sequenceFrequencies[lastSequence];

        if (!possibleNextLetters) {
            throw new Error('Sequence not found in data.');
        }

        let predictedLetter = null;
        let maxCount = 0;

        for (const [letter, count] of Object.entries(possibleNextLetters)) {
            if (count > maxCount) {
                predictedLetter = letter;
                maxCount = count;
            }
        }

        return predictedLetter;
    } catch (error) {
        console.error('Error predicting the next letter:', error);
        return null;
    }
}

// Path to your Excel file (in the same folder as this script)
const excelFilePath = path.join(__dirname, 'DATA.xlsx');

// Sample data of sequences and next letters
const data = {
    sequenceLength: 3, // Length of the sequence to consider
    validLetters: ['A', 'B', 'C'] // Valid letters that can be predicted
};

// Predict the next letter based on the sequences in the Excel file
const predictedLetter = predictNextLetter(excelFilePath, data);
if (predictedLetter) {
    console.log(`Predicted next letter is '${predictedLetter}'.`);
} else {
    console.log('Failed to predict the next letter.');
}
