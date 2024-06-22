const xlsx = require('xlsx');
const path = require('path');

// Function to read an Excel file and predict the next letter based on the 78 inputs for each row
function predictNextLetters(filePath, data) {
    try {
        // Read the file
        const workbook = xlsx.readFile(filePath);

        // Get the first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert the sheet to JSON
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        // Initialize an array to store predicted letters
        const predictions = [];
        let rowCount = 0;

        // Process each row
        jsonData.forEach(row => {
            // Initialize a map to store frequencies of each sequence
            const sequenceFrequencies = {};

            // Build frequency table for each sequence in the current row
            for (let i = 0; i < row.length - data.sequenceLength; i++) {
                const sequence = row.slice(i, i + data.sequenceLength).join('');
                const nextLetter = row[i + data.sequenceLength];
                
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
            const lastSequence = row.slice(-data.sequenceLength).join('');
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

            // Store the predicted letter for the current row
            predictions.push(predictedLetter);
            rowCount++;
        });

        return { predictions, rowCount };
    } catch (error) {
        console.error('Error predicting next letters:', error);
        return null;
    }
}

// Path to your Excel file (in the same folder as this script)
const excelFilePath = path.join(__dirname, 'DATA.xlsx');

// Sample data of sequences and next letters
const data = {
    sequenceLength: 5, // Length of the sequence to consider
    validLetters: ['A', 'B', 'C'] // Valid letters that can be predicted
};

// Predict the next letters for all rows based on the sequences in the Excel file
const { predictions, rowCount } = predictNextLetters(excelFilePath, data);
if (predictions && rowCount > 0) {
    console.log(`Predicted letters for all ${rowCount} rows:`);
    console.log(predictions);
} else {
    console.log('Failed to predict next letters for all rows.');
}
