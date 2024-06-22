const xlsx = require('xlsx');
const path = require('path');

// Function to read an Excel file, predict the next value in the 79th column, and return both predicted value and data
function predictNextColumn(filePath, sequenceLength) {
    try {
        // Read the file
        const workbook = xlsx.readFile(filePath);

        // Get the first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert the sheet to JSON
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header:  1});

        // Initialize a map to store frequencies of each sequence
        const sequenceFrequencies = {};

        // Build frequency table for each sequence from the first 77 columns
        for (let row = 0; row < jsonData.length; row++) {
            const rowData = jsonData[row];

            for (let i = 0; i <= rowData.length - sequenceLength; i++) {
                const sequence = rowData.slice(i, i + sequenceLength).join('');
                const nextColumnValue = rowData[i + sequenceLength];

                if (!sequenceFrequencies[sequence]) {
                    sequenceFrequencies[sequence] = {};
                }

                if (!sequenceFrequencies[sequence][nextColumnValue]) {
                    sequenceFrequencies[sequence][nextColumnValue] = 0;
                }

                sequenceFrequencies[sequence][nextColumnValue]++;
            }
        }

        // Predict the next value in the 79th column based on the last sequence in the data
        const lastRow = jsonData[jsonData.length - 1];
        const lastSequence = lastRow.slice(-sequenceLength).join('');
        const possibleNextValues = sequenceFrequencies[lastSequence];

        if (!possibleNextValues) {
            throw new Error('Sequence not found in data.');
        }

        let predictedValue = null;
        let maxCount = 0;

        for (const [value, count] of Object.entries(possibleNextValues)) {
            if (count > maxCount) {
                predictedValue = value;
                maxCount = count;
            }
        }

        // Return both predicted value and the entire jsonData
        return { predictedValue, jsonData };
    } catch (error) {
        console.error('Error predicting the next column value:', error);
        return { predictedValue: null, jsonData: null };
    }
}

// Path to your Excel file (in the same folder as this script)
const excelFilePath = path.join(__dirname, 'DATA.xlsx');

// Length of the sequence to consider from the first 77 columns
const sequenceLength = 3;

// Predict the next value in the 79th column based on sequences in the Excel file
const { predictedValue, jsonData } = predictNextColumn(excelFilePath, sequenceLength);

if (predictedValue !== null && jsonData !== null) {
    const predictedRow = jsonData.length;
    console.log(`Predicted next value in the 79th column at row ${predictedRow} is '${predictedValue}'.`);
} else {
    console.log('Failed to predict the next value in the 79th column.');
}
