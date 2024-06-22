const xlsx = require('xlsx');
const path = require('path');

// Function to read an Excel file and predict the next value in the 79th column for each row
function predictNextColumn(filePath, sequenceLength) {
    try {
        // Read the file
        const workbook = xlsx.readFile(filePath);

        // Get the first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert the sheet to JSON
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        // Initialize an array to store predicted values for each row
        const predictedValues = [];

        // Initialize a map to store frequencies of each sequence
        const sequenceFrequencies = {};

        // Build frequency table for each sequence from the first 77 columns
        for (let row = 0; row < jsonData.length; row++) {
            const rowData = jsonData[row];

            // Array to hold predicted values for current row
            const rowPredictions = [];

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

            // Predict the next value in the 79th column based on the last sequence in the data
            const lastSequence = rowData.slice(-sequenceLength).join('');
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

            // Store predicted value for current row
            rowPredictions.push(predictedValue);
            predictedValues.push(rowPredictions);
        }

        return predictedValues;
    } catch (error) {
        console.error('Error predicting the next column value:', error);
        return null;
    }
}

// Path to your Excel file (in the same folder as this script)
const excelFilePath = path.join(__dirname, 'DATA.xlsx');

// Length of the sequence to consider from the first 77 columns
const sequenceLength = 3;

// Predict the next value in the 79th column for each row based on sequences in the Excel file
const predictedValues = predictNextColumn(excelFilePath, sequenceLength);

if (predictedValues !== null) {
    predictedValues.forEach((predictedValue, index) => {
        console.log(`Predicted next value in the 79th column for row ${index + 1} is '${predictedValue}'.`);
    });
} else {
    console.log('Failed to predict the next value in the 79th column.');
}
