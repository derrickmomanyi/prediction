const xlsx = require('xlsx');
const path = require('path');
const tf = require('@tensorflow/tfjs-node');

async function predictNextLetters(filePath, data) {
    try {
        // Read the file
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        // Prepare sequences and labels
        const sequences = [];
        const labels = [];

        jsonData.forEach(row => {
            for (let i = 0; i < row.length - data.sequenceLength; i++) {
                const sequence = row.slice(i, i + data.sequenceLength);
                const nextLetter = row[i + data.sequenceLength];

                if (sequence.length === data.sequenceLength && data.validLetters.includes(nextLetter)) {
                    sequences.push(sequence);
                    labels.push(nextLetter);
                }
            }
        });

        // Convert sequences and labels to tensors
        const xs = tf.tensor2d(sequences);
        const ys = tf.tensor1d(labels.map(letter => data.validLetters.indexOf(letter)));

        // Define and train LSTM model
        const model = tf.sequential();
        model.add(tf.layers.lstm({ units: 50, inputShape: [data.sequenceLength, 1] }));
        model.add(tf.layers.dense({ units: data.validLetters.length, activation: 'softmax' }));
        model.compile({ optimizer: 'adam', loss: 'sparseCategoricalCrossentropy' });

        await model.fit(xs, ys, { epochs: 50 });

        // Predict next letters for all rows
        const predictions = [];
        jsonData.forEach(row => {
            const sequence = row.slice(-data.sequenceLength);
            const sequenceTensor = tf.tensor2d(sequence).expandDims();
            const prediction = model.predict(sequenceTensor);
            const predictedIndex = prediction.argMax(-1).dataSync()[0];
            const predictedLetter = data.validLetters[predictedIndex];
            predictions.push(predictedLetter);
        });

        return { predictions, rowCount: jsonData.length };
    } catch (error) {
        console.error('Error predicting next letters:', error);
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

// Predict the next letters for all rows based on the sequences in the Excel file
predictNextLetters(excelFilePath, data)
    .then(result => {
        if (result && result.predictions.length > 0) {
            console.log(`Predicted letters for all ${result.rowCount} rows:`);
            console.log(result.predictions);
        } else {
            console.log('Failed to predict next letters for all rows.');
        }
    })
    .catch(err => {
        console.error('Error predicting next letters:', err);
    });
