// Competency names and weightages
const competencies = [
    'Being Innovative',
    'Collaborative Working',
    'Communication skills',
    'Process Orientation',
    'Drive for Results',
    'Planning & organizing',
    'Product & Market Understanding',
    'Resilience',
    'Self Confidence/Driven',
    'Understanding Customer Needs and effective resolution'
];

const salesWeightages = [5.5, 5.5, 13.0, 5.5, 13.0, 5.5, 13.0, 13.0, 13.0, 13.0];
const opsWeightages = [5.0, 15.0, 5.0, 15.0, 15.0, 15.0, 5.0, 5.0, 15.0, 5.0];

let processedDataFinal = null;
// Removed duplicate processedDataRaw declaration - it's handled in rawscorescript.js

// Initialize the application
document.addEventListener('DOMContentLoaded', function () {
    initializeWeightageDisplay();
    initializeEventListeners();
});

function initializeWeightageDisplay() {
    const salesList = document.getElementById('salesWeightages');
    const opsList = document.getElementById('opsWeightages');

    competencies.forEach((competency, index) => {
        const salesItem = document.createElement('li');
        salesItem.innerHTML = `<span>${competency}</span><span>${salesWeightages[index]}%</span>`;
        salesList.appendChild(salesItem);

        const opsItem = document.createElement('li');
        opsItem.innerHTML = `<span>${competency}</span><span>${opsWeightages[index]}%</span>`;
        opsList.appendChild(opsItem);
    });
}

function initializeEventListeners() {
    // Final Score Elements
    const fileInputFinal = document.getElementById('fileInputFinal');
    const uploadBtnFinal = document.getElementById('uploadBtnFinal');
    const downloadBtnFinal = document.getElementById('downloadBtnFinal');

    // Final Score Event Listeners
    if (uploadBtnFinal) {
        uploadBtnFinal.addEventListener('click', () => fileInputFinal.click());
    }
    if (fileInputFinal) {
        fileInputFinal.addEventListener('change', (e) => handleFileUpload(e, 'final'));
    }
    if (downloadBtnFinal) {
        downloadBtnFinal.addEventListener('click', () => downloadProcessedFile('final'));
    }

    // Raw Score Event Listeners are handled in rawscorescript.js
    // Removed duplicate event listeners to avoid conflicts
}

function handleFileUpload(event, type) {
    const file = event.target.files[0];
    if (!file) return;

    // Only handle final score processing here
    if (type !== 'final') return;

    const statusId = 'statusFinal';
    updateStatus('Processing file...', 'processing', statusId);

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

            // Convert to JSON
            const rawData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

            // Normalize keys (trim headers)
            const cleanedData = rawData.map(row => {
                const cleanedRow = {};
                Object.keys(row).forEach(key => {
                    cleanedRow[key.trim()] = row[key]; // trim header names
                });
                return cleanedRow;
            });

            // Process final score data
            validateInputData(cleanedData);
            processDataFinal(cleanedData);

        } catch (error) {
            updateStatus('Error reading file: ' + error.message, 'error', statusId);
        }
    };
    reader.readAsBinaryString(file);
}

function processDataFinal(data) {
    try {
        processedDataFinal = data.map((row, rowIndex) => {
            const processedRow = { ...row };

            // Validate each competency score is between 5 and 20 (inclusive)
            for (const competency of competencies) {
                const score = parseFloat(row[competency]);
                if (isNaN(score) || score < 5 || score > 20) {
                    throw new Error(
                        `Invalid score "${row[competency]}" for "${competency}" in row ${rowIndex + 1}. Scores must be between 5 and 20.`
                    );
                }
            }

            // Calculate sales and ops weighted scores
            let salesScore = 0;
            let opsScore = 0;

            competencies.forEach((competency, index) => {
                const score = parseFloat(row[competency]) || 0;
                salesScore += (score * salesWeightages[index]) / 100;
                opsScore += (score * opsWeightages[index]) / 100;
            });

            const totalScore = salesScore + opsScore;
            const salesPercent = totalScore > 0 ? (salesScore / totalScore) * 100 : 0;
            const opsPercent = totalScore > 0 ? (opsScore / totalScore) * 100 : 0;

            // Use two decimal places instead of rounding
            const salesDecimal = Math.round(salesPercent * 100) / 100;
            const opsDecimal = Math.round(opsPercent * 100) / 100;

            processedRow['Sales %'] = salesDecimal;
            processedRow['Ops %'] = opsDecimal;

            if (opsDecimal > salesDecimal) {
                processedRow['Suitability'] = 'Operations';
            } else if (salesDecimal > opsDecimal) {
                processedRow['Suitability'] = 'Sales';
            } else {
                processedRow['Suitability'] = 'Operations/Sales';
            }

            return processedRow;
        });

        updateStatus(`Successfully processed ${processedDataFinal.length} records`, 'success', 'statusFinal');
        document.getElementById('downloadBtnFinal').style.display = 'inline-block';

    } catch (error) {
        updateStatus('Error processing data: ' + error.message, 'error', 'statusFinal');
    }
}

// Removed the conflicting processDataRaw function - it's properly implemented in rawscorescript.js

function updateStatus(message, type, statusId) {
    const status = document.getElementById(statusId);
    const resultsSectionId = statusId === 'statusFinal' ? 'resultsSectionFinal' : 'resultsSectionRaw';
    const resultsSection = document.getElementById(resultsSectionId);

    if (status) {
        status.textContent = message;
        status.className = `status ${type}`;
    }

    if (resultsSection) {
        resultsSection.style.display = 'block';
    }
}

function downloadProcessedFile(type) {
    // Only handle final score downloads here
    if (type !== 'final') return;
    
    const processedData = processedDataFinal;
    const statusId = 'statusFinal';
    
    if (!processedData) {
        updateStatus('No processed data available', 'error', statusId);
        return;
    }

    try {
        // Create a new workbook
        const wb = XLSX.utils.book_new();

        // Convert processed data to worksheet
        const ws = XLSX.utils.json_to_sheet(processedData);

        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Processed Data');

        // Generate Excel file and trigger download
        const filename = 'competency_assessment_final_processed.xlsx';
        XLSX.writeFile(wb, filename);

        updateStatus('File downloaded successfully!', 'success', statusId);

    } catch (error) {
        updateStatus('Error downloading file: ' + error.message, 'error', statusId);
    }
}

// Utility function to validate input data
function validateInputData(data) {
    if (!Array.isArray(data) || data.length === 0) {
        throw new Error('No data found in the uploaded file');
    }

    const firstRow = data[0];
    const missingCompetencies = competencies.filter(comp => !(comp in firstRow));

    if (missingCompetencies.length > 0) {
        throw new Error(`Missing competencies in the uploaded file: ${missingCompetencies.join(', ')}`);
    }

    return true;
}