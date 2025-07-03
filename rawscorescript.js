// Global variable to store raw score processed data
let processedDataRaw = null;

// Initialize raw score event listeners when DOM is loaded
document.addEventListener('DOMContentLoaded', function () {
    initializeRawScoreEventListeners();
});

function initializeRawScoreEventListeners() {
    // Raw Score Elements
    const fileInputRaw = document.getElementById('fileInputRaw');
    const uploadBtnRaw = document.getElementById('uploadBtnRaw');
    const downloadBtnRaw = document.getElementById('downloadBtnRaw');

    // Raw Score Event Listeners
    if (uploadBtnRaw) {
        uploadBtnRaw.addEventListener('click', () => fileInputRaw.click());
    }
    if (fileInputRaw) {
        fileInputRaw.addEventListener('change', (e) => handleRawFileUpload(e));
    }
    if (downloadBtnRaw) {
        downloadBtnRaw.addEventListener('click', () => downloadRawProcessedFile());
    }
}

function handleRawFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    updateStatus('Processing raw score file...', 'processing', 'statusRaw');

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

            // Validate and process raw score data
            validateRawScoreData(cleanedData);
            processDataRaw(cleanedData);

        } catch (error) {
            updateStatus('Error reading raw score file: ' + error.message, 'error', 'statusRaw');
        }
    };
    reader.readAsBinaryString(file);
}

function processDataRaw(data) {
    try {
        // Validate required columns
        if (!data || data.length === 0) {
            throw new Error('No data found in the uploaded file');
        }

        const firstRow = data[0];
        const requiredColumns = ['User Mail', 'section', 'score'];
        const missingColumns = requiredColumns.filter(col => !(col in firstRow));
        
        if (missingColumns.length > 0) {
            throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
        }

        // Group data by User Mail and section, then sum scores for each section
        const candidateScores = {};
        
        data.forEach((row, index) => {
            const userMail = row['User Mail'];
            const section = row['section'];
            const score = parseFloat(row['score']);
            
            if (!userMail || !section) {
                console.warn(`Skipping row ${index + 1}: Missing User Mail or section`);
                return;
            }
            
            if (isNaN(score)) {
                console.warn(`Skipping row ${index + 1}: Invalid score "${row['score']}"`);
                return;
            }
            
            // Initialize candidate if not exists
            if (!candidateScores[userMail]) {
                candidateScores[userMail] = {};
            }
            
            // Initialize section total if not exists, otherwise add to existing total
            if (!candidateScores[userMail][section]) {
                candidateScores[userMail][section] = 0;
            }
            
            // Add score to the section total
            candidateScores[userMail][section] += score;
        });

        // Debug: Log the aggregated scores
        console.log('Aggregated Raw Scores:', candidateScores);

        // Use the summed raw scores directly as final scores and calculate Sales%, Ops%
        const processedCandidates = [];
        
        Object.keys(candidateScores).forEach(userMail => {
            const candidateData = { 'User Mail': userMail };
            
            // For each competency, use the total raw score directly as the final score
            competencies.forEach(competency => {
                const totalRawScore = candidateScores[userMail][competency] || 0;
                
                // Verify that the summed raw score is within the valid range (5-20)
                if (totalRawScore < 5 || totalRawScore > 20) {
                    throw new Error(
                        `Invalid summed score "${totalRawScore}" for "${competency}" for candidate "${userMail}". ` +
                        `Summed scores must be between 5 and 20 (inclusive).`
                    );
                }
                
                candidateData[competency] = totalRawScore; // Use raw score sum directly
            });
            
            processedCandidates.push(candidateData);
        });

        // Now calculate Sales% and Ops% for each candidate
        processedDataRaw = processedCandidates.map((candidate, index) => {
            const processedRow = { ...candidate };
            
            // Calculate sales and ops weighted scores
            let salesScore = 0;
            let opsScore = 0;

            competencies.forEach((competency, index) => {
                const score = processedRow[competency] || 0; // Use 0 as default if missing
                salesScore += (score * salesWeightages[index]) / 100;
                opsScore += (score * opsWeightages[index]) / 100;
            });

            const totalScore = salesScore + opsScore;
            const salesPercent = totalScore > 0 ? (salesScore / totalScore) * 100 : 0;
            const opsPercent = totalScore > 0 ? (opsScore / totalScore) * 100 : 0;

            // Round percentages
            const salesRounded = Math.round(salesPercent);
            const opsRounded = Math.round(opsPercent);

            processedRow['Sales %'] = salesRounded;
            processedRow['Ops %'] = opsRounded;

            if (opsRounded > salesRounded) {
                processedRow['Suitability'] = 'Operations';
            } else if (salesRounded > opsRounded) {
                processedRow['Suitability'] = 'Sales';
            } else {
                processedRow['Suitability'] = 'Operations/Sales';
            }

            return processedRow;
        });

        // Debug: Log the final processed data
        console.log('Final Processed Data:', processedDataRaw);

        updateStatus(`Successfully processed ${processedDataRaw.length} candidates from raw scores`, 'success', 'statusRaw');
        document.getElementById('downloadBtnRaw').style.display = 'inline-block';

    } catch (error) {
        updateStatus('Error processing raw data: ' + error.message, 'error', 'statusRaw');
    }
}

// Enhanced validation for raw score data
function validateRawScoreData(data) {
    if (!Array.isArray(data) || data.length === 0) {
        throw new Error('No data found in the uploaded file');
    }

    const firstRow = data[0];
    const requiredColumns = ['User Mail', 'section', 'score'];
    const missingColumns = requiredColumns.filter(col => !(col in firstRow));

    if (missingColumns.length > 0) {
        throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
    }

    // Check if we have valid data
    const validRows = data.filter(row => 
        row['User Mail'] && 
        row['section'] && 
        !isNaN(parseFloat(row['score']))
    );

    if (validRows.length === 0) {
        throw new Error('No valid data rows found. Please check that User Mail, section, and score columns contain valid data.');
    }

    // Check if we have recognized competencies
    const uniqueSections = [...new Set(data.map(row => row['section']).filter(Boolean))];
    const unrecognizedSections = uniqueSections.filter(section => !competencies.includes(section));
    
    if (unrecognizedSections.length > 0) {
        console.warn('Unrecognized sections found (will be ignored):', unrecognizedSections);
    }

    return true;
}

function downloadRawProcessedFile() {
    if (!processedDataRaw) {
        updateStatus('No processed raw score data available', 'error', 'statusRaw');
        return;
    }

    try {
        // Create a new workbook
        const wb = XLSX.utils.book_new();

        // Convert processed data to worksheet
        const ws = XLSX.utils.json_to_sheet(processedDataRaw);

        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Raw Score Processed Data');

        // Generate Excel file and trigger download
        const filename = 'competency_assessment_raw_processed.xlsx';
        XLSX.writeFile(wb, filename);

        updateStatus('Raw score file downloaded successfully!', 'success', 'statusRaw');

    } catch (error) {
        updateStatus('Error downloading raw score file: ' + error.message, 'error', 'statusRaw');
    }
}