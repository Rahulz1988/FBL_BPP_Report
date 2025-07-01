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

let processedData = null;

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
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
    const fileInput = document.getElementById('fileInput');
    const uploadBtn = document.getElementById('uploadBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    
    uploadBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileUpload);
    downloadBtn.addEventListener('click', downloadProcessedFile);
}

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    updateStatus('Processing file...', 'processing');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            processData(jsonData);
        } catch (error) {
            updateStatus('Error reading file: ' + error.message, 'error');
        }
    };
    reader.readAsBinaryString(file);
}

function processData(data) {
    try {
        processedData = data.map(row => {
            const processedRow = { ...row };
            
            // Calculate sales weighted score
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
            
            // Round percentages: if .5 or more, round up; if below .5, round down
            processedRow['Sales %'] = Math.round(salesPercent);
            processedRow['Ops %'] = Math.round(opsPercent);
            processedRow['Suitability'] = opsPercent > salesPercent ? 'Operations' : 'Sales';
            
            return processedRow;
        });
        
        updateStatus(`Successfully processed ${processedData.length} records`, 'success');
        document.getElementById('downloadBtn').style.display = 'inline-block';
        
    } catch (error) {
        updateStatus('Error processing data: ' + error.message, 'error');
    }
}

function updateStatus(message, type) {
    const status = document.getElementById('status');
    const resultsSection = document.getElementById('resultsSection');
    
    if (status) {
        status.textContent = message;
        status.className = `status ${type}`;
    }
    
    if (resultsSection) {
        resultsSection.style.display = 'block';
    }
}

function downloadProcessedFile() {
    if (!processedData) {
        updateStatus('No processed data available', 'error');
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
        const filename = 'competency_assessment_processed.xlsx';
        XLSX.writeFile(wb, filename);
        
        updateStatus('File downloaded successfully!', 'success');
        
    } catch (error) {
        updateStatus('Error downloading file: ' + error.message, 'error');
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