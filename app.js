// Excel Helper - Web Version
// Converts comma decimals (11,2) to dot format (11.2)

let workbook = null;
let fileName = '';
let cellsToConvert = 0;

// DOM Elements
const dropZone = document.getElementById('dropZone');
const dropHint = document.getElementById('dropHint');
const fileInfo = document.getElementById('fileInfo');
const fileNameEl = document.getElementById('fileName');
const fileStatsEl = document.getElementById('fileStats');
const statusEl = document.getElementById('status');
const exportBtn = document.getElementById('exportBtn');
const fileInput = document.getElementById('fileInput');
const fileInputAlt = document.getElementById('fileInputAlt');

// Event Listeners
dropZone.addEventListener('dragover', handleDragOver);
dropZone.addEventListener('dragleave', handleDragLeave);
dropZone.addEventListener('drop', handleDrop);
dropZone.addEventListener('click', (e) => {
    if (e.target === dropZone || e.target.closest('.drop-hint')) {
        fileInput.click();
    }
});

fileInput.addEventListener('change', handleFileSelect);
fileInputAlt.addEventListener('change', handleFileSelect);
exportBtn.addEventListener('click', handleExport);

// Drag & Drop Handlers
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        loadFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        loadFile(files[0]);
    }
}

// File Loading
function loadFile(file) {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    
    if (!validExtensions.includes(ext)) {
        showStatus('❌ Please select a valid file (.xlsx, .xls or .csv)', true);
        return;
    }
    
    fileName = file.name;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array', raw: true });
            
            cellsToConvert = countCellsToConvert();
            
            // Update UI
            dropHint.classList.add('hidden');
            fileInfo.classList.remove('hidden');
            fileNameEl.textContent = fileName;
            fileStatsEl.textContent = `To convert: ${cellsToConvert} cells`;
            exportBtn.disabled = false;
            
            showStatus('✓ File loaded successfully', false);
        } catch (error) {
            showStatus('❌ Error loading file: ' + error.message, true);
        }
    };
    
    reader.onerror = function() {
        showStatus('❌ Error reading file', true);
    };
    
    reader.readAsArrayBuffer(file);
}

// Check if value is a comma decimal number (e.g., "11,2")
function isCommaDecimal(value) {
    if (typeof value !== 'string') return false;
    return /^-?\d+,\d+$/.test(value.trim());
}

// Count cells that need conversion
function countCellsToConvert() {
    let count = 0;
    
    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
        
        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = sheet[cellRef];
                
                if (cell && cell.v !== undefined) {
                    const value = String(cell.v);
                    if (isCommaDecimal(value)) {
                        count++;
                    }
                }
            }
        }
    });
    
    return count;
}

// Convert all comma decimals to dot format
function convertAllCells() {
    let count = 0;
    
    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
        
        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = sheet[cellRef];
                
                if (cell && cell.v !== undefined) {
                    const value = String(cell.v);
                    if (isCommaDecimal(value)) {
                        // Convert comma to dot and keep as string
                        const newValue = value.replace(',', '.');
                        cell.v = newValue;
                        cell.t = 's'; // Set type to string
                        cell.w = newValue; // Set formatted value
                        count++;
                    }
                }
            }
        }
    });
    
    return count;
}

// Export Handler
async function handleExport() {
    if (!workbook) {
        showStatus('❌ Please load a file first', true);
        return;
    }
    
    try {
        // Create new ExcelJS workbook
        const excelWorkbook = new ExcelJS.Workbook();
        let convertedCount = 0;
        
        // Process each sheet
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const excelSheet = excelWorkbook.addWorksheet(sheetName);
            
            const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
            
            for (let row = range.s.r; row <= range.e.r; row++) {
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
                    const cell = sheet[cellRef];
                    
                    if (cell && cell.v !== undefined) {
                        let value = String(cell.v);
                        const excelCell = excelSheet.getCell(row + 1, col + 1);
                        
                        // Convert comma decimals to dot
                        if (isCommaDecimal(value)) {
                            value = value.replace(',', '.');
                            convertedCount++;
                        }
                        
                        excelCell.value = value;
                        // Right-align all cells
                        excelCell.alignment = { horizontal: 'right' };
                    }
                }
            }
        });
        
        // Generate new filename
        const baseName = fileName.replace(/\.[^/.]+$/, '');
        const newFileName = baseName + '_converted.xlsx';
        
        // Write and download
        const buffer = await excelWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        // Create download link
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = newFileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showStatus(`✅ ${convertedCount} cells converted and saved!`, false);
        
        // Update stats
        cellsToConvert = 0;
        fileStatsEl.textContent = 'To convert: 0 cells';
        
    } catch (error) {
        showStatus('❌ Save error: ' + error.message, true);
    }
}

// Status Display
function showStatus(message, isError) {
    statusEl.textContent = message;
    statusEl.className = 'status ' + (isError ? 'error' : 'success');
}

