const fs = require('fs');
const LuckyExcel = require('./dist/luckyexcel.cjs.js');

// Read test.xlsx file
const fileBuffer = fs.readFileSync('test.xlsx');

// Create a mock File object for Node.js
const mockFile = {
    name: 'test.xlsx',
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    arrayBuffer: async () => fileBuffer.buffer
};

// Add file name property
mockFile.name = 'test.xlsx';

console.log('=== Checking T83 cell during import ===\n');

// Need to simulate browser environment for transformExcelToUniver
global.FileReader = class {
    readAsArrayBuffer(file) {
        this.result = fileBuffer.buffer;
        setTimeout(() => this.onload({ target: { result: this.result } }), 0);
    }
};

// Create proper Blob/File shim
class MockFile {
    constructor(buffer, name) {
        this.buffer = buffer;
        this.name = name;
        this.type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }
}

// Import using the library to check T83
const XLSX = require('xlsx');
const workbook = XLSX.read(fileBuffer, {type: 'buffer'});

// Find DCF sheet
const dcfSheet = workbook.Sheets['DCF'];
if (dcfSheet) {
    console.log('DCF Sheet found');
    
    // Check T83 (and surrounding cells)
    const cellsToCheck = ['S83', 'T83', 'U83', 'V83', 'W83'];
    
    for (let addr of cellsToCheck) {
        if (dcfSheet[addr]) {
            const cell = dcfSheet[addr];
            console.log(`\n${addr}:`);
            if (cell.f) {
                console.log(`  Formula: ${cell.f}`);
                if (cell.f.includes('TRANSPOSE')) {
                    console.log('  ✓ Contains TRANSPOSE');
                }
            }
            if (cell.v !== undefined) {
                console.log(`  Value: ${cell.v}`);
            }
            if (cell.t) {
                console.log(`  Type: ${cell.t}`);
            }
        }
    }
}

console.log('\n=== Now checking how it gets imported to Univer ===');

// We need to check the imported Univer data structure
// This would require running transformExcelToUniver but that needs browser environment