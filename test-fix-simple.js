#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// First, import the file to get Univer data
const { execSync } = require('child_process');

console.log('Testing shared formula fix...\n');

// Create a simple test to import and export
const testCode = `
const fs = require('fs');
const LuckyExcel = require('./dist/luckyexcel.umd.js');

// Helper to handle File API in Node
global.File = class File {
    constructor(buffer, name) {
        this.buffer = buffer;
        this.name = name;
        this.type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        this.size = buffer.length;
    }
    
    arrayBuffer() {
        return Promise.resolve(this.buffer.buffer.slice(
            this.buffer.byteOffset, 
            this.buffer.byteOffset + this.buffer.byteLength
        ));
    }
};

const fileBuffer = fs.readFileSync('test.xlsx');
const file = new File(fileBuffer, 'test.xlsx');

LuckyExcel.transformExcelToUniver(
    file,
    (univerData) => {
        console.log('Import successful');
        
        // Now export
        LuckyExcel.transformUniverToExcel({
            snapshot: univerData,
            fileName: 'test-fixed.xlsx',
            getBuffer: true,
            success: (buffer) => {
                if (buffer) {
                    fs.writeFileSync('test-fixed.xlsx', buffer);
                    console.log('Export successful: test-fixed.xlsx created');
                }
            },
            error: (err) => {
                console.error('Export failed:', err.message);
                process.exit(1);
            }
        });
    },
    (error) => {
        console.error('Import failed:', error.message);
        process.exit(1);
    }
);
`;

// Write and execute the test
fs.writeFileSync('temp-test.js', testCode);

try {
    execSync('node temp-test.js', { stdio: 'inherit' });
    
    // If export was successful, analyze the result
    if (fs.existsSync('test-fixed.xlsx')) {
        console.log('\nAnalyzing the exported file...\n');
        execSync('node analyze-fixed-export.js', { stdio: 'inherit' });
    }
} catch (error) {
    console.error('Test failed:', error.message);
} finally {
    // Clean up
    if (fs.existsSync('temp-test.js')) {
        fs.unlinkSync('temp-test.js');
    }
}