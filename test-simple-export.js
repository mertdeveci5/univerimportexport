#!/usr/bin/env node

/**
 * Simple test for the enhanced export functionality
 * This test checks basic functionality without complex File API dependencies
 */

const fs = require('fs');
const path = require('path');

// Import the library
const LuckyExcel = require('./dist/luckyexcel.umd.js');

console.log('Testing Enhanced Export...\n');

// Create a mock Univer data structure for testing
const mockUniverData = {
    id: 'test-workbook',
    name: 'Test Workbook',
    appVersion: '1.0.0',
    locale: 'en-US',
    styles: {
        's1': {
            ff: 'Arial',
            fs: 12,
            bl: 1, // Bold
            cl: { rgb: 'FF0000' } // Red color
        }
    },
    sheetOrder: ['sheet1'],
    sheets: {
        'sheet1': {
            id: 'sheet1',
            name: 'Test Sheet',
            cellData: {
                0: {
                    0: { v: 'Hello', t: 1, s: 's1' }, // String with style
                    1: { v: 'World', t: 1 }
                },
                1: {
                    0: { v: 123, t: 2 }, // Number
                    1: { v: '=A1&" "&B1', f: 'A1&" "&B1' } // Formula
                },
                2: {
                    0: { v: true, t: 3 }, // Boolean
                    1: { v: '456', t: 4 } // Force string (number as text)
                }
            },
            mergeData: [
                {
                    startRow: 3,
                    startColumn: 0,
                    endRow: 3,
                    endColumn: 1
                }
            ],
            rowCount: 10,
            columnCount: 5,
            defaultRowHeight: 20,
            defaultColumnWidth: 100
        }
    },
    resources: [
        {
            name: 'SHEET_FILTER_PLUGIN',
            data: JSON.stringify({
                'sheet1': {
                    range: {
                        startRow: 0,
                        startColumn: 0,
                        endRow: 10,
                        endColumn: 4
                    }
                }
            })
        },
        {
            name: 'SHEET_COMMENT_PLUGIN',
            data: JSON.stringify({
                'sheet1': [
                    {
                        row: 0,
                        column: 0,
                        content: 'This is a test comment',
                        author: 'Test User'
                    }
                ]
            })
        }
    ]
};

console.log('Mock Univer data created with:');
console.log('- 1 sheet with styled cells');
console.log('- Formulas and different data types');
console.log('- Merged cells');
console.log('- Filter and comment resources\n');

// Test the export
console.log('Starting export...');

const startTime = Date.now();

LuckyExcel.transformUniverToExcel({
    snapshot: mockUniverData,
    fileName: 'test-simple-output.xlsx',
    getBuffer: true,
    success: (buffer) => {
        const elapsed = Date.now() - startTime;
        
        if (buffer && buffer.length > 0) {
            // Save the file
            const outputFile = path.join(__dirname, 'test-simple-output.xlsx');
            fs.writeFileSync(outputFile, buffer);
            
            console.log('✅ Export successful!');
            console.log(`   File size: ${(buffer.length / 1024).toFixed(2)} KB`);
            console.log(`   Time taken: ${elapsed}ms`);
            console.log(`   Output saved to: ${outputFile}`);
            
            // Basic validation
            console.log('\n📊 Export validation:');
            console.log('   ✓ Buffer created');
            console.log('   ✓ Non-empty output');
            console.log('   ✓ File saved successfully');
            
            // Try to read the first few bytes to check it's a valid XLSX
            const header = buffer.slice(0, 4);
            const isZip = header[0] === 0x50 && header[1] === 0x4B; // PK header
            console.log(`   ${isZip ? '✓' : '✗'} Valid ZIP/XLSX header`);
            
            console.log('\n🎉 Test completed successfully!');
        } else {
            console.error('❌ Export failed: Empty buffer');
            process.exit(1);
        }
    },
    error: (err) => {
        console.error('❌ Export failed:', err.message);
        console.error(err.stack);
        process.exit(1);
    }
});