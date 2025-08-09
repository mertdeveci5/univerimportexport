#!/usr/bin/env node

/**
 * Test export functionality after shared formula fix
 */

const fs = require('fs');
const LuckyExcel = require('./dist/luckyexcel.umd.js');

console.log('Testing export with shared formula fix...\n');

// Create a File-like object for Node.js
class NodeFile {
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
}

// Read the original test file
const fileBuffer = fs.readFileSync('test.xlsx');
const file = new NodeFile(Buffer.from(fileBuffer), 'test.xlsx');

console.log('Step 1: Importing test.xlsx to Univer format...');

LuckyExcel.transformExcelToUniver(
    file,
    (univerData) => {
        console.log('✅ Import successful\n');
        
        // Log some details about the data
        console.log('Workbook details:');
        console.log(`  Sheets: ${univerData.sheetOrder.length}`);
        console.log(`  Sheet names: ${univerData.sheetOrder.map(id => univerData.sheets[id].name).join(', ')}`);
        
        // Check for shared formulas in Operational Assumptions sheet
        const opSheet = Object.values(univerData.sheets).find(s => s.name === 'Operational Assumptions');
        if (opSheet && opSheet.cellData) {
            let sharedFormulaCount = 0;
            for (const row in opSheet.cellData) {
                for (const col in opSheet.cellData[row]) {
                    const cell = opSheet.cellData[row][col];
                    if (cell && cell.si) {
                        sharedFormulaCount++;
                    }
                }
            }
            console.log(`  Shared formulas in Operational Assumptions: ${sharedFormulaCount}`);
        }
        
        console.log('\nStep 2: Exporting to Excel with fixed shared formula handling...');
        
        // Export back to Excel
        LuckyExcel.transformUniverToExcel({
            snapshot: univerData,
            fileName: 'test-fixed.xlsx',
            getBuffer: true,
            success: (buffer) => {
                if (buffer) {
                    fs.writeFileSync('test-fixed.xlsx', buffer);
                    console.log('✅ Export successful: test-fixed.xlsx created\n');
                    
                    console.log('Step 3: Running formula comparison analysis...');
                    
                    // Run the analysis tool to check if formulas are preserved
                    const { exec } = require('child_process');
                    exec('node analyze-fixed-export.js', (error, stdout, stderr) => {
                        if (error) {
                            console.error('❌ Analysis failed:', error);
                            return;
                        }
                        console.log(stdout);
                        if (stderr) console.error(stderr);
                    });
                } else {
                    console.error('❌ Export failed: No buffer returned');
                }
            },
            error: (err) => {
                console.error('❌ Export failed:', err.message);
            }
        });
    },
    (error) => {
        console.error('❌ Import failed:', error.message);
    }
);