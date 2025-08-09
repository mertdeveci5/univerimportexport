#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// Mock browser APIs
global.File = class File extends Blob {
    constructor(chunks, name, opts = {}) {
        super(chunks, opts);
        this.name = name;
    }
};

global.Blob = class Blob {
    constructor(chunks, opts = {}) {
        this.chunks = chunks;
        this.type = opts.type || '';
    }
    
    async arrayBuffer() {
        if (this.chunks[0] instanceof ArrayBuffer) return this.chunks[0];
        if (this.chunks[0] instanceof Buffer) return this.chunks[0].buffer.slice(this.chunks[0].byteOffset, this.chunks[0].byteOffset + this.chunks[0].byteLength);
        return new ArrayBuffer(0);
    }
};

const LuckyExcel = require('./dist/luckyexcel.umd.js');

const testFile = 'test.xlsx';
const fileBuffer = fs.readFileSync(testFile);
const file = new File([fileBuffer], testFile, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
});

console.log('Importing test.xlsx...\n');

LuckyExcel.transformExcelToUniver(
    file,
    (univerData) => {
        console.log('✅ Import successful\n');
        
        // Find formulas with issues
        console.log('=== SEARCHING FOR FORMULA PATTERNS ===\n');
        
        univerData.sheetOrder.forEach(sheetId => {
            const sheet = univerData.sheets[sheetId];
            if (!sheet.cellData) return;
            
            console.log(`\nSheet: ${sheet.name}`);
            console.log('-'.repeat(50));
            
            // Look for CHOOSE formulas
            const chooseFormulas = [];
            const sheetRefFormulas = [];
            
            for (const row in sheet.cellData) {
                for (const col in sheet.cellData[row]) {
                    const cell = sheet.cellData[row][col];
                    if (cell && cell.f) {
                        const cellAddr = String.fromCharCode(65 + parseInt(col)) + (parseInt(row) + 1);
                        
                        // Check for CHOOSE formulas
                        if (cell.f.includes('CHOOSE')) {
                            chooseFormulas.push({
                                cell: cellAddr,
                                formula: cell.f,
                                si: cell.si
                            });
                        }
                        
                        // Check for sheet references with +'
                        if (cell.f.includes("+'") || cell.f.includes("'+")) {
                            sheetRefFormulas.push({
                                cell: cellAddr,
                                formula: cell.f
                            });
                        }
                    }
                }
            }
            
            if (chooseFormulas.length > 0) {
                console.log('\nCHOOSE Formulas:');
                chooseFormulas.forEach(f => {
                    console.log(`  ${f.cell}: ${f.formula} ${f.si ? `(si=${f.si})` : ''}`);
                });
            }
            
            if (sheetRefFormulas.length > 0) {
                console.log('\nSheet Reference Formulas with +\':');
                sheetRefFormulas.forEach(f => {
                    console.log(`  ${f.cell}: ${f.formula}`);
                });
            }
        });
        
        // Write the imported data to file for inspection
        fs.writeFileSync('imported-data.json', JSON.stringify(univerData, null, 2));
        console.log('\n✅ Imported data written to imported-data.json');
    },
    (error) => {
        console.error('❌ Import failed:', error.message);
    }
);