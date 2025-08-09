#!/usr/bin/env node

/**
 * Test export with proper File simulation for Node.js
 */

const fs = require('fs');
const XLSX = require('xlsx');

// Mock browser File API for Node.js
global.File = class File extends Blob {
    constructor(chunks, name, opts = {}) {
        super(chunks, opts);
        this.name = name;
        this.lastModified = opts.lastModified || Date.now();
    }
};

global.Blob = class Blob {
    constructor(chunks, opts = {}) {
        this.chunks = chunks;
        this.type = opts.type || '';
        this.size = chunks.reduce((acc, chunk) => {
            if (chunk instanceof ArrayBuffer) return acc + chunk.byteLength;
            if (chunk instanceof Uint8Array) return acc + chunk.length;
            if (typeof chunk === 'string') return acc + chunk.length;
            return acc;
        }, 0);
    }
    
    async arrayBuffer() {
        // Combine all chunks into a single ArrayBuffer
        const totalSize = this.chunks.reduce((acc, chunk) => {
            if (chunk instanceof ArrayBuffer) return acc + chunk.byteLength;
            if (chunk instanceof Uint8Array) return acc + chunk.length;
            if (typeof chunk === 'string') return acc + chunk.length;
            return acc;
        }, 0);
        
        const result = new ArrayBuffer(totalSize);
        const view = new Uint8Array(result);
        let offset = 0;
        
        for (const chunk of this.chunks) {
            if (chunk instanceof ArrayBuffer) {
                view.set(new Uint8Array(chunk), offset);
                offset += chunk.byteLength;
            } else if (chunk instanceof Uint8Array) {
                view.set(chunk, offset);
                offset += chunk.length;
            } else if (typeof chunk === 'string') {
                const encoder = new TextEncoder();
                const encoded = encoder.encode(chunk);
                view.set(encoded, offset);
                offset += encoded.length;
            }
        }
        
        return result;
    }
    
    async text() {
        const buffer = await this.arrayBuffer();
        const decoder = new TextDecoder();
        return decoder.decode(buffer);
    }
};

console.log('Testing export with shared formula fix...\n');

const LuckyExcel = require('./dist/luckyexcel.umd.js');

// Read the test file
const fileBuffer = fs.readFileSync('test.xlsx');
const file = new File([fileBuffer], 'test.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
});

console.log('Step 1: Importing test.xlsx...');

LuckyExcel.transformExcelToUniver(
    file,
    (univerData) => {
        console.log('✅ Import successful\n');
        
        // Check for shared formulas
        const opSheet = Object.values(univerData.sheets).find(s => s.name === 'Operational Assumptions');
        if (opSheet && opSheet.cellData) {
            let sharedFormulaCount = 0;
            let masterCells = 0;
            for (const row in opSheet.cellData) {
                for (const col in opSheet.cellData[row]) {
                    const cell = opSheet.cellData[row][col];
                    if (cell && cell.si) {
                        sharedFormulaCount++;
                        if (cell.f) masterCells++;
                    }
                }
            }
            console.log(`Shared formula cells: ${sharedFormulaCount}`);
            console.log(`Master cells (with formula): ${masterCells}\n`);
        }
        
        console.log('Step 2: Exporting with fixed shared formula handling...');
        
        LuckyExcel.transformUniverToExcel({
            snapshot: univerData,
            fileName: 'test-fixed.xlsx',
            getBuffer: true,
            success: (buffer) => {
                if (buffer) {
                    fs.writeFileSync('test-fixed.xlsx', buffer);
                    console.log('✅ Export successful: test-fixed.xlsx created\n');
                    
                    // Now analyze the result
                    console.log('Step 3: Analyzing formulas in exported file...\n');
                    
                    const originalWB = XLSX.readFile('test.xlsx');
                    const fixedWB = XLSX.readFile('test-fixed.xlsx');
                    
                    const sheetName = 'Operational Assumptions';
                    const originalSheet = originalWB.Sheets[sheetName];
                    const fixedSheet = fixedWB.Sheets[sheetName];
                    
                    // Check specific problem cells
                    const testCells = ['O16', 'P16', 'Q16', 'R16', 'O54', 'P54', 'Q54', 'R54'];
                    let allCorrect = true;
                    
                    testCells.forEach(cell => {
                        const orig = originalSheet[cell];
                        const fixed = fixedSheet[cell];
                        
                        if (orig && orig.f) {
                            if (fixed && fixed.f && orig.f === fixed.f) {
                                console.log(`✅ ${cell}: Formula preserved correctly`);
                            } else {
                                allCorrect = false;
                                console.log(`❌ ${cell}: Formula corrupted!`);
                                console.log(`   Original: ${orig.f}`);
                                console.log(`   Fixed:    ${fixed ? fixed.f : 'Missing'}`);
                            }
                        }
                    });
                    
                    if (allCorrect) {
                        console.log('\n🎉 SUCCESS! Shared formula fix is working!');
                    } else {
                        console.log('\n⚠️  Some formulas are still corrupted.');
                    }
                } else {
                    console.error('❌ No buffer returned from export');
                }
            },
            error: (err) => {
                console.error('❌ Export failed:', err.message);
            }
        });
    },
    (error) => {
        console.error('❌ Import failed:', error.message);
        console.error('Stack:', error.stack);
    }
);