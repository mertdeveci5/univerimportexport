#!/usr/bin/env node

/**
 * Debug script to identify the real export issue
 */

const fs = require('fs');
const path = require('path');

// Mock browser APIs for Node.js
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
};

const LuckyExcel = require('./dist/luckyexcel.umd.js');

// Look for test files
const testFiles = [
    'test.xlsx',
    'DCF and LBO Model Template.xlsx',
    path.join('..', 'alphafrontend', 'docs', 'excel', 'test.xlsx')
];

const testFile = testFiles.find(f => {
    try {
        return fs.existsSync(f);
    } catch {
        return false;
    }
});

if (!testFile) {
    console.error('❌ No test file found. Please provide a test Excel file.');
    process.exit(1);
}

console.log(`📁 Using test file: ${testFile}\n`);

const fileBuffer = fs.readFileSync(testFile);
const file = new File([fileBuffer], path.basename(testFile), {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
});

console.log('Step 1: Importing to Univer format...');

LuckyExcel.transformExcelToUniver(
    file,
    (univerData) => {
        console.log('✅ Import successful\n');
        
        // Analyze the data structure
        console.log('📊 Workbook Analysis:');
        console.log(`  Total sheets: ${univerData.sheetOrder.length}`);
        
        // Check each sheet for formulas and shared formulas
        univerData.sheetOrder.forEach((sheetId, index) => {
            const sheet = univerData.sheets[sheetId];
            console.log(`\n  Sheet ${index + 1}: "${sheet.name}"`);
            
            let formulaCount = 0;
            let sharedFormulaCount = 0;
            let masterCells = 0;
            let dependentCells = 0;
            const formulaTypes = new Set();
            const problematicFormulas = [];
            
            if (sheet.cellData) {
                for (const row in sheet.cellData) {
                    for (const col in sheet.cellData[row]) {
                        const cell = sheet.cellData[row][col];
                        if (cell) {
                            if (cell.f) {
                                formulaCount++;
                                // Extract formula function name
                                const match = cell.f.match(/^=?([A-Z]+)\(/);
                                if (match) {
                                    formulaTypes.add(match[1]);
                                }
                                
                                // Check for potentially problematic formulas
                                if (cell.f.includes('TRANSPOSE') || 
                                    cell.f.includes('#') ||
                                    cell.f.includes('{') ||
                                    cell.f.includes('}')) {
                                    problematicFormulas.push({
                                        cell: `${String.fromCharCode(65 + parseInt(col))}${parseInt(row) + 1}`,
                                        formula: cell.f,
                                        hasSi: !!cell.si
                                    });
                                }
                            }
                            
                            if (cell.si) {
                                sharedFormulaCount++;
                                if (cell.f) {
                                    masterCells++;
                                } else {
                                    dependentCells++;
                                }
                            }
                        }
                    }
                }
            }
            
            console.log(`    Total formulas: ${formulaCount}`);
            console.log(`    Shared formula cells: ${sharedFormulaCount}`);
            console.log(`      - Master cells (si + f): ${masterCells}`);
            console.log(`      - Dependent cells (si only): ${dependentCells}`);
            console.log(`    Formula types used: ${Array.from(formulaTypes).join(', ')}`);
            
            if (problematicFormulas.length > 0) {
                console.log(`    ⚠️  Potentially problematic formulas: ${problematicFormulas.length}`);
                problematicFormulas.slice(0, 3).forEach(p => {
                    console.log(`      ${p.cell}: ${p.formula.substring(0, 50)}...`);
                });
            }
            
            // Check for array formulas
            if (sheet.arrayFormulas && sheet.arrayFormulas.length > 0) {
                console.log(`    Array formulas: ${sheet.arrayFormulas.length}`);
                sheet.arrayFormulas.slice(0, 2).forEach(af => {
                    console.log(`      Formula ID: ${af.formulaId}`);
                    console.log(`      Formula: ${af.formula?.substring(0, 50) || 'N/A'}...`);
                });
            }
        });
        
        console.log('\n' + '='.repeat(60));
        console.log('Step 2: Exporting back to Excel...');
        
        // Add extensive logging to the export
        const originalConsoleLog = console.log;
        const exportLogs = [];
        console.log = (...args) => {
            exportLogs.push(args.join(' '));
            originalConsoleLog.apply(console, args);
        };
        
        LuckyExcel.transformUniverToExcel({
            snapshot: univerData,
            fileName: 'debug-output.xlsx',
            getBuffer: true,
            success: (buffer) => {
                // Restore console.log
                console.log = originalConsoleLog;
                
                if (buffer) {
                    fs.writeFileSync('debug-output.xlsx', buffer);
                    console.log('✅ Export completed: debug-output.xlsx\n');
                    
                    // Check for any error logs
                    const errorLogs = exportLogs.filter(log => 
                        log.includes('ERROR') || 
                        log.includes('Error') || 
                        log.includes('error') ||
                        log.includes('Invalid') ||
                        log.includes('undefined')
                    );
                    
                    if (errorLogs.length > 0) {
                        console.log('⚠️  Errors/warnings during export:');
                        errorLogs.forEach(log => console.log(`  ${log}`));
                    }
                    
                    console.log('\n📋 Next steps:');
                    console.log('1. Open debug-output.xlsx in Excel');
                    console.log('2. Check if Excel shows corruption warning');
                    console.log('3. If corrupted, check which sheet has issues');
                    console.log('4. Compare with the analysis above');
                } else {
                    console.error('❌ Export failed: No buffer returned');
                }
            },
            error: (err) => {
                console.log = originalConsoleLog;
                console.error('❌ Export failed:', err.message);
                console.error('Stack:', err.stack);
            }
        });
    },
    (error) => {
        console.error('❌ Import failed:', error.message);
    }
);