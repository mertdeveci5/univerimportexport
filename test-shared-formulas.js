#!/usr/bin/env node

/**
 * Test to understand shared formula handling
 */

const fs = require('fs');
const LuckyExcel = require('./dist/luckyexcel.umd.js');

console.log('Testing shared formula handling...\n');

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
const file = new NodeFile(fileBuffer, 'test.xlsx');

console.log('Importing test.xlsx to analyze shared formulas...');

LuckyExcel.transformExcelToUniver(
    file,
    (univerData) => {
        console.log('✅ Import successful\n');
        
        // Look for shared formulas in Operational Assumptions sheet
        const targetSheet = Object.values(univerData.sheets).find(s => 
            s.name === 'Operational Assumptions'
        );
        
        if (!targetSheet) {
            console.log('❌ Could not find Operational Assumptions sheet');
            return;
        }
        
        console.log(`Analyzing sheet: ${targetSheet.name}`);
        console.log('Looking for cells with shared formulas (si field)...\n');
        
        // Check row 16 cells (O16-V16) which have the corruption issue
        const targetRows = [14, 16, 54, 60, 74, 86]; // 0-based (subtract 2 for Excel row numbers)
        const targetCols = [14, 15, 16, 17, 18, 19, 20, 21]; // O-V columns (0-based)
        
        let sharedFormulaMap = {};
        let masterCells = {};
        
        // First pass: collect all cells with si field
        if (targetSheet.cellData) {
            for (const rowStr in targetSheet.cellData) {
                const row = parseInt(rowStr);
                if (!targetRows.includes(row)) continue;
                
                const rowData = targetSheet.cellData[rowStr];
                for (const colStr in rowData) {
                    const col = parseInt(colStr);
                    if (!targetCols.includes(col)) continue;
                    
                    const cell = rowData[colStr];
                    if (cell) {
                        const cellAddr = String.fromCharCode(65 + col) + (row + 1);
                        
                        if (cell.si) {
                            if (!sharedFormulaMap[cell.si]) {
                                sharedFormulaMap[cell.si] = [];
                            }
                            sharedFormulaMap[cell.si].push({
                                addr: cellAddr,
                                row: row,
                                col: col,
                                hasFormula: !!cell.f,
                                formula: cell.f,
                                value: cell.v
                            });
                            
                            if (cell.f) {
                                masterCells[cell.si] = cellAddr;
                            }
                        } else if (cell.f) {
                            console.log(`Regular formula at ${cellAddr}: ${cell.f}`);
                        }
                    }
                }
            }
        }
        
        // Report findings
        console.log('\n📊 Shared Formula Analysis:');
        console.log(`Found ${Object.keys(sharedFormulaMap).length} shared formula groups\n`);
        
        for (const [formulaId, cells] of Object.entries(sharedFormulaMap)) {
            console.log(`\nShared Formula ID: ${formulaId}`);
            console.log(`  Master Cell: ${masterCells[formulaId] || 'Unknown'}`);
            console.log(`  Cell Count: ${cells.length}`);
            console.log('  Cells:');
            cells.forEach(c => {
                const type = c.hasFormula ? '(master)' : '(dependent)';
                console.log(`    ${c.addr} ${type}: formula=${c.formula || 'none'}, value=${c.value}`);
            });
        }
        
        // Check arrayFormulas if present
        if (targetSheet.arrayFormulas) {
            console.log('\n📐 Array Formulas:');
            console.log(`Found ${targetSheet.arrayFormulas.length} array formulas`);
            targetSheet.arrayFormulas.forEach(af => {
                console.log(`  Formula ID: ${af.formulaId}`);
                console.log(`    Range: ${JSON.stringify(af.range)}`);
                console.log(`    Formula: ${af.formula}`);
                console.log(`    Master: [${af.masterRow}, ${af.masterCol}]`);
            });
        }
        
        // Save for inspection
        fs.writeFileSync('shared-formula-analysis.json', JSON.stringify({
            sharedFormulaMap,
            masterCells,
            arrayFormulas: targetSheet.arrayFormulas || []
        }, null, 2));
        
        console.log('\n💾 Analysis saved to shared-formula-analysis.json');
    },
    (error) => {
        console.error('❌ Import failed:', error.message);
    }
);