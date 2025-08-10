const fs = require('fs');
const LuckyExcel = require('./dist/luckyexcel.cjs.js');

/**
 * Test what actually gets imported from test.xlsx
 */

console.log('===========================================');
console.log('IMPORT ANALYSIS - WHAT DATA DO WE GET?');
console.log('===========================================\n');

const fileBuffer = fs.readFileSync('test.xlsx');

// Mock browser FileReader
global.FileReader = class {
    readAsArrayBuffer(file) {
        this.result = fileBuffer.buffer;
        setTimeout(() => this.onload({ target: { result: this.result } }), 0);
    }
};

// Create a mock File object
const mockFile = {
    name: 'test.xlsx',
    size: fileBuffer.length,
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    arrayBuffer: () => Promise.resolve(fileBuffer.buffer)
};

// Import the file
LuckyExcel.transformExcelToUniver(
    mockFile,
    function(univerData) {
        console.log('IMPORT SUCCESSFUL\n');
        
        // 1. Check for defined names
        console.log('1. DEFINED NAMES CHECK');
        console.log('----------------------');
        console.log('univerData.namedRanges:', univerData.namedRanges);
        console.log('univerData.definedNames:', univerData.definedNames);
        console.log('univerData.names:', univerData.names);
        
        // Check all top-level properties
        console.log('\nAll top-level properties:');
        Object.keys(univerData).forEach(key => {
            if (key.toLowerCase().includes('name') || key.toLowerCase().includes('defined')) {
                console.log(`  ${key}:`, univerData[key]);
            }
        });
        
        // 2. Check for styles/borders
        console.log('\n2. STYLES/BORDERS CHECK');
        console.log('------------------------');
        console.log('univerData.styles exists:', !!univerData.styles);
        
        if (univerData.styles) {
            let borderCount = 0;
            for (let styleId in univerData.styles) {
                if (univerData.styles[styleId]?.bd) {
                    borderCount++;
                }
            }
            console.log(`Found ${borderCount} styles with borders`);
            
            // Show first few border examples
            let shown = 0;
            for (let styleId in univerData.styles) {
                if (univerData.styles[styleId]?.bd && shown < 3) {
                    console.log(`  Style ${styleId}:`, JSON.stringify(univerData.styles[styleId].bd));
                    shown++;
                }
            }
        }
        
        // 3. Check array formulas
        console.log('\n3. ARRAY FORMULAS CHECK');
        console.log('------------------------');
        
        let totalArrayFormulas = 0;
        let transposeFormulas = [];
        
        for (let sheetId in univerData.sheets) {
            const sheet = univerData.sheets[sheetId];
            console.log(`\nSheet: ${sheet.name || sheetId}`);
            
            // Check for array formulas
            if (sheet.arrayFormulas) {
                console.log(`  Array formulas: ${sheet.arrayFormulas.length}`);
                totalArrayFormulas += sheet.arrayFormulas.length;
                
                sheet.arrayFormulas.forEach((af, idx) => {
                    if (af.formula && af.formula.includes('TRANSPOSE')) {
                        transposeFormulas.push({
                            sheet: sheet.name || sheetId,
                            formula: af.formula,
                            range: af.range
                        });
                        if (idx < 2) {
                            console.log(`    [${idx}] ${af.formula} at ${JSON.stringify(af.range)}`);
                        }
                    }
                });
            }
            
            // Also check cellData for formulas
            if (sheet.cellData) {
                let formulaCount = 0;
                for (let row in sheet.cellData) {
                    for (let col in sheet.cellData[row]) {
                        const cell = sheet.cellData[row][col];
                        if (cell.f) {
                            formulaCount++;
                            if (cell.f.includes('TRANSPOSE') && formulaCount <= 2) {
                                console.log(`  Cell formula at [${row},${col}]: ${cell.f}`);
                            }
                        }
                    }
                }
                if (formulaCount > 0) {
                    console.log(`  Total formulas in cellData: ${formulaCount}`);
                }
            }
        }
        
        console.log(`\nTotal array formulas: ${totalArrayFormulas}`);
        console.log(`TRANSPOSE formulas: ${transposeFormulas.length}`);
        
        // 4. Check workbook structure
        console.log('\n4. WORKBOOK STRUCTURE');
        console.log('----------------------');
        console.log('Top-level keys:', Object.keys(univerData));
        
        // Save the data for inspection
        fs.writeFileSync('import-data.json', JSON.stringify(univerData, null, 2));
        console.log('\n✅ Full data saved to import-data.json for inspection');
    },
    function(error) {
        console.error('Import failed:', error);
    }
);