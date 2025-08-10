const ExcelJS = require('@zwight/exceljs');
const fs = require('fs');

async function checkExportedFile() {
    console.log('=================================================');
    console.log('CHECKING EXPORTED FILE WITH EXCELJS');
    console.log('=================================================\n');
    
    // Read the exported file with ExcelJS
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('export.xlsx');
    
    // 1. Check for defined names
    console.log('1. DEFINED NAMES IN EXCELJS');
    console.log('----------------------------');
    if (workbook.definedNames) {
        console.log('Defined names count:', Object.keys(workbook.definedNames).length);
        Object.keys(workbook.definedNames).slice(0, 5).forEach(name => {
            console.log(`  - ${name}: ${workbook.definedNames[name]}`);
        });
    } else {
        console.log('No defined names found');
    }
    
    // 2. Check DCF sheet for TRANSPOSE formulas
    console.log('\n2. TRANSPOSE FORMULAS IN DCF (via ExcelJS)');
    console.log('-------------------------------------------');
    const dcfSheet = workbook.getWorksheet('DCF');
    
    if (dcfSheet) {
        const transposeCells = [
            {row: 83, col: 20}, // T83 (T=20)
            {row: 83, col: 23}, // W83 (W=23)
            {row: 83, col: 26}, // Z83 (Z=26)
            {row: 84, col: 20}, // T84
            {row: 100, col: 20}, // T100
        ];
        
        transposeCells.forEach(({row, col}) => {
            const cell = dcfSheet.getCell(row, col);
            const colLetter = String.fromCharCode(64 + col);
            
            if (cell.formula) {
                console.log(`${colLetter}${row}:`);
                console.log(`  Formula: ${cell.formula}`);
                console.log(`  Type: ${cell.type}`);
                console.log(`  Value: ${cell.value}`);
                
                // Check the actual formula object
                if (typeof cell.value === 'object' && cell.value.formula) {
                    console.log(`  Value.formula: ${cell.value.formula}`);
                }
            }
        });
        
        // 3. Check borders
        console.log('\n3. BORDER STYLES (via ExcelJS)');
        console.log('-------------------------------');
        let borderCount = 0;
        let sampleCount = 0;
        
        for (let row = 1; row <= 20; row++) {
            for (let col = 1; col <= 10; col++) {
                const cell = dcfSheet.getCell(row, col);
                sampleCount++;
                
                if (cell.border && Object.keys(cell.border).length > 0) {
                    borderCount++;
                    if (borderCount <= 3) {
                        const colLetter = String.fromCharCode(64 + col);
                        console.log(`  ${colLetter}${row} has borders:`, JSON.stringify(cell.border));
                    }
                }
            }
        }
        
        console.log(`Cells with borders: ${borderCount}/${sampleCount}`);
    }
    
    // 4. Write a test file with ExcelJS to see how it handles TRANSPOSE
    console.log('\n4. TESTING EXCELJS TRANSPOSE HANDLING');
    console.log('--------------------------------------');
    const testWb = new ExcelJS.Workbook();
    const testSheet = testWb.addWorksheet('Test');
    
    // Test different ways to set TRANSPOSE
    testSheet.getCell('A1').value = {
        formula: 'TRANSPOSE(D1:D3)',
        result: null
    };
    
    testSheet.getCell('A5').value = {
        formula: '=TRANSPOSE(E1:E3)',
        result: null
    };
    
    // Write and read back
    const buffer = await testWb.xlsx.writeBuffer();
    fs.writeFileSync('exceljs-transpose-test.xlsx', buffer);
    
    // Read back with XLSX to see what was written
    const XLSX = require('xlsx');
    const readBack = XLSX.read(buffer, {type: 'buffer'});
    const testSheetRead = readBack.Sheets['Test'];
    
    console.log('After ExcelJS writes TRANSPOSE:');
    if (testSheetRead['A1']) {
        console.log(`  A1 formula: ${testSheetRead['A1'].f}`);
    }
    if (testSheetRead['A5']) {
        console.log(`  A5 formula: ${testSheetRead['A5'].f}`);
    }
}

checkExportedFile().catch(console.error);