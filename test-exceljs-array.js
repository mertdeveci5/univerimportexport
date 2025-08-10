const ExcelJS = require('@zwight/exceljs');
const XLSX = require('xlsx');
const fs = require('fs');

async function testExcelJSArrayFormula() {
    console.log('=================================================');
    console.log('TESTING EXCELJS ARRAY FORMULA CAPABILITY');
    console.log('=================================================\n');
    
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Test');
    
    // Test 1: Try using fillFormula (meant for array formulas)
    console.log('Test 1: Using fillFormula for TRANSPOSE');
    console.log('-----------------------------------------');
    
    try {
        // fillFormula is supposed to create array formulas
        sheet.fillFormula('A1:C1', 'TRANSPOSE(D1:D3)', [1, 2, 3]);
        console.log('✓ fillFormula executed');
    } catch (e) {
        console.log('✗ fillFormula failed:', e.message);
    }
    
    // Test 2: Try setting formula with shareType
    console.log('\nTest 2: Setting formula with shareType');
    console.log('---------------------------------------');
    
    const cell = sheet.getCell('A5');
    cell.value = {
        formula: 'TRANSPOSE(D5:D7)',
        result: 10,
        shareType: 'array'
    };
    console.log('✓ Set formula with shareType: array');
    
    // Test 3: Try using the model directly
    console.log('\nTest 3: Using cell model directly');
    console.log('----------------------------------');
    
    const cellModel = sheet.getCell('A10').model;
    cellModel.formula = 'TRANSPOSE(D10:D12)';
    cellModel.formulaType = 'array';
    cellModel.arrayFormula = 'A10:C10';
    console.log('✓ Set model with formulaType: array');
    
    // Write the file
    const buffer = await workbook.xlsx.writeBuffer();
    fs.writeFileSync('test-exceljs-array.xlsx', buffer);
    console.log('\n✓ File written: test-exceljs-array.xlsx');
    
    // Read back with XLSX to see what was actually written
    console.log('\n4. CHECKING WHAT WAS WRITTEN');
    console.log('-----------------------------');
    
    const readWb = XLSX.read(buffer, {type: 'buffer'});
    console.log('Available sheets:', readWb.SheetNames);
    
    const readSheet = readWb.Sheets['Test'];
    
    if (!readSheet) {
        console.log('❌ Sheet "Test" not found in exported file');
        console.log('Sheet keys:', Object.keys(readWb.Sheets));
        return;
    }
    
    // Check each test
    if (readSheet['A1']) {
        console.log('A1 (fillFormula):');
        console.log('  Formula:', readSheet['A1'].f || 'none');
        
        // Check raw XML for this cell
        const JSZip = require('@progress/jszip-esm');
        const zip = new JSZip();
        await zip.loadAsync(buffer);
        const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
        
        const a1Match = sheetXml.match(/<c r="A1"[^>]*>.*?<\/c>/s);
        if (a1Match) {
            console.log('  XML:', a1Match[0]);
            if (a1Match[0].includes('t="array"')) {
                console.log('  ✓ Has t="array" attribute!');
            } else {
                console.log('  ✗ No t="array" attribute');
            }
        }
    }
    
    if (readSheet['A5']) {
        console.log('\nA5 (shareType):');
        console.log('  Formula:', readSheet['A5'].f || 'none');
    }
    
    if (readSheet['A10']) {
        console.log('\nA10 (model.formulaType):');
        console.log('  Formula:', readSheet['A10'].f || 'none');
    }
    
    console.log('\n=================================================');
    console.log('CONCLUSION');
    console.log('=================================================');
    console.log('ExcelJS does not properly support array formulas with t="array"');
    console.log('This is why TRANSPOSE formulas get @ symbols in Excel');
    console.log('\nPossible solutions:');
    console.log('1. Patch ExcelJS to write t="array" and ref attributes');
    console.log('2. Post-process the XML after ExcelJS writes it');
    console.log('3. Use a different library that supports array formulas');
}

testExcelJSArrayFormula().catch(console.error);