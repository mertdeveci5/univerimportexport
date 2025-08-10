const ExcelJS = require('@zwight/exceljs');
const fs = require('fs');

async function testExcelJSTranspose() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Test');
    
    // Test 1: Set TRANSPOSE formula directly
    console.log('Test 1: Setting TRANSPOSE formula directly');
    const cell = worksheet.getCell('A1');
    cell.value = {
        formula: 'TRANSPOSE(D1:D3)',
        result: null
    };
    console.log('  Formula set:', cell.value);
    
    // Test 2: Set TRANSPOSE with @ removed
    console.log('\nTest 2: Setting TRANSPOSE with explicit @ removal');
    const cell2 = worksheet.getCell('A5');
    const formula = 'TRANSPOSE(E1:E3)'.replace(/@/g, '');
    cell2.value = {
        formula: formula,
        result: null
    };
    console.log('  Formula set:', cell2.value);
    
    // Write to file
    const buffer = await workbook.xlsx.writeBuffer();
    fs.writeFileSync('test-exceljs-output.xlsx', buffer);
    console.log('\nFile written: test-exceljs-output.xlsx');
    
    // Read back to see what was actually written
    const XLSX = require('xlsx');
    const readWorkbook = XLSX.read(buffer, {type: 'buffer'});
    const readSheet = readWorkbook.Sheets['Test'];
    
    console.log('\nReading back the written file:');
    if (readSheet['A1']) {
        console.log('A1 formula:', readSheet['A1'].f);
        if (readSheet['A1'].f && readSheet['A1'].f.includes('@')) {
            console.log('  ⚠️ Contains @ symbol!');
        }
    }
    
    if (readSheet['A5']) {
        console.log('A5 formula:', readSheet['A5'].f);
        if (readSheet['A5'].f && readSheet['A5'].f.includes('@')) {
            console.log('  ⚠️ Contains @ symbol!');
        }
    }
}

testExcelJSTranspose().catch(console.error);