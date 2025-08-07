const ExcelJS = require('@zwight/exceljs');
const fs = require('fs');
const path = require('path');

async function testExcelJS() {
    const testFilePath = '/Users/mertdeveci/Desktop/Code/alphafrontend/docs/excel/test.xlsx';
    
    console.log('Testing with ExcelJS directly...\n');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(testFilePath);
    
    console.log('Workbook properties:');
    console.log('- Title:', workbook.title);
    console.log('- Company:', workbook.company);
    console.log('- Created:', workbook.created);
    console.log('- Modified:', workbook.modified);
    console.log('');
    
    console.log('Total worksheets found by ExcelJS:', workbook.worksheets.length);
    console.log('');
    
    console.log('All worksheets:');
    workbook.eachSheet((worksheet, sheetId) => {
        console.log(`Sheet ${sheetId}: "${worksheet.name}"`);
        console.log(`  - State: ${worksheet.state}`);
        console.log(`  - Row count: ${worksheet.rowCount}`);
        console.log(`  - Column count: ${worksheet.columnCount}`);
        console.log(`  - Actual row count: ${worksheet.actualRowCount}`);
        console.log(`  - Actual column count: ${worksheet.actualColumnCount}`);
        console.log(`  - Has data: ${worksheet.actualRowCount > 0 || worksheet.actualColumnCount > 0}`);
        
        // Count actual cells with data
        let cellCount = 0;
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                cellCount++;
            });
        });
        console.log(`  - Cells with data: ${cellCount}`);
        console.log('');
    });
    
    // Check if there are any hidden sheets
    const hiddenSheets = workbook.worksheets.filter(ws => ws.state === 'hidden');
    console.log('Hidden sheets:', hiddenSheets.length);
    
    // Check for sheets with special characters
    const sheetsWithSpecialChars = workbook.worksheets.filter(ws => ws.name.includes('>>>'));
    console.log('Sheets with ">>>":', sheetsWithSpecialChars.map(ws => ws.name));
}

testExcelJS().catch(console.error);