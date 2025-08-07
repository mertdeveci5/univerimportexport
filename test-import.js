const { LuckyExcel } = require('./dist/main.js');
const fs = require('fs');
const path = require('path');

async function testImport() {
    const testFilePath = '/Users/mertdeveci/Desktop/Code/alphafrontend/docs/excel/test.xlsx';
    
    if (!fs.existsSync(testFilePath)) {
        console.error('Test file not found:', testFilePath);
        return;
    }
    
    const buffer = fs.readFileSync(testFilePath);
    const file = new File([buffer], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    return new Promise((resolve, reject) => {
        LuckyExcel.transformExcelToUniver(
            file,
            (workbookData) => {
                console.log('=== IMPORT COMPLETE ===');
                console.log('Total sheets imported:', Object.keys(workbookData.sheets).length);
                console.log('\nSheet details:');
                Object.entries(workbookData.sheets).forEach(([id, sheet], index) => {
                    const cellCount = sheet.cellData ? Object.keys(sheet.cellData).reduce((acc, row) => {
                        return acc + (sheet.cellData[row] ? Object.keys(sheet.cellData[row]).length : 0);
                    }, 0) : 0;
                    console.log(`${index + 1}. ${sheet.name} (ID: ${id})`);
                    console.log(`   - Cells: ${cellCount}`);
                    console.log(`   - Dimensions: ${sheet.rowCount}x${sheet.columnCount}`);
                });
                resolve(workbookData);
            },
            (error) => {
                console.error('Import error:', error);
                reject(error);
            }
        );
    });
}

// Run test
testImport().catch(console.error);