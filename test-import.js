const fs = require('fs');
const path = require('path');
const LuckyExcel = require('./dist/luckyexcel.cjs.js').LuckyExcel || require('./dist/luckyexcel.cjs.js');

// Read test.xlsx file
const filePath = path.join(__dirname, 'test.xlsx');
const fileContent = fs.readFileSync(filePath);

// Create a File-like object
const file = {
    arrayBuffer: () => Promise.resolve(fileContent),
    name: 'test.xlsx'
};

console.log('Testing import of test.xlsx...');

LuckyExcel.transformExcelToUniver(
    file,
    (result) => {
        console.log('✅ Import successful!');
        if (result && result.sheets) {
            const sheets = Object.values(result.sheets);
            console.log(`Found ${sheets.length} sheets:`);
            sheets.forEach(sheet => {
                console.log(`  - ${sheet.name}`);
            });
        }
    },
    (error) => {
        console.error('❌ Import failed:', error);
        console.error('Stack:', error.stack);
    }
);