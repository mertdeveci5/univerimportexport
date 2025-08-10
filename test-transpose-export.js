const fs = require('fs');
const LuckyExcel = require('./dist/luckyexcel.cjs.js');

async function testTransposeExport() {
    const file = fs.readFileSync('test.xlsx');
    
    return new Promise((resolve, reject) => {
        LuckyExcel.transformExcelToUniver(
            file,
            function(workbookData) {
                console.log('=== Import complete, checking DCF sheet ===');
                
                // Find DCF sheet
                const sheets = workbookData.sheets;
                let dcfSheet = null;
                for (let sheetId in sheets) {
                    if (sheets[sheetId].name === 'DCF') {
                        dcfSheet = sheets[sheetId];
                        break;
                    }
                }
                
                if (!dcfSheet) {
                    console.log('DCF sheet not found!');
                    return;
                }
                
                console.log('Found DCF sheet, checking TRANSPOSE formulas:');
                
                // Check row 83 (index 82)
                const row82 = dcfSheet.cellData && dcfSheet.cellData[82];
                if (row82) {
                    // T83 = column 19, U83 = column 20, V83 = column 21, W83 = column 22
                    console.log('\nRow 83 formulas:');
                    if (row82[19]) console.log('T83 (col 19):', JSON.stringify(row82[19]));
                    if (row82[20]) console.log('U83 (col 20):', JSON.stringify(row82[20]));
                    if (row82[21]) console.log('V83 (col 21):', JSON.stringify(row82[21]));
                    if (row82[22]) console.log('W83 (col 22):', JSON.stringify(row82[22]));
                    if (row82[25]) console.log('Z83 (col 25):', JSON.stringify(row82[25]));
                }
                
                // Now export back to Excel
                console.log('\n=== Exporting back to Excel ===');
                LuckyExcel.transformUniverToExcel({
                    snapshot: workbookData,
                    fileName: 'test-transpose-export.xlsx',
                    success: function(buffer) {
                        fs.writeFileSync('test-transpose-export.xlsx', buffer);
                        console.log('Export complete! File saved as test-transpose-export.xlsx');
                        
                        // Now read the exported file to check formulas
                        const XLSX = require('xlsx');
                        const exportedWb = XLSX.read(buffer, {type: 'buffer'});
                        const exportedDcf = exportedWb.Sheets['DCF'];
                        
                        if (exportedDcf) {
                            console.log('\n=== Checking exported TRANSPOSE formulas ===');
                            const checkCells = ['T83', 'W83', 'Z83', 'T84', 'T100', 'W100', 'Z100'];
                            for (let cell of checkCells) {
                                if (exportedDcf[cell] && exportedDcf[cell].f) {
                                    console.log(`${cell}: ${exportedDcf[cell].f}`);
                                    if (exportedDcf[cell].f.includes('@')) {
                                        console.log('  ⚠️ Contains @ symbol!');
                                    }
                                }
                            }
                        }
                        
                        resolve();
                    },
                    error: reject
                });
            },
            reject
        );
    });
}

testTransposeExport().catch(console.error);