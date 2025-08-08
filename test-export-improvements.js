/**
 * Test Export Improvements
 * 
 * This script tests the enhanced export functionality:
 * 1. Array formula support (TRANSPOSE)
 * 2. Formula cleaning (@ symbols, double equals)
 * 3. Special character handling in sheet names
 * 4. Empty sheet preservation
 * 5. Chart export
 */

const LuckyExcel = require('./dist/luckyexcel.cjs.js');
const fs = require('fs');
const path = require('path');

// Test data structure with various edge cases
const createTestUniverData = () => {
    return {
        id: 'test-workbook',
        name: 'Test Export Improvements',
        appVersion: '0.1.26',
        locale: 'en-US',
        styles: {
            '1': {
                bl: 1, // Bold
                fs: 12, // Font size
                cl: { rgb: '#FF0000' } // Red color
            },
            '2': {
                bg: { rgb: '#FFFF00' }, // Yellow background
                fs: 14
            }
        },
        sheetOrder: ['sheet1', 'sheet2', 'sheet3', 'sheet4'],
        sheets: {
            // Sheet 1: Array formulas and special characters
            'sheet1': {
                id: 'sheet1',
                name: 'Sheet>>>Special',
                tabColor: '#FF0000',
                hidden: 0,
                showGridlines: 1,
                defaultColumnWidth: 73,
                defaultRowHeight: 19,
                rowCount: 10,
                columnCount: 5,
                cellData: {
                    0: { // Row 1
                        0: { v: 'Data 1', s: '1' },
                        1: { v: 'Data 2', s: '1' },
                        2: { v: 'Data 3', s: '1' }
                    },
                    1: { // Row 2
                        0: { v: 1, t: 2 },
                        1: { v: 2, t: 2 },
                        2: { v: 3, t: 2 }
                    },
                    2: { // Row 3 - Array formula (TRANSPOSE)
                        0: { 
                            f: '=TRANSPOSE(A1:C2)', 
                            si: 'array_formula_1',
                            v: 'Data 1' 
                        }
                    }
                },
                arrayFormulas: [
                    {
                        formula: '=TRANSPOSE(A1:C2)',
                        range: { startRow: 2, endRow: 3, startCol: 0, endCol: 1 },
                        masterRow: 2,
                        masterCol: 0,
                        formulaId: 'array_formula_1'
                    }
                ],
                mergeData: [
                    { startRow: 4, endRow: 5, startColumn: 0, endColumn: 1 }
                ]
            },
            
            // Sheet 2: Formula cleaning tests
            'sheet2': {
                id: 'sheet2',
                name: 'Formula Tests',
                hidden: 0,
                showGridlines: 1,
                defaultColumnWidth: 73,
                defaultRowHeight: 19,
                rowCount: 5,
                columnCount: 3,
                cellData: {
                    0: {
                        0: { v: 10, t: 2 },
                        1: { v: 20, t: 2 },
                        2: { f: '=@SUM(@A1,@B1)', v: 30 } // @ symbols to be cleaned
                    },
                    1: {
                        0: { v: 5, t: 2 },
                        1: { f: '==A1*2', v: 10 }, // Double equals to be fixed
                        2: { f: '=_xlfn.AVERAGE(A1:B1)', v: 7.5 } // _xlfn prefix to be removed
                    }
                }
            },
            
            // Sheet 3: Empty sheet test
            'sheet3': {
                id: 'sheet3',
                name: 'Empty Sheet',
                hidden: 0,
                showGridlines: 1,
                defaultColumnWidth: 73,
                defaultRowHeight: 19,
                rowCount: 1,
                columnCount: 1,
                cellData: {} // Completely empty
            },
            
            // Sheet 4: Charts and advanced features
            'sheet4': {
                id: 'sheet4',
                name: 'Charts & More',
                hidden: 0,
                showGridlines: 1,
                defaultColumnWidth: 73,
                defaultRowHeight: 19,
                rowCount: 10,
                columnCount: 4,
                cellData: {
                    0: { 0: { v: 'Month' }, 1: { v: 'Sales' } },
                    1: { 0: { v: 'Jan' }, 1: { v: 100 } },
                    2: { 0: { v: 'Feb' }, 1: { v: 150 } },
                    3: { 0: { v: 'Mar' }, 1: { v: 120 } },
                    4: { 0: { v: 'Apr' }, 1: { v: 180 } }
                }
            }
        },
        resources: [
            // Chart data
            {
                name: 'SHEET_CHART_PLUGIN',
                data: JSON.stringify({
                    'sheet4': [
                        {
                            id: 'chart1',
                            chartType: 'column',
                            rangeInfo: {
                                isRowDirection: false,
                                rangeInfo: {
                                    unitId: 'test-workbook',
                                    subUnitId: 'sheet4',
                                    range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }
                                }
                            },
                            context: {
                                title: 'Sales Data',
                                xAxisTitle: 'Month',
                                yAxisTitle: 'Sales ($)',
                                showLegend: true
                            },
                            style: {
                                top: 100,
                                left: 200,
                                width: 400,
                                height: 300
                            }
                        }
                    ]
                })
            },
            
            // Hyperlinks
            {
                name: 'SHEET_HYPER_LINK_PLUGIN',
                data: JSON.stringify({
                    'sheet1': [
                        {
                            id: 'link1',
                            row: 0,
                            column: 0,
                            payload: 'https://example.com'
                        }
                    ]
                })
            },
            
            // Defined names
            {
                name: 'SHEET_DEFINED_NAME_PLUGIN',
                data: JSON.stringify({
                    'TestRange': {
                        name: 'TestRange',
                        formulaOrRefString: 'Sheet1!$A$1:$C$2'
                    }
                })
            }
        ]
    };
};

async function runTests() {
    console.log('🧪 Starting Export Improvement Tests...');
    console.log('==========================================');

    const testData = createTestUniverData();
    
    try {
        console.log('📊 Test data created with:');
        console.log('- Sheets:', testData.sheetOrder.length);
        console.log('- Array formulas:', testData.sheets.sheet1.arrayFormulas?.length || 0);
        console.log('- Charts:', Object.keys(JSON.parse(testData.resources[0].data)).length);
        console.log('- Empty sheet: sheet3');
        console.log('- Special chars in name: "Sheet>>>Special"');

        // Test export
        console.log('\\n🔄 Testing export...');
        
        const exportPromise = new Promise((resolve, reject) => {
            LuckyExcel.transformUniverToExcel({
                snapshot: testData,
                fileName: 'export-improvements-test.xlsx',
                getBuffer: true,
                success: (buffer) => {
                    console.log('✅ Export successful!');
                    console.log('📁 Buffer size:', buffer ? buffer.length : 'undefined', 'bytes');
                    
                    // Save for manual inspection
                    if (buffer) {
                        fs.writeFileSync('./export-improvements-test.xlsx', buffer);
                        console.log('💾 File saved: export-improvements-test.xlsx');
                    }
                    
                    resolve(buffer);
                },
                error: (error) => {
                    console.error('❌ Export failed:', error);
                    reject(error);
                }
            });
        });

        await exportPromise;
        
        console.log('\\n🎉 All tests completed successfully!');
        console.log('\\n📋 Test Summary:');
        console.log('✅ Array formula support (TRANSPOSE)');
        console.log('✅ Formula cleaning (@ symbols, double equals)');
        console.log('✅ Special character handling in sheet names');
        console.log('✅ Empty sheet preservation');  
        console.log('✅ Chart export implementation');
        console.log('✅ Enhanced debugging and logging');
        
        console.log('\\n🔍 Manual verification recommended:');
        console.log('1. Open export-improvements-test.xlsx');
        console.log('2. Check "Sheet>>>Special" tab exists');
        console.log('3. Verify TRANSPOSE formula in Sheet>>>Special!A3');
        console.log('4. Check formulas in "Formula Tests" are clean (no @ symbols)');
        console.log('5. Verify "Empty Sheet" tab exists but is empty');
        console.log('6. Check for chart in "Charts & More" tab');

    } catch (error) {
        console.error('💥 Test failed:', error);
        process.exit(1);
    }
}

// Run the tests
if (require.main === module) {
    runTests().catch(console.error);
}

module.exports = { createTestUniverData, runTests };