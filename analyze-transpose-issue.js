const ExcelJS = require('@zwight/exceljs');
const XLSX = require('xlsx');
const fs = require('fs');

async function analyzeTransposeIssue() {
    console.log('=================================================');
    console.log('TRANSPOSE @ SYMBOL ANALYSIS');
    console.log('=================================================\n');
    
    // 1. Read the original Excel file to see the exact formula
    console.log('1. ORIGINAL FILE (test.xlsx)');
    console.log('-----------------------------');
    const originalBuffer = fs.readFileSync('test.xlsx');
    const originalWb = XLSX.read(originalBuffer, {type: 'buffer'});
    const originalDCF = originalWb.Sheets['DCF'];
    
    if (originalDCF && originalDCF['T83']) {
        console.log('T83 in original:');
        console.log('  Formula:', originalDCF['T83'].f);
        console.log('  Value:', originalDCF['T83'].v);
        console.log('  Type:', originalDCF['T83'].t);
    }
    
    // 2. Create test files with different TRANSPOSE approaches
    console.log('\n2. TESTING DIFFERENT TRANSPOSE APPROACHES');
    console.log('------------------------------------------');
    
    // Test A: Using ExcelJS with regular formula
    const testA = new ExcelJS.Workbook();
    const sheetA = testA.addWorksheet('Test');
    sheetA.getCell('A1').value = {
        formula: 'TRANSPOSE(D1:D3)',
        result: [1, 2, 3]
    };
    
    const bufferA = await testA.xlsx.writeBuffer();
    fs.writeFileSync('test-transpose-regular.xlsx', bufferA);
    
    // Read back what was written
    const readA = XLSX.read(bufferA, {type: 'buffer'});
    console.log('Regular formula approach:');
    console.log('  Written formula:', readA.Sheets['Test']['A1']?.f || 'none');
    
    // Test B: Using ExcelJS with array formula notation
    const testB = new ExcelJS.Workbook();
    const sheetB = testB.addWorksheet('Test');
    
    // Try setting as array formula using fillFormula
    sheetB.fillFormula('A1:C1', 'TRANSPOSE(D1:D3)', [1, 2, 3]);
    
    const bufferB = await testB.xlsx.writeBuffer();
    fs.writeFileSync('test-transpose-array.xlsx', bufferB);
    
    const readB = XLSX.read(bufferB, {type: 'buffer'});
    console.log('\nArray formula approach (fillFormula):');
    console.log('  A1 formula:', readB.Sheets['Test']['A1']?.f || 'none');
    
    // Test C: Write raw XML to understand the format
    console.log('\n3. CHECKING RAW XML STRUCTURE');
    console.log('------------------------------');
    
    // Unzip and check the worksheet XML
    const JSZip = require('@progress/jszip-esm');
    
    // Check export.xlsx raw XML
    const exportZip = new JSZip();
    await exportZip.loadAsync(fs.readFileSync('export.xlsx'));
    
    // Get DCF worksheet XML
    const dcfXml = await exportZip.file('xl/worksheets/sheet7.xml').async('string');
    
    // Find T83 formula in XML (T=column 20, row 83)
    const t83Match = dcfXml.match(/<c r="T83"[^>]*>.*?<\/c>/s);
    if (t83Match) {
        console.log('T83 in export.xlsx XML:');
        console.log(t83Match[0].substring(0, 200));
        
        // Check if it has array formula attributes
        if (t83Match[0].includes('t="array"')) {
            console.log('  ✓ Has array type');
        } else {
            console.log('  ✗ No array type');
        }
        
        // Check formula format
        const formulaMatch = t83Match[0].match(/<f[^>]*>(.*?)<\/f>/);
        if (formulaMatch) {
            console.log('  Formula in XML:', formulaMatch[1]);
        }
    }
    
    // 4. Check what modern Excel expects
    console.log('\n4. MODERN EXCEL DYNAMIC ARRAY EXPECTATIONS');
    console.log('-------------------------------------------');
    console.log('Modern Excel (365/2021) uses @ for implicit intersection.');
    console.log('TRANSPOSE should be a dynamic array formula (spills).');
    console.log('Issue: If not marked as array formula, Excel adds @ symbols.');
    
    // Test D: Create a proper dynamic array formula
    const testD = new ExcelJS.Workbook();
    const sheetD = testD.addWorksheet('Test');
    
    // Set cell with formula but mark it properly
    const cellD = sheetD.getCell('A1');
    cellD.value = {
        formula: 'TRANSPOSE(D1:D3)',
        result: null,
        shareType: 'array'  // This might help
    };
    
    const bufferD = await testD.xlsx.writeBuffer();
    fs.writeFileSync('test-transpose-dynamic.xlsx', bufferD);
    
    console.log('\nCreated test files:');
    console.log('  - test-transpose-regular.xlsx');
    console.log('  - test-transpose-array.xlsx');
    console.log('  - test-transpose-dynamic.xlsx');
    console.log('\nOpen these in Excel to see which one works correctly.');
}

analyzeTransposeIssue().catch(console.error);