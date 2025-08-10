const fs = require('fs');
const JSZip = require('@progress/jszip-esm');

/**
 * Debug script to trace defined names through the export process
 */

async function testDefinedNamesExport() {
    console.log('=====================================');
    console.log('TESTING DEFINED NAMES EXPORT');
    console.log('=====================================\n');
    
    // Check if ExcelJS is adding defined names correctly
    const ExcelJS = require('@zwight/exceljs');
    const workbook = new ExcelJS.Workbook();
    
    console.log('1. TESTING EXCELJS DEFINED NAMES');
    console.log('---------------------------------');
    
    // Try to add a simple defined name
    console.log('Adding test defined name...');
    try {
        workbook.definedNames.add('TestName', 'Sheet1!$A$1');
        console.log('✅ Successfully added defined name');
        console.log('definedNames.model:', workbook.definedNames.model);
    } catch (error) {
        console.log('❌ Error adding defined name:', error.message);
    }
    
    // Export and check
    console.log('\n2. EXPORTING TEST WORKBOOK');
    console.log('---------------------------');
    
    // Add a worksheet (required for export)
    const worksheet = workbook.addWorksheet('Sheet1');
    worksheet.getCell('A1').value = 'Test';
    
    // Export to buffer
    const buffer = await workbook.xlsx.writeBuffer();
    console.log('Export buffer size:', buffer.byteLength);
    
    // Check the exported XML
    console.log('\n3. CHECKING EXPORTED XML');
    console.log('------------------------');
    
    const zip = new JSZip();
    await zip.loadAsync(buffer);
    
    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    
    // Look for definedNames section
    if (workbookXml.includes('<definedNames>')) {
        console.log('✅ Found <definedNames> section in XML');
        const definedNamesMatch = workbookXml.match(/<definedNames>.*?<\/definedNames>/s);
        if (definedNamesMatch) {
            console.log('Content:', definedNamesMatch[0]);
        }
    } else {
        console.log('❌ No <definedNames> section in XML');
        console.log('Workbook.xml preview:');
        console.log(workbookXml.substring(0, 500));
    }
    
    // Save test file
    fs.writeFileSync('test-defined-names-export.xlsx', buffer);
    console.log('\n✅ Test file saved as test-defined-names-export.xlsx');
    
    // Now test with our actual export
    console.log('\n4. CHECKING OUR EXPORT.XLSX');
    console.log('----------------------------');
    
    if (fs.existsSync('export.xlsx')) {
        const exportZip = new JSZip();
        const exportContent = fs.readFileSync('export.xlsx');
        await exportZip.loadAsync(exportContent);
        
        const exportWorkbookXml = await exportZip.file('xl/workbook.xml').async('string');
        
        if (exportWorkbookXml.includes('<definedNames>')) {
            console.log('✅ Found <definedNames> section in export.xlsx');
        } else {
            console.log('❌ No <definedNames> section in export.xlsx');
            
            // Check if there's a definedNames element at all
            console.log('Searching for any defined name references...');
            if (exportWorkbookXml.includes('definedName')) {
                console.log('Found "definedName" string in XML');
            } else {
                console.log('No "definedName" string found anywhere');
            }
        }
    } else {
        console.log('export.xlsx not found');
    }
    
    console.log('\n=====================================\n');
}

testDefinedNamesExport().catch(console.error);