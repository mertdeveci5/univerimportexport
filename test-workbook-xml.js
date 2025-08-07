const JSZip = require('@progress/jszip-esm');
const fs = require('fs');

async function testWorkbookXML() {
    const testFilePath = '/Users/mertdeveci/Desktop/Code/alphafrontend/docs/excel/test.xlsx';
    const buffer = fs.readFileSync(testFilePath);
    
    const zip = new JSZip();
    const contents = await zip.loadAsync(buffer);
    
    // Get workbook.xml content
    const workbookXML = await contents.files['xl/workbook.xml'].async('string');
    
    console.log('=== WORKBOOK.XML CONTENT ===\n');
    
    // Find all sheet references
    const sheetRegex = /<sheet[^>]*>/g;
    const sheets = workbookXML.match(sheetRegex);
    
    console.log('Sheet references found in workbook.xml:');
    sheets.forEach((sheet, index) => {
        console.log(`${index + 1}. ${sheet}`);
    });
    
    console.log('\n=== ALL WORKSHEET FILES IN ZIP ===\n');
    
    // List all worksheet files
    const worksheetFiles = [];
    for (let fileName in contents.files) {
        if (fileName.match(/xl\/worksheets\/sheet\d+\.xml$/)) {
            worksheetFiles.push(fileName);
        }
    }
    
    worksheetFiles.sort((a, b) => {
        const numA = parseInt(a.match(/sheet(\d+)\.xml/)[1]);
        const numB = parseInt(b.match(/sheet(\d+)\.xml/)[1]);
        return numA - numB;
    });
    
    console.log('Worksheet files found:');
    worksheetFiles.forEach(file => {
        console.log(`  - ${file}`);
    });
    
    console.log('\n=== COMPARISON ===\n');
    console.log(`Sheets in workbook.xml: ${sheets.length}`);
    console.log(`Worksheet files in ZIP: ${worksheetFiles.length}`);
    
    // Check which sheet files are missing from workbook.xml
    const referencedSheets = new Set();
    sheets.forEach(sheet => {
        const match = sheet.match(/r:id="(rId\d+)"/);
        if (match) {
            referencedSheets.add(match[1]);
        }
    });
    
    // Get workbook.xml.rels to map rIds to sheet files
    const relsXML = await contents.files['xl/_rels/workbook.xml.rels'].async('string');
    console.log('\n=== WORKBOOK RELS ===\n');
    
    const relRegex = /<Relationship[^>]*>/g;
    const rels = relsXML.match(relRegex);
    
    const sheetRels = {};
    rels.forEach(rel => {
        if (rel.includes('worksheets/sheet')) {
            const idMatch = rel.match(/Id="(rId\d+)"/);
            const targetMatch = rel.match(/Target="worksheets\/(sheet\d+\.xml)"/);
            if (idMatch && targetMatch) {
                sheetRels[idMatch[1]] = targetMatch[1];
                console.log(`${idMatch[1]} -> ${targetMatch[1]}`);
            }
        }
    });
    
    console.log('\n=== MISSING SHEETS ===\n');
    
    // Find which sheet files are not referenced
    const referencedFiles = new Set();
    for (let rId of referencedSheets) {
        if (sheetRels[rId]) {
            referencedFiles.add(`xl/worksheets/${sheetRels[rId]}`);
        }
    }
    
    const missingFiles = worksheetFiles.filter(file => !referencedFiles.has(file));
    console.log('Sheet files NOT referenced in workbook.xml:');
    missingFiles.forEach(file => {
        console.log(`  - ${file} (MISSING FROM WORKBOOK.XML!)`);
    });
}

testWorkbookXML().catch(console.error);