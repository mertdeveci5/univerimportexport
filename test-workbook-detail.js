const JSZip = require('@progress/jszip-esm');
const fs = require('fs');

async function testWorkbookDetail() {
    const testFilePath = '/Users/mertdeveci/Desktop/Code/alphafrontend/docs/excel/test.xlsx';
    const buffer = fs.readFileSync(testFilePath);
    
    const zip = new JSZip();
    const contents = await zip.loadAsync(buffer);
    
    // Get workbook.xml content
    const workbookXML = await contents.files['xl/workbook.xml'].async('string');
    
    console.log('=== RAW WORKBOOK.XML SHEETS SECTION ===\n');
    
    // Extract just the sheets section
    const sheetsStart = workbookXML.indexOf('<sheets>');
    const sheetsEnd = workbookXML.indexOf('</sheets>') + '</sheets>'.length;
    const sheetsSection = workbookXML.substring(sheetsStart, sheetsEnd);
    
    // Pretty print it
    console.log(sheetsSection.replace(/></g, '>\n<'));
    
    console.log('\n=== PARSED SHEETS ===\n');
    
    // Parse each sheet element properly
    const sheetRegex = /<sheet ([^>]+)\/>/g;
    let match;
    let sheetCount = 0;
    const parsedSheets = [];
    
    while ((match = sheetRegex.exec(sheetsSection)) !== null) {
        sheetCount++;
        const attrs = match[1];
        
        // Parse attributes
        const nameMatch = attrs.match(/name="([^"]+)"/);
        const sheetIdMatch = attrs.match(/sheetId="([^"]+)"/);
        const ridMatch = attrs.match(/r:id="([^"]+)"/);
        
        const sheetInfo = {
            name: nameMatch ? nameMatch[1] : 'UNKNOWN',
            sheetId: sheetIdMatch ? sheetIdMatch[1] : 'UNKNOWN',
            rId: ridMatch ? ridMatch[1] : 'UNKNOWN'
        };
        
        parsedSheets.push(sheetInfo);
        console.log(`${sheetCount}. Name: "${sheetInfo.name}", SheetId: ${sheetInfo.sheetId}, rId: ${sheetInfo.rId}`);
    }
    
    console.log('\n=== WORKBOOK.XML.RELS MAPPING ===\n');
    
    // Get the rels file
    const relsXML = await contents.files['xl/_rels/workbook.xml.rels'].async('string');
    const relRegex = /<Relationship ([^>]+)\/>/g;
    const ridToFile = {};
    
    while ((match = relRegex.exec(relsXML)) !== null) {
        const attrs = match[1];
        const idMatch = attrs.match(/Id="([^"]+)"/);
        const targetMatch = attrs.match(/Target="([^"]+)"/);
        
        if (idMatch && targetMatch && targetMatch[1].includes('worksheets/sheet')) {
            ridToFile[idMatch[1]] = targetMatch[1];
        }
    }
    
    console.log('rId to File mapping:');
    for (let rId in ridToFile) {
        console.log(`  ${rId} -> ${ridToFile[rId]}`);
    }
    
    console.log('\n=== FINAL ANALYSIS ===\n');
    
    // Map sheets to files
    console.log('Sheets and their files:');
    parsedSheets.forEach((sheet, index) => {
        const file = ridToFile[sheet.rId];
        console.log(`${index + 1}. "${sheet.name}" (ID: ${sheet.sheetId}) -> ${file || 'NO FILE FOUND!'}`);
    });
    
    // Find orphaned files
    console.log('\n=== ORPHANED SHEET FILES ===\n');
    const referencedFiles = new Set(Object.values(ridToFile));
    const allSheetFiles = [];
    
    for (let fileName in contents.files) {
        if (fileName.match(/xl\/worksheets\/sheet\d+\.xml$/)) {
            const shortName = fileName.replace('xl/', '');
            allSheetFiles.push(shortName);
            
            if (!referencedFiles.has(shortName)) {
                console.log(`ORPHANED: ${fileName}`);
                
                // Check if it's empty
                const content = await contents.files[fileName].async('string');
                const hasData = content.includes('<c ') || content.includes('<v>');
                console.log(`  - Has data: ${hasData}`);
                console.log(`  - File size: ${content.length} bytes`);
            }
        }
    }
}

testWorkbookDetail().catch(console.error);