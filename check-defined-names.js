const fs = require('fs');
const JSZip = require('@progress/jszip-esm');

async function checkDefinedNames() {
    console.log('=================================================');
    console.log('DEFINED NAMES INVESTIGATION');
    console.log('=================================================\n');
    
    // 1. Check original file for defined names in XML
    console.log('1. ORIGINAL FILE (test.xlsx) - Defined Names');
    console.log('----------------------------------------------');
    
    const originalZip = new JSZip();
    await originalZip.loadAsync(fs.readFileSync('test.xlsx'));
    
    // Check workbook.xml for defined names
    const workbookXml = await originalZip.file('xl/workbook.xml').async('string');
    
    // Extract defined names
    const definedNamesMatch = workbookXml.match(/<definedNames>.*?<\/definedNames>/s);
    if (definedNamesMatch) {
        console.log('Found definedNames section in workbook.xml');
        
        // Count individual names
        const nameMatches = definedNamesMatch[0].match(/<definedName[^>]*>/g);
        if (nameMatches) {
            console.log(`Total defined names: ${nameMatches.length}`);
            
            // Show first few
            nameMatches.slice(0, 5).forEach(nameTag => {
                const nameMatch = nameTag.match(/name="([^"]+)"/);
                const refMatch = definedNamesMatch[0].match(new RegExp(nameTag + '([^<]*)<'));
                if (nameMatch && refMatch) {
                    console.log(`  - ${nameMatch[1]}: ${refMatch[1]}`);
                }
            });
        }
    }
    
    // 2. Check exported file
    console.log('\n2. EXPORTED FILE (export.xlsx) - Defined Names');
    console.log('------------------------------------------------');
    
    const exportZip = new JSZip();
    await exportZip.loadAsync(fs.readFileSync('export.xlsx'));
    
    const exportWorkbookXml = await exportZip.file('xl/workbook.xml').async('string');
    
    const exportDefinedNamesMatch = exportWorkbookXml.match(/<definedNames>.*?<\/definedNames>/s);
    if (exportDefinedNamesMatch) {
        console.log('Found definedNames section in export workbook.xml');
        
        const exportNameMatches = exportDefinedNamesMatch[0].match(/<definedName[^>]*>/g);
        if (exportNameMatches) {
            console.log(`Total defined names: ${exportNameMatches.length}`);
            
            exportNameMatches.forEach(nameTag => {
                const nameMatch = nameTag.match(/name="([^"]+)"/);
                if (nameMatch) {
                    console.log(`  - ${nameMatch[1]}`);
                }
            });
        }
    } else {
        console.log('No definedNames section found in export workbook.xml');
    }
    
    // 3. Check import process
    console.log('\n3. CHECK IMPORT CODE FOR DEFINED NAMES');
    console.log('----------------------------------------');
    console.log('Searching for defined name handling in import code...');
    
    // Read LuckyFile.ts to see if it processes defined names
    const luckyFileContent = fs.readFileSync('src/ToLuckySheet/LuckyFile.ts', 'utf8');
    if (luckyFileContent.includes('definedName')) {
        console.log('✓ LuckyFile.ts mentions definedName');
    } else {
        console.log('✗ LuckyFile.ts does NOT mention definedName');
    }
    
    // Check ReadXml.ts
    const readXmlContent = fs.readFileSync('src/ToLuckySheet/ReadXml.ts', 'utf8');
    if (readXmlContent.includes('definedName')) {
        console.log('✓ ReadXml.ts mentions definedName');
    } else {
        console.log('✗ ReadXml.ts does NOT mention definedName');
    }
    
    // 4. Check borders in original vs export
    console.log('\n4. BORDER STYLES IN XML');
    console.log('------------------------');
    
    // Check styles.xml in original
    const originalStylesXml = await originalZip.file('xl/styles.xml').async('string');
    const originalBordersMatch = originalStylesXml.match(/<borders[^>]*count="(\d+)"/);
    if (originalBordersMatch) {
        console.log(`Original file: ${originalBordersMatch[1]} border definitions`);
    }
    
    // Check styles.xml in export
    const exportStylesXml = await exportZip.file('xl/styles.xml').async('string');
    const exportBordersMatch = exportStylesXml.match(/<borders[^>]*count="(\d+)"/);
    if (exportBordersMatch) {
        console.log(`Exported file: ${exportBordersMatch[1]} border definitions`);
    }
}

checkDefinedNames().catch(console.error);