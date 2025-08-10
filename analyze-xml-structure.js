const fs = require('fs');
const JSZip = require('@progress/jszip-esm');

async function analyzeXmlStructure() {
    console.log('=================================================');
    console.log('XML STRUCTURE ANALYSIS FOR TRANSPOSE');
    console.log('=================================================\n');
    
    // 1. Check original Excel file's TRANSPOSE XML
    console.log('1. ORIGINAL test.xlsx - T83 CELL XML');
    console.log('--------------------------------------');
    
    const originalZip = new JSZip();
    await originalZip.loadAsync(fs.readFileSync('test.xlsx'));
    
    // Get DCF worksheet (sheet7.xml based on sheet order)
    const worksheetFiles = originalZip.file(/xl\/worksheets\/sheet\d+\.xml/);
    console.log('Available sheets:', worksheetFiles.map(f => f.name));
    
    // Find DCF sheet
    const workbookRels = await originalZip.file('xl/_rels/workbook.xml.rels').async('string');
    const sheetRels = workbookRels.match(/Relationship[^>]*sheet\d+\.xml/g);
    
    // Get sheet7.xml (DCF is typically sheet 7)
    let dcfSheetXml = null;
    for (let i = 1; i <= 13; i++) {
        try {
            const sheetXml = await originalZip.file(`xl/worksheets/sheet${i}.xml`).async('string');
            if (sheetXml.includes('TRANSPOSE')) {
                console.log(`Found TRANSPOSE in sheet${i}.xml`);
                dcfSheetXml = sheetXml;
                
                // Extract T83 cell
                const t83Match = sheetXml.match(/<c r="T83"[^>]*>.*?<\/c>/s);
                if (t83Match) {
                    console.log('\nT83 cell XML in original:');
                    console.log(t83Match[0]);
                    
                    // Check attributes
                    const attributes = t83Match[0].match(/<c[^>]*>/)[0];
                    console.log('\nCell attributes:', attributes);
                    
                    // Check formula
                    const formulaMatch = t83Match[0].match(/<f[^>]*>(.*?)<\/f>/);
                    if (formulaMatch) {
                        console.log('Formula element:', formulaMatch[0]);
                        
                        // Check for array formula attributes
                        if (formulaMatch[0].includes('t="array"')) {
                            console.log('  ✓ Has t="array" attribute');
                        } else {
                            console.log('  ✗ No t="array" attribute');
                        }
                        
                        if (formulaMatch[0].includes('ref=')) {
                            const refMatch = formulaMatch[0].match(/ref="([^"]+)"/);
                            console.log('  ✓ Has ref attribute:', refMatch?.[1]);
                        } else {
                            console.log('  ✗ No ref attribute');
                        }
                    }
                }
                
                // Also check U83 and V83 (the spill cells)
                const u83Match = sheetXml.match(/<c r="U83"[^>]*>.*?<\/c>/s);
                const v83Match = sheetXml.match(/<c r="V83"[^>]*>.*?<\/c>/s);
                
                if (u83Match) {
                    console.log('\nU83 (spill cell) XML:');
                    console.log(u83Match[0].substring(0, 200));
                }
                
                break;
            }
        } catch (e) {
            // Sheet doesn't exist
        }
    }
    
    // 2. Check exported file's XML
    console.log('\n2. EXPORTED export.xlsx - T83 CELL XML');
    console.log('----------------------------------------');
    
    const exportZip = new JSZip();
    await exportZip.loadAsync(fs.readFileSync('export.xlsx'));
    
    // Find DCF sheet in export
    for (let i = 1; i <= 13; i++) {
        try {
            const sheetXml = await exportZip.file(`xl/worksheets/sheet${i}.xml`).async('string');
            if (sheetXml.includes('TRANSPOSE')) {
                console.log(`Found TRANSPOSE in sheet${i}.xml`);
                
                // Extract T83 cell
                const t83Match = sheetXml.match(/<c r="T83"[^>]*>.*?<\/c>/s);
                if (t83Match) {
                    console.log('\nT83 cell XML in export:');
                    console.log(t83Match[0]);
                    
                    // Check formula
                    const formulaMatch = t83Match[0].match(/<f[^>]*>(.*?)<\/f>/);
                    if (formulaMatch) {
                        console.log('\nFormula element:', formulaMatch[0]);
                        
                        // Check for array attributes
                        if (!formulaMatch[0].includes('t="array"')) {
                            console.log('⚠️ MISSING t="array" - This causes Excel to add @ symbols!');
                        }
                        if (!formulaMatch[0].includes('ref=')) {
                            console.log('⚠️ MISSING ref attribute - This prevents proper spilling!');
                        }
                    }
                }
                
                // Check U83 and V83
                const u83Match = sheetXml.match(/<c r="U83"[^>]*>.*?<\/c>/s);
                if (u83Match) {
                    console.log('\nU83 in export:');
                    console.log(u83Match[0].substring(0, 200));
                }
                
                break;
            }
        } catch (e) {
            // Sheet doesn't exist
        }
    }
    
    // 3. What Excel expects for dynamic array formulas
    console.log('\n3. EXCEL DYNAMIC ARRAY FORMULA REQUIREMENTS');
    console.log('---------------------------------------------');
    console.log('For TRANSPOSE to work without @ symbols:');
    console.log('1. Formula cell needs: <f t="array" ref="T83:V83">TRANSPOSE(...)</f>');
    console.log('2. The ref attribute defines the spill range');
    console.log('3. Spill cells (U83, V83) should only have values, no formulas');
    console.log('4. Without t="array", Excel treats it as a regular formula and adds @');
    
    // 4. Create a correct example
    console.log('\n4. CREATING CORRECT TRANSPOSE EXAMPLE');
    console.log('--------------------------------------');
    
    // We need to manually construct the correct XML structure
    // This is what ExcelJS should be producing
    console.log('Correct formula XML should be:');
    console.log('<c r="T83" s="1" t="n">');
    console.log('  <f t="array" ref="T83:V83">TRANSPOSE($N$43:$N$45)</f>');
    console.log('  <v>32.8786327536365</v>');
    console.log('</c>');
}

analyzeXmlStructure().catch(console.error);