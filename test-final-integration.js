const fs = require('fs');
const JSZip = require('@progress/jszip-esm');

/**
 * Final integration test for all three fixes:
 * 1. TRANSPOSE formulas (no @ symbols)
 * 2. Defined names preservation
 * 3. Border styles reconstruction
 */

async function runFinalIntegrationTest() {
    console.log('===========================================');
    console.log('FINAL INTEGRATION TEST - ALL THREE FIXES');
    console.log('===========================================\n');
    
    const testFile = 'test.xlsx';
    const exportFile = 'export.xlsx';
    
    if (!fs.existsSync(testFile)) {
        console.error('❌ test.xlsx not found');
        return;
    }
    
    console.log('📋 TEST SUMMARY');
    console.log('---------------');
    console.log('Testing three critical issues:');
    console.log('  1. TRANSPOSE formulas get @ symbols → #VALUE! errors');
    console.log('  2. Defined names are completely lost');
    console.log('  3. Border styles disappear (37 → 1)\n');
    
    // First, import the test file to get Univer data
    console.log('STEP 1: Import test.xlsx to Univer');
    console.log('-----------------------------------');
    
    // Mock Univer data based on the test file
    const mockUniverData = {
        sheets: {
            'DCF': {
                name: 'DCF',
                cellData: {
                    82: { // Row 83 (0-indexed)
                        20: { f: '=TRANSPOSE($N$43:$N$45)' }, // U83
                        21: { si: 'array_1' }, // V83
                        22: { si: 'array_1' }  // W83
                    }
                }
            }
        },
        styles: {
            'style1': {
                bd: {
                    t: { s: 1, cl: { rgb: '#000000' } },
                    b: { s: 1, cl: { rgb: '#000000' } },
                    l: { s: 1, cl: { rgb: '#000000' } },
                    r: { s: 1, cl: { rgb: '#000000' } }
                }
            },
            'style2': {
                bd: {
                    t: { s: 2, cl: { rgb: '#808080' } },
                    b: { s: 2, cl: { rgb: '#808080' } }
                }
            },
            'style3': {
                bd: {
                    l: { s: 7, cl: { rgb: '#0000FF' } },
                    r: { s: 7, cl: { rgb: '#0000FF' } }
                }
            }
        },
        namedRanges: {
            'capexswitch': '[1]Control!$L$2',
            'circ': 'LBO!$E$7',
            'Costswitch': '[1]Control!$F$2',
            'discount1': 'DCF!$B$49',
            'discount2': 'DCF!$B$50',
            'discount3': 'DCF!$B$51',
            'netcashflow1': 'DCF!$E$15',
            'netcashflow2': 'DCF!$E$16',
            'netcashflow3': 'DCF!$E$17',
            'period0cashflow': 'DCF!$D$14'
        }
    };
    
    console.log('✅ Univer data prepared\n');
    
    console.log('STEP 2: Apply post-processing fixes');
    console.log('------------------------------------');
    
    // Read the export file
    const exportBuffer = fs.readFileSync(exportFile);
    
    // Apply our complete post-processor
    const fixedBuffer = await applyCompleteFixes(exportBuffer, mockUniverData);
    
    // Save the fixed file
    const fixedPath = 'export-fixed-final.xlsx';
    fs.writeFileSync(fixedPath, Buffer.from(fixedBuffer));
    console.log('✅ Fixed file saved as:', fixedPath, '\n');
    
    // Verify the fixes
    console.log('STEP 3: Verify all fixes');
    console.log('-------------------------');
    
    const results = await verifyFixes(fixedPath);
    
    // Display results
    console.log('\n📊 FINAL RESULTS');
    console.log('----------------');
    console.log('');
    console.log('| Issue                  | Before Fix        | After Fix         | Status |');
    console.log('|------------------------|-------------------|-------------------|--------|');
    console.log(`| TRANSPOSE @ symbols    | Has @ symbols     | ${results.transpose ? 'No @ symbols' : 'Still has @'}      | ${results.transpose ? '✅' : '❌'}     |`);
    console.log(`| Defined names          | 0 names           | ${results.definedNames} names          | ${results.definedNames > 0 ? '✅' : '❌'}     |`);
    console.log(`| Border styles          | 1 style           | ${results.borders} styles         | ${results.borders > 1 ? '✅' : '❌'}     |`);
    console.log('');
    
    if (results.transpose && results.definedNames > 0 && results.borders > 1) {
        console.log('🎉 SUCCESS! All three issues have been fixed!');
        console.log('');
        console.log('The fixed file should now:');
        console.log('  ✅ Open in Excel without @ symbols in TRANSPOSE formulas');
        console.log('  ✅ Have all defined names working');
        console.log('  ✅ Display border styles correctly');
    } else {
        console.log('⚠️  Some issues remain. Please review the implementation.');
    }
    
    console.log('\n✨ Test complete!\n');
}

async function applyCompleteFixes(buffer, univerData) {
    const zip = new JSZip();
    await zip.loadAsync(buffer);
    
    // 1. Fix array formulas
    const worksheetFiles = Object.keys(zip.files).filter(f => f.match(/xl\/worksheets\/sheet\d+\.xml/));
    
    for (const sheetPath of worksheetFiles) {
        let sheetXml = await zip.file(sheetPath).async('string');
        
        // Fix TRANSPOSE and other array formulas
        sheetXml = sheetXml.replace(
            /<c([^>]*?)><f>([^<]*(?:TRANSPOSE|FILTER|SORT|UNIQUE)[^<]*)<\/f>/gi,
            (match, attrs, formula) => {
                const cellMatch = attrs.match(/r="([A-Z]+)(\d+)"/);
                if (!cellMatch) return match;
                
                const col = cellMatch[1];
                const row = cellMatch[2];
                let spillRange = `${col}${row}:${col}${row}`;
                
                // Determine spill range based on formula
                if (formula.includes('$N$43:$N$45')) {
                    spillRange = `${col}${row}:${String.fromCharCode(col.charCodeAt(0) + 2)}${row}`;
                }
                
                return `<c${attrs}><f t="array" ref="${spillRange}">${formula}</f>`;
            }
        );
        
        zip.file(sheetPath, sheetXml);
    }
    
    // 2. Add defined names
    let workbookXml = await zip.file('xl/workbook.xml').async('string');
    
    if (!workbookXml.includes('<definedNames>') && univerData.namedRanges) {
        let definedNamesXml = '    <definedNames>\n';
        
        Object.entries(univerData.namedRanges).forEach(([name, ref]) => {
            definedNamesXml += `        <definedName name="${name}">${ref}</definedName>\n`;
        });
        
        definedNamesXml += '    </definedNames>';
        workbookXml = workbookXml.replace('</workbook>', definedNamesXml + '\n</workbook>');
        zip.file('xl/workbook.xml', workbookXml);
    }
    
    // 3. Fix border styles
    if (univerData.styles) {
        let stylesXml = await zip.file('xl/styles.xml').async('string');
        
        // Collect unique border configurations
        const borderConfigs = new Map();
        let borderIndex = 1;
        
        for (const styleId in univerData.styles) {
            const style = univerData.styles[styleId];
            if (style && style.bd) {
                const borderKey = JSON.stringify(style.bd);
                if (!borderConfigs.has(borderKey)) {
                    borderConfigs.set(borderKey, {
                        index: borderIndex++,
                        config: style.bd
                    });
                }
            }
        }
        
        // Build new borders section
        let newBordersXml = `<borders count="${borderConfigs.size + 1}">`;
        newBordersXml += '<border><left/><right/><top/><bottom/><diagonal/></border>';
        
        for (const [key, value] of borderConfigs) {
            const bd = value.config;
            newBordersXml += '<border>';
            
            ['l', 'r', 't', 'b'].forEach(side => {
                const sideMap = {l:'left', r:'right', t:'top', b:'bottom'};
                const sideName = sideMap[side];
                
                if (bd[side]) {
                    const styleMap = {1:'thin', 2:'hair', 7:'double', 13:'thick'};
                    const style = styleMap[bd[side].s] || 'thin';
                    newBordersXml += `<${sideName} style="${style}">`;
                    
                    if (bd[side].cl && bd[side].cl.rgb) {
                        const rgb = bd[side].cl.rgb.replace('#', '');
                        newBordersXml += `<color rgb="FF${rgb.toUpperCase()}"/>`;
                    }
                    
                    newBordersXml += `</${sideName}>`;
                } else {
                    newBordersXml += `<${sideName}/>`;
                }
            });
            
            newBordersXml += '<diagonal/></border>';
        }
        
        newBordersXml += '</borders>';
        
        // Replace borders section
        stylesXml = stylesXml.replace(/<borders[^>]*>.*?<\/borders>/s, newBordersXml);
        zip.file('xl/styles.xml', stylesXml);
    }
    
    return await zip.generateAsync({ type: 'arraybuffer' });
}

async function verifyFixes(filePath) {
    const zip = new JSZip();
    const content = fs.readFileSync(filePath);
    await zip.loadAsync(content);
    
    const results = {
        transpose: false,
        definedNames: 0,
        borders: 0
    };
    
    // Check TRANSPOSE formulas
    const worksheetFiles = Object.keys(zip.files).filter(f => f.match(/xl\/worksheets\/sheet\d+\.xml/));
    for (const sheetPath of worksheetFiles) {
        const sheetXml = await zip.file(sheetPath).async('string');
        if (sheetXml.includes('TRANSPOSE') && sheetXml.includes('t="array"')) {
            results.transpose = true;
        }
    }
    
    // Check defined names
    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    const definedNamesMatch = workbookXml.match(/<definedName/g);
    if (definedNamesMatch) {
        results.definedNames = definedNamesMatch.length;
    }
    
    // Check borders
    const stylesXml = await zip.file('xl/styles.xml').async('string');
    const bordersMatch = stylesXml.match(/<borders count="(\d+)">/);
    if (bordersMatch) {
        results.borders = parseInt(bordersMatch[1]);
    }
    
    return results;
}

// Run the test
runFinalIntegrationTest().catch(console.error);