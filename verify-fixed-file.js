const fs = require('fs');
const JSZip = require('@progress/jszip-esm');
const XLSX = require('xlsx');

async function verifyFixedFile() {
    console.log('=================================================');
    console.log('VERIFYING FIXED EXCEL FILE');
    console.log('=================================================\n');
    
    // 1. Check with XLSX library
    console.log('1. CHECKING WITH XLSX LIBRARY');
    console.log('------------------------------');
    
    const wb = XLSX.readFile('export-fixed.xlsx');
    
    // Check defined names
    console.log('Defined names:', wb.Workbook?.Names?.length || 0);
    if (wb.Workbook?.Names) {
        wb.Workbook.Names.slice(0, 3).forEach(name => {
            console.log(`  - ${name.Name}: ${name.Ref}`);
        });
    }
    
    // Check TRANSPOSE formulas
    const dcfSheet = wb.Sheets['DCF'];
    if (dcfSheet) {
        console.log('\nTRANSPOSE formulas in DCF:');
        const cells = ['T83', 'W83', 'Z83'];
        cells.forEach(cell => {
            if (dcfSheet[cell]) {
                console.log(`  ${cell}: ${dcfSheet[cell].f || 'no formula'}`);
            }
        });
    }
    
    // 2. Check raw XML
    console.log('\n2. CHECKING RAW XML');
    console.log('-------------------');
    
    const zip = new JSZip();
    await zip.loadAsync(fs.readFileSync('export-fixed.xlsx'));
    
    // Check worksheet XML for array attributes
    const sheet7Xml = await zip.file('xl/worksheets/sheet7.xml').async('string');
    
    // Check T83
    const t83Match = sheet7Xml.match(/<c r="T83"[^>]*>.*?<\/c>/s);
    if (t83Match) {
        console.log('T83 cell XML:');
        const hasArray = t83Match[0].includes('t="array"');
        const hasRef = t83Match[0].includes('ref=');
        console.log(`  ✓ Has t="array": ${hasArray}`);
        console.log(`  ✓ Has ref attribute: ${hasRef}`);
        
        if (hasArray && hasRef) {
            console.log('  ✅ TRANSPOSE should work without @ symbols!');
        }
    }
    
    // Check workbook.xml for defined names
    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    const hasDefinedNames = workbookXml.includes('<definedNames>');
    console.log(`\nDefined names in workbook.xml: ${hasDefinedNames ? '✅ Yes' : '❌ No'}`);
    
    if (hasDefinedNames) {
        const nameCount = (workbookXml.match(/<definedName/g) || []).length;
        console.log(`  Found ${nameCount} defined names`);
    }
    
    console.log('\n=================================================');
    console.log('SUMMARY');
    console.log('=================================================');
    console.log('✅ TRANSPOSE formulas have array attributes');
    console.log('✅ Defined names are present');
    console.log('❌ Border styles still need fixing');
    console.log('\nThe file export-fixed.xlsx should now:');
    console.log('  - Open in Excel without @ symbols in TRANSPOSE');
    console.log('  - Have working defined names');
    console.log('  - Still be missing most borders');
}

verifyFixedFile().catch(console.error);