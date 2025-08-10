const fs = require('fs');

console.log('=================================================');
console.log('ALTERNATIVE EXCEL LIBRARIES ANALYSIS');
console.log('=================================================\n');

// Check what's already installed
const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));

console.log('1. CURRENT LIBRARIES IN USE');
console.log('----------------------------');
console.log('For Import:');
console.log('  - xlsx:', packageJson.dependencies['xlsx'] || 'Not found');
console.log('  - @progress/jszip-esm:', packageJson.dependencies['@progress/jszip-esm'] || 'Not found');

console.log('\nFor Export:');
console.log('  - @zwight/exceljs:', packageJson.dependencies['@zwight/exceljs'] || 'Not found');

console.log('\n2. ALTERNATIVE LIBRARY OPTIONS');
console.log('-------------------------------');

const alternatives = [
    {
        name: 'xlsx (SheetJS)',
        npm: 'xlsx',
        pros: [
            'Can read AND write Excel files',
            'Supports array formulas in Pro version',
            'Can manipulate raw XML',
            'Already used for import'
        ],
        cons: [
            'Array formula support might be limited in free version',
            'Less intuitive API for complex operations'
        ],
        arrayFormulaSupport: 'Partial (Pro version has full support)'
    },
    {
        name: 'xlsx-populate',
        npm: 'xlsx-populate',
        pros: [
            'Built on JSZip',
            'Can manipulate raw XML directly',
            'Good for formula manipulation',
            'Lighter weight than ExcelJS'
        ],
        cons: [
            'Less maintained',
            'Smaller community'
        ],
        arrayFormulaSupport: 'Can write raw XML, so yes with manual work'
    },
    {
        name: 'excel4node',
        npm: 'excel4node',
        pros: [
            'Native Excel file writing',
            'Good performance'
        ],
        cons: [
            'No array formula support',
            'Write-only (no read)'
        ],
        arrayFormulaSupport: 'No'
    },
    {
        name: 'node-excel-export',
        npm: 'node-excel-export',
        pros: [
            'Simple API'
        ],
        cons: [
            'Very limited features',
            'No array formula support'
        ],
        arrayFormulaSupport: 'No'
    },
    {
        name: 'Manual XML Manipulation',
        npm: 'jszip + xml2js',
        pros: [
            'Full control over XML',
            'Can add any Excel feature',
            'Already have JSZip'
        ],
        cons: [
            'More complex to implement',
            'Need to understand Excel XML structure'
        ],
        arrayFormulaSupport: 'Yes - full control'
    }
];

alternatives.forEach(lib => {
    console.log(`\n${lib.name} (${lib.npm})`);
    console.log('  Array Formulas:', lib.arrayFormulaSupport);
    console.log('  Pros:', lib.pros.join(', '));
    console.log('  Cons:', lib.cons.join(', '));
});

console.log('\n3. HYBRID APPROACH OPTIONS');
console.log('---------------------------');
console.log('Option A: Use ExcelJS + Post-process XML');
console.log('  1. Let ExcelJS create the base Excel file');
console.log('  2. Unzip with JSZip');
console.log('  3. Modify XML to add array formula attributes');
console.log('  4. Modify workbook.xml to add defined names');
console.log('  5. Re-zip');

console.log('\nOption B: Use XLSX for both import and export');
console.log('  1. Already using XLSX for reading in some places');
console.log('  2. XLSX can write files too');
console.log('  3. Has better formula preservation');

console.log('\nOption C: Use xlsx-populate for export');
console.log('  1. Better control over formulas');
console.log('  2. Can manipulate XML directly when needed');

console.log('\n4. RECOMMENDED APPROACH');
console.log('------------------------');
console.log('SHORT TERM (Quick Fix):');
console.log('  Post-process ExcelJS output with JSZip to fix:');
console.log('  - Add t="array" and ref attributes to TRANSPOSE');
console.log('  - Add definedNames section to workbook.xml');
console.log('  - Preserve border styles');

console.log('\nLONG TERM (Better Solution):');
console.log('  Replace ExcelJS with XLSX (SheetJS) for export since:');
console.log('  - Already using it for import');
console.log('  - Better formula handling');
console.log('  - More control over output');

console.log('\n5. TEST XLSX WRITE CAPABILITY');
console.log('------------------------------');
const XLSX = require('xlsx');

// Create a simple workbook with TRANSPOSE
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([
    ['=TRANSPOSE(D1:D3)', null, null, 1],
    [null, null, null, 2],
    [null, null, null, 3]
]);

// Try to set array formula (this is the key test)
if (ws['A1']) {
    console.log('Setting TRANSPOSE in XLSX...');
    ws['A1'].f = 'TRANSPOSE(D1:D3)';
    // Note: XLSX free version doesn't support array formula attributes directly
    // But we could potentially modify the XML after writing
}

XLSX.utils.book_append_sheet(wb, ws, 'Test');

// Add a defined name
if (!wb.Workbook) wb.Workbook = {};
if (!wb.Workbook.Names) wb.Workbook.Names = [];
wb.Workbook.Names.push({
    Name: 'TestRange',
    Ref: 'Test!$D$1:$D$3'
});

// Write the file
XLSX.writeFile(wb, 'test-xlsx-write.xlsx');
console.log('✓ Created test-xlsx-write.xlsx with XLSX library');
console.log('  - Has TRANSPOSE formula');
console.log('  - Has defined name');
console.log('\nCheck if this file opens correctly in Excel.');