const fs = require('fs');

// Check how borders are being handled in the import process
console.log('=================================================');
console.log('BORDER IMPORT INVESTIGATION');
console.log('=================================================\n');

// 1. Check style.ts for border handling
const styleContent = fs.readFileSync('src/ToLuckySheet/style.ts', 'utf8');

console.log('1. BORDER HANDLING IN style.ts');
console.log('--------------------------------');

// Look for border-related functions
const borderFunctions = styleContent.match(/function.*border.*\{/gi);
if (borderFunctions) {
    console.log('Border-related functions found:');
    borderFunctions.forEach(func => console.log(`  - ${func}`));
}

// Look for border property mapping
if (styleContent.includes('borderTop') || styleContent.includes('border-top')) {
    console.log('✓ Handles border-top');
}
if (styleContent.includes('borderBottom') || styleContent.includes('border-bottom')) {
    console.log('✓ Handles border-bottom');
}
if (styleContent.includes('borderLeft') || styleContent.includes('border-left')) {
    console.log('✓ Handles border-left');
}
if (styleContent.includes('borderRight') || styleContent.includes('border-right')) {
    console.log('✓ Handles border-right');
}

// 2. Check LuckySheet.ts
console.log('\n2. BORDER HANDLING IN LuckySheet.ts');
console.log('------------------------------------');

const luckySheetContent = fs.readFileSync('src/ToLuckySheet/LuckySheet.ts', 'utf8');

// Look for borderInfo
if (luckySheetContent.includes('borderInfo')) {
    console.log('✓ LuckySheet.ts uses borderInfo');
    
    // Find how it's used
    const borderInfoLines = luckySheetContent.split('\n').filter(line => line.includes('borderInfo'));
    console.log(`Found ${borderInfoLines.length} references to borderInfo`);
}

// 3. Check export side
console.log('\n3. BORDER EXPORT IN UniverToExcel');
console.log('----------------------------------');

const univerUtilsContent = fs.readFileSync('src/UniverToExcel/univerUtils.ts', 'utf8');

if (univerUtilsContent.includes('convertBorderToExcel')) {
    console.log('✓ Has convertBorderToExcel function');
}

// 4. Check LuckyToUniver conversion
console.log('\n4. BORDER CONVERSION IN LuckyToUniver');
console.log('--------------------------------------');

const univerSheetContent = fs.readFileSync('src/LuckyToUniver/UniverSheet.ts', 'utf8');

// Check if borders are handled
if (univerSheetContent.includes('border')) {
    const borderLines = univerSheetContent.split('\n').filter(line => line.toLowerCase().includes('border'));
    console.log(`Found ${borderLines.length} references to border`);
    
    // Show first few
    borderLines.slice(0, 3).forEach(line => {
        console.log(`  Line: ${line.trim()}`);
    });
}

// 5. Test a simple border import/export
console.log('\n5. TESTING BORDER FLOW');
console.log('-----------------------');
console.log('To fully test borders, we need to:');
console.log('1. Import test.xlsx and check if borderInfo is populated');
console.log('2. Check if borders are in the Univer data structure');
console.log('3. Check if borders are exported to Excel');