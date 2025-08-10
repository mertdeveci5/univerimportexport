const fs = require('fs');
const XLSX = require('xlsx');

console.log('=================================================');
console.log('FILE COMPARISON: Original vs Exported');
console.log('=================================================\n');

// Read both files
const originalBuffer = fs.readFileSync('test.xlsx');
const exportedBuffer = fs.readFileSync('export.xlsx');

const originalWb = XLSX.read(originalBuffer, {type: 'buffer', cellStyles: true});
const exportedWb = XLSX.read(exportedBuffer, {type: 'buffer', cellStyles: true});

// 1. Compare Defined Names
console.log('1. DEFINED NAMES COMPARISON');
console.log('----------------------------');
console.log('Original defined names:', originalWb.Workbook?.Names?.length || 0);
if (originalWb.Workbook?.Names) {
    originalWb.Workbook.Names.slice(0, 5).forEach(name => {
        console.log(`  - ${name.Name}: ${name.Ref}`);
    });
    if (originalWb.Workbook.Names.length > 5) {
        console.log(`  ... and ${originalWb.Workbook.Names.length - 5} more`);
    }
}

console.log('\nExported defined names:', exportedWb.Workbook?.Names?.length || 0);
if (exportedWb.Workbook?.Names) {
    exportedWb.Workbook.Names.slice(0, 5).forEach(name => {
        console.log(`  - ${name.Name}: ${name.Ref}`);
    });
}

// 2. Check TRANSPOSE formulas in DCF sheet
console.log('\n2. TRANSPOSE FORMULAS IN DCF SHEET');
console.log('-----------------------------------');
const originalDCF = originalWb.Sheets['DCF'];
const exportedDCF = exportedWb.Sheets['DCF'];

if (originalDCF && exportedDCF) {
    const transposeCells = ['T83', 'W83', 'Z83', 'T84', 'T100', 'W100', 'Z100'];
    
    console.log('Cell    | Original Formula                  | Exported Formula');
    console.log('--------|-----------------------------------|----------------------------------');
    
    transposeCells.forEach(cell => {
        const origFormula = originalDCF[cell]?.f || 'N/A';
        const expFormula = exportedDCF[cell]?.f || 'N/A';
        const hasAt = expFormula.includes('@');
        const status = hasAt ? '❌' : (expFormula === 'N/A' ? '⚠️' : '✅');
        console.log(`${cell}    | ${origFormula.padEnd(33)} | ${status} ${expFormula}`);
    });
}

// 3. Check borders in a sample of cells
console.log('\n3. BORDER STYLES COMPARISON (Sample from DCF sheet)');
console.log('----------------------------------------------------');

function hasBorder(cell) {
    if (!cell || !cell.s) return false;
    return cell.s.border && Object.keys(cell.s.border).length > 0;
}

// Check first 100 cells for borders
let originalBorderCount = 0;
let exportedBorderCount = 0;
const sampleCells = [];

for (let row = 1; row <= 20; row++) {
    for (let col = 0; col < 10; col++) {
        const addr = XLSX.utils.encode_cell({r: row, c: col});
        sampleCells.push(addr);
    }
}

sampleCells.forEach(addr => {
    if (originalDCF && originalDCF[addr] && hasBorder(originalDCF[addr])) {
        originalBorderCount++;
    }
    if (exportedDCF && exportedDCF[addr] && hasBorder(exportedDCF[addr])) {
        exportedBorderCount++;
    }
});

console.log(`Original cells with borders: ${originalBorderCount}/${sampleCells.length}`);
console.log(`Exported cells with borders: ${exportedBorderCount}/${sampleCells.length}`);
console.log(`Border retention rate: ${exportedBorderCount > 0 ? Math.round(exportedBorderCount/originalBorderCount * 100) : 0}%`);

// 4. Sheet count comparison
console.log('\n4. SHEETS COMPARISON');
console.log('--------------------');
console.log('Original sheets:', originalWb.SheetNames.length);
console.log('Exported sheets:', exportedWb.SheetNames.length);

if (originalWb.SheetNames.length !== exportedWb.SheetNames.length) {
    console.log('⚠️ Sheet count mismatch!');
    console.log('Original:', originalWb.SheetNames);
    console.log('Exported:', exportedWb.SheetNames);
}

// 5. Formula count comparison
console.log('\n5. FORMULA COUNT COMPARISON');
console.log('---------------------------');
function countFormulas(sheet) {
    let count = 0;
    for (let cell in sheet) {
        if (cell[0] !== '!' && sheet[cell].f) count++;
    }
    return count;
}

originalWb.SheetNames.forEach(sheetName => {
    const origSheet = originalWb.Sheets[sheetName];
    const expSheet = exportedWb.Sheets[sheetName];
    
    const origCount = countFormulas(origSheet);
    const expCount = expSheet ? countFormulas(expSheet) : 0;
    
    if (origCount > 0 || expCount > 0) {
        const status = origCount === expCount ? '✅' : '⚠️';
        console.log(`${status} ${sheetName}: ${origCount} → ${expCount} formulas`);
    }
});

console.log('\n=================================================');
console.log('END OF COMPARISON');
console.log('=================================================');