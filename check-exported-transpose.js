const fs = require('fs');
const XLSX = require('xlsx');

// Read the exported file
const data = fs.readFileSync('export.xlsx');
const workbook = XLSX.read(data, {type: 'buffer'});

// Find the DCF sheet
const dcfSheet = workbook.Sheets['DCF'];
if (!dcfSheet) {
  console.log('DCF sheet not found in export.xlsx');
  process.exit(1);
}

console.log('=== Checking TRANSPOSE formulas in exported file (export.xlsx) ===\n');

// Check known TRANSPOSE locations
const transposeCells = [
  'T83', 'W83', 'Z83', 'T84', 
  'T100', 'W100', 'Z100',
  'T107', 'W107', 'Z107',
  'T108', 'W108', 'Z108',
  'T124', 'W124', 'Z124'
];

let hasAtSymbol = false;

for (let cell of transposeCells) {
  if (dcfSheet[cell] && dcfSheet[cell].f) {
    const formula = dcfSheet[cell].f;
    console.log(`${cell}: ${formula}`);
    
    if (formula.includes('@')) {
      console.log(`  ⚠️ HAS @ SYMBOL!`);
      hasAtSymbol = true;
      
      // Show what it should be
      const cleanedFormula = formula.replace(/@/g, '');
      console.log(`  ✅ Should be: ${cleanedFormula}`);
    }
  }
}

if (hasAtSymbol) {
  console.log('\n❌ PROBLEM CONFIRMED: Exported file contains @ symbols in TRANSPOSE formulas!');
  console.log('This causes Excel to reject the file when opened.');
} else {
  console.log('\n✅ No @ symbols found in TRANSPOSE formulas.');
}