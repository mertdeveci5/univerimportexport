const fs = require('fs');
const XLSX = require('xlsx');

const data = fs.readFileSync('test.xlsx');
const workbook = XLSX.read(data, {type: 'buffer'});

// Find the DCF sheet
const dcfSheet = workbook.Sheets['DCF'];
if (!dcfSheet) {
  console.log('DCF sheet not found in test.xlsx');
  process.exit(1);
}

console.log('=== Searching for TRANSPOSE formulas in DCF sheet ===');
let transposeCount = 0;

// Scan all cells
for (let cell in dcfSheet) {
  if (cell[0] === '!') continue; // Skip metadata
  
  const cellData = dcfSheet[cell];
  if (cellData.f && cellData.f.toUpperCase().includes('TRANSPOSE')) {
    console.log(`Found TRANSPOSE at ${cell}:`, cellData.f);
    transposeCount++;
  }
}

if (transposeCount === 0) {
  console.log('No TRANSPOSE formulas found in DCF sheet');
  console.log('\nChecking all formulas in column U and V around row 83:');
  
  // Check U and V columns around row 83
  const checkCells = ['T82', 'T83', 'T84', 'U82', 'U83', 'U84', 'V82', 'V83', 'V84'];
  for (let addr of checkCells) {
    if (dcfSheet[addr]) {
      console.log(`${addr}:`, JSON.stringify(dcfSheet[addr], null, 2));
    }
  }
}

console.log('\n=== Now checking import to see what Univer sees ===');