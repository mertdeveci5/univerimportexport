#!/usr/bin/env node

/**
 * Analyze the fixed export to verify formulas are preserved
 */

const fs = require('fs');
const XLSX = require('xlsx');

console.log('Analyzing fixed export...\n');

// Read both files
const originalWB = XLSX.readFile('test.xlsx');
const fixedWB = XLSX.readFile('test-fixed.xlsx');

// Focus on Operational Assumptions sheet
const sheetName = 'Operational Assumptions';
const originalSheet = originalWB.Sheets[sheetName];
const fixedSheet = fixedWB.Sheets[sheetName];

if (!originalSheet || !fixedSheet) {
    console.log('❌ Could not find Operational Assumptions sheet');
    process.exit(1);
}

// Check specific cells that had corruption issues
const testCells = [
    'O16', 'P16', 'Q16', 'R16', 'S16', 'T16', 'U16', 'V16',
    'O54', 'P54', 'Q54', 'R54', 'S54', 'T54', 'U54', 'V54',
    'N60', 'O60', 'P60', 'Q60', 'R60', 'S60', 'T60', 'U60', 'V60',
    'O74', 'P74', 'Q74', 'R74', 'S74', 'T74', 'U74', 'V74',
    'O86', 'P86', 'Q86', 'R86', 'S86', 'T86', 'U86', 'V86'
];

let totalChecked = 0;
let totalMatched = 0;
let totalMismatched = 0;
const mismatches = [];

console.log('Checking formulas in key cells...\n');

testCells.forEach(cell => {
    const origCell = originalSheet[cell];
    const fixedCell = fixedSheet[cell];
    
    if (origCell && origCell.f) {
        totalChecked++;
        
        if (fixedCell && fixedCell.f) {
            if (origCell.f === fixedCell.f) {
                totalMatched++;
                console.log(`✅ ${cell}: Formula preserved`);
            } else {
                totalMismatched++;
                mismatches.push({
                    cell,
                    original: origCell.f,
                    fixed: fixedCell.f
                });
                console.log(`❌ ${cell}: Formula CORRUPTED`);
                console.log(`   Original: ${origCell.f}`);
                console.log(`   Fixed:    ${fixedCell.f}`);
            }
        } else {
            totalMismatched++;
            mismatches.push({
                cell,
                original: origCell.f,
                fixed: 'Missing'
            });
            console.log(`❌ ${cell}: Formula MISSING in fixed export`);
        }
    }
});

console.log('\n' + '='.repeat(60));
console.log('SUMMARY');
console.log('='.repeat(60));
console.log(`Total formulas checked: ${totalChecked}`);
console.log(`Formulas preserved correctly: ${totalMatched}`);
console.log(`Formulas corrupted/missing: ${totalMismatched}`);

if (totalMismatched === 0) {
    console.log('\n🎉 SUCCESS! All formulas were preserved correctly!');
} else {
    console.log(`\n⚠️  FAILURE: ${totalMismatched} formulas were corrupted.`);
    console.log('\nCorruption pattern analysis:');
    
    // Analyze the corruption pattern
    const patterns = {};
    mismatches.forEach(m => {
        if (m.fixed !== 'Missing') {
            // Extract the changed references
            const origRefs = (m.original.match(/\$?[A-Z]+\$?\d+/g) || []);
            const fixedRefs = (m.fixed.match(/\$?[A-Z]+\$?\d+/g) || []);
            
            for (let i = 0; i < Math.min(origRefs.length, fixedRefs.length); i++) {
                if (origRefs[i] !== fixedRefs[i]) {
                    const pattern = `${origRefs[i]} -> ${fixedRefs[i]}`;
                    if (!patterns[pattern]) {
                        patterns[pattern] = 0;
                    }
                    patterns[pattern]++;
                }
            }
        }
    });
    
    Object.entries(patterns).forEach(([pattern, count]) => {
        console.log(`  ${pattern}: ${count} occurrences`);
    });
}

process.exit(totalMismatched > 0 ? 1 : 0);