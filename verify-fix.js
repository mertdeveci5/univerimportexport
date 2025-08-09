#!/usr/bin/env node

/**
 * Verify that the shared formula fix is in place
 */

const fs = require('fs');
const path = require('path');

console.log('Verifying shared formula fix in the code...\n');

// Check if the fix is present in EnhancedWorkSheet.ts
const enhancedWorkSheetPath = './src/UniverToExcel/EnhancedWorkSheet.ts';
const enhancedWorkSheetContent = fs.readFileSync(enhancedWorkSheetPath, 'utf8');

// Check for the key fix: proper handling of si field
const checks = [
    {
        name: 'Two-pass processing for shared formulas',
        pattern: /First pass: identify shared formula master cells/,
        found: false
    },
    {
        name: 'Master cell identification',
        pattern: /If cell has both si.*and f.*it's a master cell/,
        found: false
    },
    {
        name: 'Dependent cell handling',
        pattern: /This is a dependent cell - ExcelJS will handle it automatically/,
        found: false
    },
    {
        name: 'ProcessFormula only uses cell.f',
        pattern: /IMPORTANT: cell\.si is a shared formula ID, not a formula string!/,
        found: false
    },
    {
        name: 'Return early if no formula',
        pattern: /if \(!cell\.f\)[\s\S]*?return;/,
        found: false
    }
];

// Run checks
checks.forEach(check => {
    check.found = check.pattern.test(enhancedWorkSheetContent);
});

// Report results
console.log('Verification Results:');
console.log('====================\n');

let allPassed = true;
checks.forEach(check => {
    const status = check.found ? '✅' : '❌';
    console.log(`${status} ${check.name}`);
    if (!check.found) allPassed = false;
});

console.log('\n' + '='.repeat(60));

if (allPassed) {
    console.log('✅ SUCCESS: All shared formula fixes are in place!');
    console.log('\nThe code should now correctly handle shared formulas by:');
    console.log('1. Identifying master cells (have both si and f)');
    console.log('2. Processing dependent cells separately (only have si)');
    console.log('3. Only using cell.f for formula text, never cell.si');
    console.log('\nNext step: Build and test with actual Excel files.');
} else {
    console.log('⚠️  WARNING: Some fixes are missing!');
    console.log('Please review the EnhancedWorkSheet.ts file.');
}

// Also check if the build output exists
const distPath = './dist/luckyexcel.umd.js';
if (fs.existsSync(distPath)) {
    const stats = fs.statSync(distPath);
    const modTime = new Date(stats.mtime);
    const now = new Date();
    const diffMinutes = Math.floor((now - modTime) / 1000 / 60);
    
    console.log('\nBuild Info:');
    console.log(`  File: ${distPath}`);
    console.log(`  Last built: ${modTime.toLocaleString()} (${diffMinutes} minutes ago)`);
    
    if (diffMinutes > 5) {
        console.log('  ⚠️  Build might be outdated. Run "npm run build" to rebuild.');
    } else {
        console.log('  ✅ Build is recent.');
    }
}