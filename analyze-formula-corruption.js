#!/usr/bin/env node

/**
 * Formula Corruption Analysis Tool
 * Compares formulas between original and exported Excel files
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Color codes for terminal output
const colors = {
    reset: '\x1b[0m',
    bright: '\x1b[1m',
    red: '\x1b[31m',
    green: '\x1b[32m',
    yellow: '\x1b[33m',
    blue: '\x1b[34m',
    magenta: '\x1b[35m',
    cyan: '\x1b[36m'
};

function log(message, color = 'reset') {
    console.log(`${colors[color]}${message}${colors.reset}`);
}

function logSection(title) {
    console.log('\n' + '='.repeat(80));
    log(title, 'bright');
    console.log('='.repeat(80));
}

// Convert column number to letter (0-based)
function columnToLetter(col) {
    let result = '';
    let num = col;
    while (num >= 0) {
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26) - 1;
        if (num < 0) break;
    }
    return result;
}

// Parse cell address
function parseCellAddress(address) {
    const match = address.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;
    return {
        col: match[1],
        row: parseInt(match[2])
    };
}

// Analyze formula differences in detail
function analyzeFormula(original, exported, cell) {
    if (!original && !exported) return null;
    if (!original) return { type: 'added', exported };
    if (!exported) return { type: 'removed', original };
    if (original === exported) return null;
    
    // Analyze the difference
    const diff = {
        type: 'modified',
        original,
        exported,
        changes: []
    };
    
    // Check for reference changes
    const origRefs = (original.match(/[A-Z]+\d+/g) || []).sort();
    const expRefs = (exported.match(/[A-Z]+\d+/g) || []).sort();
    
    const addedRefs = expRefs.filter(r => !origRefs.includes(r));
    const removedRefs = origRefs.filter(r => !expRefs.includes(r));
    const commonRefs = origRefs.filter(r => expRefs.includes(r));
    
    if (addedRefs.length > 0) {
        diff.changes.push({ type: 'added_refs', refs: addedRefs });
    }
    if (removedRefs.length > 0) {
        diff.changes.push({ type: 'removed_refs', refs: removedRefs });
    }
    
    // Check if it's just a reference shift pattern
    if (commonRefs.length > 0 && (addedRefs.length > 0 || removedRefs.length > 0)) {
        // Analyze if there's a consistent shift pattern
        const shifts = [];
        for (let i = 0; i < Math.min(origRefs.length, expRefs.length); i++) {
            const origParsed = parseCellAddress(origRefs[i]);
            const expParsed = parseCellAddress(expRefs[i]);
            if (origParsed && expParsed) {
                const colShift = expParsed.col.charCodeAt(0) - origParsed.col.charCodeAt(0);
                const rowShift = expParsed.row - origParsed.row;
                shifts.push({ col: colShift, row: rowShift });
            }
        }
        
        // Check if all shifts are the same
        if (shifts.length > 0) {
            const firstShift = shifts[0];
            const consistent = shifts.every(s => s.col === firstShift.col && s.row === firstShift.row);
            if (consistent && (firstShift.col !== 0 || firstShift.row !== 0)) {
                diff.changes.push({
                    type: 'reference_shift',
                    colShift: firstShift.col,
                    rowShift: firstShift.row
                });
            }
        }
    }
    
    // Check for function changes
    const origFuncs = (original.match(/[A-Z]+(?=\()/g) || []).sort();
    const expFuncs = (exported.match(/[A-Z]+(?=\()/g) || []).sort();
    
    const addedFuncs = expFuncs.filter(f => !origFuncs.includes(f));
    const removedFuncs = origFuncs.filter(f => !expFuncs.includes(f));
    
    if (addedFuncs.length > 0) {
        diff.changes.push({ type: 'added_functions', functions: addedFuncs });
    }
    if (removedFuncs.length > 0) {
        diff.changes.push({ type: 'removed_functions', functions: removedFuncs });
    }
    
    return diff;
}

// Main analysis function
function analyzeFiles(originalFile, exportedFile) {
    logSection('FORMULA CORRUPTION ANALYSIS');
    
    // Read both files
    log('\n📖 Reading files...', 'cyan');
    const originalWB = XLSX.readFile(originalFile);
    const exportedWB = XLSX.readFile(exportedFile);
    
    log(`  Original: ${originalFile}`, 'green');
    log(`  Exported: ${exportedFile}`, 'yellow');
    
    // Analysis results
    const results = {
        totalSheets: 0,
        totalCells: 0,
        totalFormulas: 0,
        corruptedFormulas: 0,
        sheetAnalysis: {},
        patterns: {
            referenceShifts: [],
            missingFormulas: [],
            addedFormulas: [],
            modifiedFormulas: []
        }
    };
    
    // Compare each sheet
    originalWB.SheetNames.forEach(sheetName => {
        if (!exportedWB.SheetNames.includes(sheetName)) {
            log(`\n⚠️  Sheet '${sheetName}' missing in exported file!`, 'red');
            return;
        }
        
        results.totalSheets++;
        
        const origSheet = originalWB.Sheets[sheetName];
        const expSheet = exportedWB.Sheets[sheetName];
        
        const sheetDiffs = [];
        
        // Get all cells from both sheets
        const origCells = Object.keys(origSheet).filter(k => !k.startsWith('!'));
        const expCells = Object.keys(expSheet).filter(k => !k.startsWith('!'));
        const allCells = new Set([...origCells, ...expCells]);
        
        results.totalCells += allCells.size;
        
        // Check each cell
        allCells.forEach(cell => {
            const origCell = origSheet[cell];
            const expCell = expSheet[cell];
            
            // Focus on formulas
            const origFormula = origCell?.f;
            const expFormula = expCell?.f;
            
            if (origFormula || expFormula) {
                results.totalFormulas++;
                
                const diff = analyzeFormula(origFormula, expFormula, cell);
                if (diff) {
                    results.corruptedFormulas++;
                    sheetDiffs.push({
                        cell,
                        ...diff
                    });
                    
                    // Categorize the pattern
                    if (diff.type === 'removed') {
                        results.patterns.missingFormulas.push({
                            sheet: sheetName,
                            cell,
                            formula: diff.original
                        });
                    } else if (diff.type === 'added') {
                        results.patterns.addedFormulas.push({
                            sheet: sheetName,
                            cell,
                            formula: diff.exported
                        });
                    } else if (diff.type === 'modified') {
                        const shiftChange = diff.changes.find(c => c.type === 'reference_shift');
                        if (shiftChange) {
                            results.patterns.referenceShifts.push({
                                sheet: sheetName,
                                cell,
                                colShift: shiftChange.colShift,
                                rowShift: shiftChange.rowShift,
                                original: diff.original,
                                exported: diff.exported
                            });
                        } else {
                            results.patterns.modifiedFormulas.push({
                                sheet: sheetName,
                                cell,
                                original: diff.original,
                                exported: diff.exported,
                                changes: diff.changes
                            });
                        }
                    }
                }
            }
        });
        
        if (sheetDiffs.length > 0) {
            results.sheetAnalysis[sheetName] = sheetDiffs;
        }
    });
    
    return results;
}

// Generate detailed report
function generateReport(results) {
    logSection('ANALYSIS SUMMARY');
    
    log(`\n📊 Overall Statistics:`, 'cyan');
    log(`  Total Sheets Analyzed: ${results.totalSheets}`);
    log(`  Total Cells Checked: ${results.totalCells}`);
    log(`  Total Formulas: ${results.totalFormulas}`);
    log(`  Corrupted Formulas: ${results.corruptedFormulas}`, results.corruptedFormulas > 0 ? 'red' : 'green');
    
    if (results.corruptedFormulas > 0) {
        const corruptionRate = ((results.corruptedFormulas / results.totalFormulas) * 100).toFixed(2);
        log(`  Corruption Rate: ${corruptionRate}%`, 'yellow');
    }
    
    // Pattern Analysis
    logSection('CORRUPTION PATTERNS');
    
    if (results.patterns.referenceShifts.length > 0) {
        log('\n🔄 Reference Shifts:', 'yellow');
        const shiftGroups = {};
        results.patterns.referenceShifts.forEach(shift => {
            const key = `${shift.colShift},${shift.rowShift}`;
            if (!shiftGroups[key]) {
                shiftGroups[key] = [];
            }
            shiftGroups[key].push(shift);
        });
        
        Object.entries(shiftGroups).forEach(([key, shifts]) => {
            const [colShift, rowShift] = key.split(',').map(Number);
            log(`  Shift Pattern: Col ${colShift > 0 ? '+' : ''}${colShift}, Row ${rowShift > 0 ? '+' : ''}${rowShift}`);
            log(`    Affected: ${shifts.length} formula(s)`);
            shifts.slice(0, 3).forEach(s => {
                log(`      ${s.sheet}!${s.cell}: ${s.original} → ${s.exported}`, 'red');
            });
            if (shifts.length > 3) {
                log(`      ... and ${shifts.length - 3} more`, 'cyan');
            }
        });
    }
    
    if (results.patterns.missingFormulas.length > 0) {
        log('\n❌ Missing Formulas:', 'red');
        results.patterns.missingFormulas.slice(0, 5).forEach(f => {
            log(`  ${f.sheet}!${f.cell}: ${f.formula}`);
        });
        if (results.patterns.missingFormulas.length > 5) {
            log(`  ... and ${results.patterns.missingFormulas.length - 5} more`, 'cyan');
        }
    }
    
    if (results.patterns.modifiedFormulas.length > 0) {
        log('\n🔧 Modified Formulas (not just shifts):', 'yellow');
        results.patterns.modifiedFormulas.slice(0, 5).forEach(f => {
            log(`  ${f.sheet}!${f.cell}:`);
            log(`    Original: ${f.original}`, 'green');
            log(`    Exported: ${f.exported}`, 'red');
            f.changes.forEach(c => {
                if (c.type === 'added_refs') {
                    log(`    Added refs: ${c.refs.join(', ')}`, 'yellow');
                } else if (c.type === 'removed_refs') {
                    log(`    Removed refs: ${c.refs.join(', ')}`, 'yellow');
                }
            });
        });
        if (results.patterns.modifiedFormulas.length > 5) {
            log(`  ... and ${results.patterns.modifiedFormulas.length - 5} more`, 'cyan');
        }
    }
    
    // Sheet-specific issues
    logSection('SHEET-SPECIFIC ISSUES');
    
    Object.entries(results.sheetAnalysis).forEach(([sheetName, diffs]) => {
        if (diffs.length > 0) {
            log(`\n📋 ${sheetName}: ${diffs.length} formula issue(s)`, 'yellow');
            
            // Focus on the specific cells mentioned
            const importantCells = ['O14', 'P14', 'Q14', 'R14', 'S14', 'N14', 'M14', 'L14'];
            const relevantDiffs = diffs.filter(d => {
                const parsed = parseCellAddress(d.cell);
                if (!parsed) return false;
                return importantCells.includes(d.cell) || parsed.row === 14;
            });
            
            if (relevantDiffs.length > 0) {
                log('  Row 14 issues:', 'red');
                relevantDiffs.sort((a, b) => a.cell.localeCompare(b.cell)).forEach(d => {
                    log(`    ${d.cell}:`, 'cyan');
                    if (d.original) log(`      Original: ${d.original}`, 'green');
                    if (d.exported) log(`      Exported: ${d.exported}`, 'red');
                });
            }
            
            // Show first few other issues
            const otherDiffs = diffs.filter(d => !relevantDiffs.includes(d)).slice(0, 3);
            if (otherDiffs.length > 0) {
                log('  Other issues:', 'yellow');
                otherDiffs.forEach(d => {
                    log(`    ${d.cell}: ${d.type}`);
                });
            }
        }
    });
    
    return results;
}

// Save report to markdown
function saveMarkdownReport(results, outputFile) {
    let md = '# Formula Corruption Analysis Report\n\n';
    md += `Generated: ${new Date().toISOString()}\n\n`;
    
    md += '## Summary\n\n';
    md += `- **Total Sheets Analyzed:** ${results.totalSheets}\n`;
    md += `- **Total Cells Checked:** ${results.totalCells}\n`;
    md += `- **Total Formulas:** ${results.totalFormulas}\n`;
    md += `- **Corrupted Formulas:** ${results.corruptedFormulas}\n`;
    if (results.corruptedFormulas > 0) {
        const rate = ((results.corruptedFormulas / results.totalFormulas) * 100).toFixed(2);
        md += `- **Corruption Rate:** ${rate}%\n`;
    }
    md += '\n';
    
    md += '## Corruption Patterns\n\n';
    
    if (results.patterns.referenceShifts.length > 0) {
        md += '### Reference Shifts\n\n';
        const shiftGroups = {};
        results.patterns.referenceShifts.forEach(shift => {
            const key = `${shift.colShift},${shift.rowShift}`;
            if (!shiftGroups[key]) {
                shiftGroups[key] = [];
            }
            shiftGroups[key].push(shift);
        });
        
        Object.entries(shiftGroups).forEach(([key, shifts]) => {
            const [colShift, rowShift] = key.split(',').map(Number);
            md += `#### Shift: Column ${colShift > 0 ? '+' : ''}${colShift}, Row ${rowShift > 0 ? '+' : ''}${rowShift}\n\n`;
            md += `Affected formulas: ${shifts.length}\n\n`;
            md += '| Sheet | Cell | Original | Exported |\n';
            md += '|-------|------|----------|----------|\n';
            shifts.forEach(s => {
                md += `| ${s.sheet} | ${s.cell} | \`${s.original}\` | \`${s.exported}\` |\n`;
            });
            md += '\n';
        });
    }
    
    if (results.patterns.modifiedFormulas.length > 0) {
        md += '### Modified Formulas\n\n';
        md += '| Sheet | Cell | Original | Exported | Changes |\n';
        md += '|-------|------|----------|----------|----------|\n';
        results.patterns.modifiedFormulas.forEach(f => {
            const changes = f.changes.map(c => {
                if (c.type === 'added_refs') return `Added: ${c.refs.join(', ')}`;
                if (c.type === 'removed_refs') return `Removed: ${c.refs.join(', ')}`;
                return c.type;
            }).join('; ');
            md += `| ${f.sheet} | ${f.cell} | \`${f.original}\` | \`${f.exported}\` | ${changes} |\n`;
        });
        md += '\n';
    }
    
    if (results.patterns.missingFormulas.length > 0) {
        md += '### Missing Formulas\n\n';
        md += '| Sheet | Cell | Formula |\n';
        md += '|-------|------|----------|\n';
        results.patterns.missingFormulas.forEach(f => {
            md += `| ${f.sheet} | ${f.cell} | \`${f.formula}\` |\n`;
        });
        md += '\n';
    }
    
    md += '## Sheet-by-Sheet Analysis\n\n';
    Object.entries(results.sheetAnalysis).forEach(([sheetName, diffs]) => {
        if (diffs.length > 0) {
            md += `### ${sheetName}\n\n`;
            md += `Total issues: ${diffs.length}\n\n`;
            
            // Group by row for better analysis
            const rowGroups = {};
            diffs.forEach(d => {
                const parsed = parseCellAddress(d.cell);
                if (parsed) {
                    const row = parsed.row;
                    if (!rowGroups[row]) {
                        rowGroups[row] = [];
                    }
                    rowGroups[row].push(d);
                }
            });
            
            // Sort rows and show details
            Object.keys(rowGroups).sort((a, b) => Number(a) - Number(b)).forEach(row => {
                const rowDiffs = rowGroups[row];
                md += `#### Row ${row}\n\n`;
                md += '| Cell | Type | Original | Exported |\n';
                md += '|------|------|----------|----------|\n';
                rowDiffs.sort((a, b) => a.cell.localeCompare(b.cell)).forEach(d => {
                    md += `| ${d.cell} | ${d.type} | \`${d.original || 'N/A'}\` | \`${d.exported || 'N/A'}\` |\n`;
                });
                md += '\n';
            });
        }
    });
    
    fs.writeFileSync(outputFile, md);
    log(`\n💾 Report saved to: ${outputFile}`, 'green');
}

// Main execution
function main() {
    const originalFile = 'test.xlsx';
    const exportedFile = 'test-corrupt.xlsx';
    const reportFile = 'FORMULA_CORRUPTION_REPORT.md';
    
    if (!fs.existsSync(originalFile)) {
        log(`❌ Original file not found: ${originalFile}`, 'red');
        process.exit(1);
    }
    
    if (!fs.existsSync(exportedFile)) {
        log(`❌ Exported file not found: ${exportedFile}`, 'red');
        process.exit(1);
    }
    
    const results = analyzeFiles(originalFile, exportedFile);
    generateReport(results);
    saveMarkdownReport(results, reportFile);
    
    logSection('ANALYSIS COMPLETE');
    
    if (results.corruptedFormulas > 0) {
        log('⚠️  Formula corruption detected! Check the report for details.', 'red');
        process.exit(1);
    } else {
        log('✅ No formula corruption detected.', 'green');
    }
}

// Run the analysis
main();