const fs = require('fs');
const JSZip = require('@progress/jszip-esm');

/**
 * Analyze what's in the export.xlsx to understand the issues
 */

async function analyzeExport() {
    console.log('=====================================');
    console.log('ANALYZING EXPORT.XLSX');
    console.log('=====================================\n');
    
    const zip = new JSZip();
    const content = fs.readFileSync('export.xlsx');
    await zip.loadAsync(content);
    
    // 1. Check for TRANSPOSE formulas
    console.log('1. CHECKING TRANSPOSE FORMULAS');
    console.log('-------------------------------');
    
    const worksheetFiles = Object.keys(zip.files).filter(f => f.match(/xl\/worksheets\/sheet\d+\.xml/));
    
    for (const sheetPath of worksheetFiles) {
        const sheetXml = await zip.file(sheetPath).async('string');
        
        // Look for TRANSPOSE
        const transposeMatches = sheetXml.match(/<f[^>]*>([^<]*TRANSPOSE[^<]*)<\/f>/gi);
        
        if (transposeMatches) {
            console.log(`\nFound in ${sheetPath}:`);
            transposeMatches.slice(0, 3).forEach(match => {
                // Extract just the formula
                const formulaMatch = match.match(/<f[^>]*>([^<]*)<\/f>/);
                if (formulaMatch) {
                    const formula = formulaMatch[1];
                    const hasArrayAttr = match.includes('t="array"');
                    const hasRef = match.includes('ref=');
                    const hasCurlyBraces = formula.includes('{') || formula.includes('}');
                    const hasAtSymbol = formula.includes('@');
                    
                    console.log(`  Formula: ${formula}`);
                    console.log(`    Has t="array": ${hasArrayAttr}`);
                    console.log(`    Has ref attr: ${hasRef}`);
                    console.log(`    Has curly braces: ${hasCurlyBraces}`);
                    console.log(`    Has @ symbol: ${hasAtSymbol}`);
                }
            });
        }
    }
    
    // 2. Check workbook.xml for defined names
    console.log('\n2. CHECKING DEFINED NAMES');
    console.log('--------------------------');
    
    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    const definedNamesMatch = workbookXml.match(/<definedNames>.*?<\/definedNames>/s);
    
    if (definedNamesMatch) {
        const names = definedNamesMatch[0].match(/<definedName[^>]*>[^<]*<\/definedName>/g);
        console.log(`Found ${names ? names.length : 0} defined names`);
        if (names) {
            names.slice(0, 3).forEach(name => {
                console.log(`  ${name}`);
            });
        }
    } else {
        console.log('No <definedNames> section found');
    }
    
    // 3. Check styles.xml for borders
    console.log('\n3. CHECKING BORDER STYLES');
    console.log('-------------------------');
    
    const stylesXml = await zip.file('xl/styles.xml').async('string');
    const bordersMatch = stylesXml.match(/<borders count="(\d+)">/);
    
    if (bordersMatch) {
        console.log(`Border count: ${bordersMatch[1]}`);
        
        // Check if there are actual border definitions
        const borderDefs = stylesXml.match(/<border[^>]*>.*?<\/border>/gs);
        if (borderDefs) {
            console.log(`Actual border definitions: ${borderDefs.length}`);
            
            // Count non-empty borders
            let nonEmptyBorders = 0;
            borderDefs.forEach(border => {
                if (border.includes('style=')) {
                    nonEmptyBorders++;
                }
            });
            console.log(`Non-empty borders: ${nonEmptyBorders}`);
        }
    }
    
    // 4. Check for external links
    console.log('\n4. CHECKING FOR EXTERNAL LINKS');
    console.log('-------------------------------');
    
    // Check if externalLinks directory exists
    const hasExternalLinks = Object.keys(zip.files).some(f => f.includes('externalLinks'));
    console.log(`Has externalLinks directory: ${hasExternalLinks}`);
    
    // Check for external references in formulas
    for (const sheetPath of worksheetFiles) {
        const sheetXml = await zip.file(sheetPath).async('string');
        
        // Look for external references (e.g., [1], [2], etc.)
        const externalRefs = sheetXml.match(/\[(\d+)\]/g);
        if (externalRefs) {
            console.log(`\nFound external references in ${sheetPath}:`);
            const unique = [...new Set(externalRefs)];
            console.log(`  References: ${unique.join(', ')}`);
            
            // Show a sample formula with external ref
            const sampleFormula = sheetXml.match(/<f[^>]*>[^<]*\[\d+\][^<]*<\/f>/);
            if (sampleFormula) {
                console.log(`  Sample formula: ${sampleFormula[0]}`);
            }
        }
    }
    
    console.log('\n=====================================\n');
}

analyzeExport().catch(console.error);