const fs = require('fs');
const JSZip = require('@progress/jszip-esm');

async function analyzeBorderProblem() {
    console.log('=================================================');
    console.log('BORDER STYLES DEEP ANALYSIS');
    console.log('=================================================\n');
    
    // 1. Analyze original file borders
    console.log('1. ORIGINAL FILE (test.xlsx) BORDERS');
    console.log('-------------------------------------');
    
    const originalZip = new JSZip();
    await originalZip.loadAsync(fs.readFileSync('test.xlsx'));
    
    const originalStyles = await originalZip.file('xl/styles.xml').async('string');
    
    // Extract border definitions
    const originalBordersMatch = originalStyles.match(/<borders[^>]*>(.*?)<\/borders>/s);
    if (originalBordersMatch) {
        const borderCount = (originalBordersMatch[0].match(/<border>/g) || []).length;
        console.log(`Border definitions: ${borderCount}`);
        
        // Show a sample border
        const sampleBorder = originalStyles.match(/<border>.*?<\/border>/);
        if (sampleBorder) {
            console.log('Sample border XML:');
            console.log(sampleBorder[0].substring(0, 200) + '...');
        }
    }
    
    // Check how many cells use borders
    const sheet7Xml = await originalZip.file('xl/worksheets/sheet7.xml').async('string');
    const cellsWithStyle = sheet7Xml.match(/s="\d+"/g) || [];
    console.log(`\nCells with style attribute: ${cellsWithStyle.length}`);
    
    // Get unique style IDs
    const styleIds = new Set(cellsWithStyle.map(s => s.match(/\d+/)[0]));
    console.log(`Unique style IDs used: ${styleIds.size}`);
    console.log('Sample style IDs:', Array.from(styleIds).slice(0, 10).join(', '));
    
    // 2. Analyze exported file borders
    console.log('\n2. EXPORTED FILE (export.xlsx) BORDERS');
    console.log('---------------------------------------');
    
    const exportZip = new JSZip();
    await exportZip.loadAsync(fs.readFileSync('export.xlsx'));
    
    const exportStyles = await exportZip.file('xl/styles.xml').async('string');
    
    const exportBordersMatch = exportStyles.match(/<borders[^>]*>(.*?)<\/borders>/s);
    if (exportBordersMatch) {
        const borderCount = (exportBordersMatch[0].match(/<border>/g) || []).length;
        console.log(`Border definitions: ${borderCount}`);
        
        // Show the borders section
        console.log('Borders XML:');
        console.log(exportBordersMatch[0].substring(0, 500));
    }
    
    // 3. Check what's happening during import
    console.log('\n3. IMPORT PROCESS ANALYSIS');
    console.log('--------------------------');
    
    // Read the import code to understand border handling
    const luckySheetCode = fs.readFileSync('src/ToLuckySheet/LuckySheet.ts', 'utf8');
    
    // Check if borderInfo is being populated
    if (luckySheetCode.includes('borderInfo')) {
        console.log('✓ LuckySheet.ts processes borderInfo');
        
        // Find how borders are read
        const borderMatches = luckySheetCode.match(/borderInfo.*?[\n\r]/g);
        if (borderMatches) {
            console.log(`Found ${borderMatches.length} references to borderInfo`);
        }
    }
    
    // 4. Check Univer data structure
    console.log('\n4. UNIVER DATA STRUCTURE');
    console.log('------------------------');
    console.log('To check: Does Univer store border data in:');
    console.log('  - styles object?');
    console.log('  - cell.s property?');
    console.log('  - borderInfo array?');
    
    // 5. Solution approach
    console.log('\n5. SOLUTION APPROACH');
    console.log('--------------------');
    console.log('Option A: Preserve original styles.xml');
    console.log('  1. Keep original styles.xml from input file');
    console.log('  2. Map Univer style IDs to original style IDs');
    console.log('  3. Use original styles in export');
    
    console.log('\nOption B: Rebuild styles from Univer data');
    console.log('  1. Extract all border info from Univer');
    console.log('  2. Generate complete styles.xml');
    console.log('  3. Update cell style references');
    
    console.log('\nOption C: Merge styles');
    console.log('  1. Take ExcelJS generated styles');
    console.log('  2. Add missing border definitions from Univer');
    console.log('  3. Remap style IDs');
}

analyzeBorderProblem().catch(console.error);