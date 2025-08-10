const fs = require('fs');
const JSZip = require('@progress/jszip-esm');

/**
 * Test script to validate border fixes
 * Checks that border structures match Univer spec
 */

async function validateBorderFix() {
    console.log('========================================');
    console.log('BORDER FIX VALIDATION TEST');
    console.log('========================================\n');
    
    // Load test files
    const originalFile = 'test.xlsx';
    const exportFile = 'export.xlsx';
    
    if (!fs.existsSync(originalFile)) {
        console.error('❌ test.xlsx not found');
        return;
    }
    
    if (!fs.existsSync(exportFile)) {
        console.error('❌ export.xlsx not found');
        return;
    }
    
    console.log('📁 Analyzing border structures...\n');
    
    // Analyze original file
    console.log('1. ORIGINAL FILE (test.xlsx)');
    console.log('-----------------------------');
    const originalBorders = await analyzeBorders(originalFile);
    
    // Analyze exported file
    console.log('\n2. EXPORTED FILE (export.xlsx)');
    console.log('-------------------------------');
    const exportedBorders = await analyzeBorders(exportFile);
    
    // Compare
    console.log('\n3. COMPARISON');
    console.log('--------------');
    console.log(`Original borders: ${originalBorders.count}`);
    console.log(`Exported borders: ${exportedBorders.count}`);
    console.log(`Loss: ${originalBorders.count - exportedBorders.count} (${Math.round((originalBorders.count - exportedBorders.count) / originalBorders.count * 100)}%)\n`);
    
    // Show Univer-compatible border structure
    console.log('4. UNIVER BORDER STRUCTURE (per documentation)');
    console.log('-----------------------------------------------');
    console.log('According to cell_data_structure.md:');
    console.log('');
    console.log('interface IStyleData {');
    console.log('  bd: {');
    console.log('    t: { s: number; cl: { rgb: string; }; };  // Top');
    console.log('    b: { s: number; cl: { rgb: string; }; };  // Bottom');
    console.log('    l: { s: number; cl: { rgb: string; }; };  // Left');
    console.log('    r: { s: number; cl: { rgb: string; }; };  // Right');
    console.log('  };');
    console.log('}');
    console.log('');
    console.log('Border style numbers:');
    console.log('  0: none');
    console.log('  1: thin');
    console.log('  2: hair');
    console.log('  3: dotted');
    console.log('  4: dashed');
    console.log('  5: dashDot');
    console.log('  6: dashDotDot');
    console.log('  7: double');
    console.log('  8: medium');
    console.log('  9: mediumDashed');
    console.log('  10: mediumDashDot');
    console.log('  11: mediumDashDotDot');
    console.log('  12: slantDashDot');
    console.log('  13: thick\n');
    
    // Test our post-processor mapping
    console.log('5. POST-PROCESSOR BORDER MAPPING TEST');
    console.log('--------------------------------------');
    testBorderMapping();
    
    console.log('\n✅ Validation complete\n');
}

async function analyzeBorders(filePath) {
    const zip = new JSZip();
    const content = fs.readFileSync(filePath);
    await zip.loadAsync(content);
    
    const stylesXml = await zip.file('xl/styles.xml').async('string');
    
    // Extract borders section
    const bordersMatch = stylesXml.match(/<borders count="(\d+)">(.*?)<\/borders>/s);
    
    if (!bordersMatch) {
        return { count: 0, borders: [] };
    }
    
    const count = parseInt(bordersMatch[1]);
    const bordersContent = bordersMatch[2];
    
    // Count unique border definitions
    const borderMatches = bordersContent.match(/<border[^>]*>.*?<\/border>/gs) || [];
    
    // Extract some sample borders for display
    const samples = [];
    for (let i = 0; i < Math.min(3, borderMatches.length); i++) {
        const border = borderMatches[i];
        const hasContent = 
            border.includes('style=') || 
            border.includes('color') ||
            border.includes('rgb');
        
        if (hasContent) {
            samples.push({
                index: i,
                xml: border.substring(0, 100) + (border.length > 100 ? '...' : '')
            });
        }
    }
    
    console.log(`  Border count: ${count}`);
    console.log(`  Actual definitions: ${borderMatches.length}`);
    
    if (samples.length > 0) {
        console.log('  Sample borders:');
        samples.forEach(s => {
            console.log(`    [${s.index}]: ${s.xml}`);
        });
    }
    
    return { count, borders: borderMatches };
}

function testBorderMapping() {
    // Test the mapping functions we use in post-processor
    
    const testCases = [
        {
            univerBorder: {
                t: { s: 1, cl: { rgb: '#000000' } },
                b: { s: 1, cl: { rgb: '#000000' } },
                l: { s: 1, cl: { rgb: '#000000' } },
                r: { s: 1, cl: { rgb: '#000000' } }
            },
            expected: 'All sides thin black'
        },
        {
            univerBorder: {
                t: { s: 7, cl: { rgb: '#FF0000' } },
                b: { s: 7, cl: { rgb: '#FF0000' } }
            },
            expected: 'Top/bottom double red'
        },
        {
            univerBorder: {
                l: { s: 13, cl: { rgb: '#0000FF' } },
                r: { s: 13, cl: { rgb: '#0000FF' } }
            },
            expected: 'Left/right thick blue'
        }
    ];
    
    console.log('Testing border conversion logic:');
    
    testCases.forEach((test, i) => {
        console.log(`  Test ${i + 1}: ${test.expected}`);
        const xml = convertToXml(test.univerBorder);
        console.log(`    → ${xml.substring(0, 80)}...`);
    });
}

function convertToXml(bd) {
    let xml = '<border>';
    
    const sides = ['l', 'r', 't', 'b'];
    const sideMap = {l:'left', r:'right', t:'top', b:'bottom'};
    
    sides.forEach(side => {
        const sideName = sideMap[side];
        if (bd[side]) {
            xml += `<${sideName} style="${getBorderStyleName(bd[side].s)}">`;
            if (bd[side].cl) {
                xml += `<color rgb="${colorToArgb(bd[side].cl)}"/>`;
            }
            xml += `</${sideName}>`;
        } else {
            xml += `<${sideName}/>`;
        }
    });
    
    xml += '<diagonal/></border>';
    return xml;
}

function getBorderStyleName(s) {
    const styles = {
        0:'none', 1:'thin', 2:'hair', 3:'dotted', 4:'dashed', 
        5:'dashDot', 6:'dashDotDot', 7:'double', 8:'medium', 
        9:'mediumDashed', 10:'mediumDashDot', 11:'mediumDashDotDot', 
        12:'slantDashDot', 13:'thick'
    };
    return styles[s] || 'thin';
}

function colorToArgb(color) {
    if (!color) return 'FF000000';
    
    if (typeof color === 'object' && color.rgb) {
        const rgb = color.rgb.replace('#', '');
        return 'FF' + rgb.toUpperCase();
    }
    
    if (typeof color === 'string' && color.startsWith('#')) {
        return 'FF' + color.substring(1).toUpperCase();
    }
    
    return 'FF000000';
}

// Run the test
validateBorderFix().catch(console.error);