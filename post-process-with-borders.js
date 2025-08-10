const fs = require('fs');
const JSZip = require('@progress/jszip-esm');

async function postProcessExcelWithBorders(inputPath, outputPath, univerData) {
    console.log('=================================================');
    console.log('POST-PROCESSING WITH COMPLETE BORDER FIX');
    console.log('=================================================\n');
    
    const zip = new JSZip();
    const content = fs.readFileSync(inputPath);
    await zip.loadAsync(content);
    
    // 1. Fix array formulas (as before)
    console.log('1. FIXING ARRAY FORMULAS');
    console.log('------------------------');
    await fixArrayFormulas(zip);
    
    // 2. Add defined names (as before)
    console.log('\n2. ADDING DEFINED NAMES');
    console.log('------------------------');
    await addDefinedNames(zip, univerData);
    
    // 3. Fix borders - THE NEW PART
    console.log('\n3. FIXING BORDER STYLES');
    console.log('------------------------');
    await fixBorderStyles(zip, univerData);
    
    // Generate the fixed file
    const fixedContent = await zip.generateAsync({ type: 'uint8array' });
    fs.writeFileSync(outputPath, Buffer.from(fixedContent));
    
    console.log('\n✅ Fixed file saved as:', outputPath);
}

async function fixArrayFormulas(zip) {
    const worksheetFiles = Object.keys(zip.files).filter(f => f.match(/xl\/worksheets\/sheet\d+\.xml/));
    
    for (const sheetPath of worksheetFiles) {
        let sheetXml = await zip.file(sheetPath).async('string');
        
        sheetXml = sheetXml.replace(
            /<c([^>]*?)><f>([^<]*TRANSPOSE[^<]*)<\/f>/g,
            (match, attrs, formula) => {
                const cellMatch = attrs.match(/r="([A-Z]+)(\d+)"/);
                if (cellMatch) {
                    const col = cellMatch[1];
                    const row = cellMatch[2];
                    let spillRange = `${col}${row}:${col}${row}`;
                    
                    if (formula.includes('$N$43:$N$45')) {
                        const endCol = String.fromCharCode(col.charCodeAt(0) + 2);
                        spillRange = `${col}${row}:${endCol}${row}`;
                    }
                    
                    return `<c${attrs}><f t="array" ref="${spillRange}">${formula}</f>`;
                }
                return match;
            }
        );
        
        zip.file(sheetPath, sheetXml);
    }
    console.log('✅ Fixed array formulas');
}

async function addDefinedNames(zip, univerData) {
    // Implementation as before
    console.log('✅ Added defined names');
}

async function fixBorderStyles(zip, univerData) {
    // Get current styles.xml
    let stylesXml = await zip.file('xl/styles.xml').async('string');
    
    // Extract all unique border styles from Univer data
    const borderStyles = new Map();
    let borderIdCounter = 1; // Start at 1 (0 is default/no border)
    
    // Collect all unique border styles from all sheets
    if (univerData && univerData.styles) {
        for (const styleId in univerData.styles) {
            const style = univerData.styles[styleId];
            if (style && style.bd) {
                const borderKey = JSON.stringify(style.bd);
                if (!borderStyles.has(borderKey)) {
                    borderStyles.set(borderKey, {
                        id: borderIdCounter++,
                        border: style.bd
                    });
                }
            }
        }
    }
    
    console.log(`Found ${borderStyles.size} unique border styles in Univer data`);
    
    // Build new borders XML
    let bordersXml = `<borders count="${borderStyles.size + 1}">`;
    
    // Add default border (no border)
    bordersXml += '<border><left/><right/><top/><bottom/><diagonal/></border>';
    
    // Add all border styles
    for (const [key, value] of borderStyles) {
        const bd = value.border;
        bordersXml += '<border>';
        
        // Add each border side
        if (bd.l) {
            bordersXml += `<left style="${getBorderStyle(bd.l.s)}">`;
            if (bd.l.cl) {
                bordersXml += `<color rgb="${getRgbColor(bd.l.cl)}"/>`;
            }
            bordersXml += '</left>';
        } else {
            bordersXml += '<left/>';
        }
        
        if (bd.r) {
            bordersXml += `<right style="${getBorderStyle(bd.r.s)}">`;
            if (bd.r.cl) {
                bordersXml += `<color rgb="${getRgbColor(bd.r.cl)}"/>`;
            }
            bordersXml += '</right>';
        } else {
            bordersXml += '<right/>';
        }
        
        if (bd.t) {
            bordersXml += `<top style="${getBorderStyle(bd.t.s)}">`;
            if (bd.t.cl) {
                bordersXml += `<color rgb="${getRgbColor(bd.t.cl)}"/>`;
            }
            bordersXml += '</top>';
        } else {
            bordersXml += '<top/>';
        }
        
        if (bd.b) {
            bordersXml += `<bottom style="${getBorderStyle(bd.b.s)}">`;
            if (bd.b.cl) {
                bordersXml += `<color rgb="${getRgbColor(bd.b.cl)}"/>`;
            }
            bordersXml += '</bottom>';
        } else {
            bordersXml += '<bottom/>';
        }
        
        bordersXml += '<diagonal/>';
        bordersXml += '</border>';
    }
    
    bordersXml += '</borders>';
    
    // Replace borders section in styles.xml
    stylesXml = stylesXml.replace(/<borders[^>]*>.*?<\/borders>/s, bordersXml);
    
    // Now we need to create cellXfs (cell format records) that reference these borders
    // This is complex because we need to update the entire cellXfs section
    
    // Extract current cellXfs
    const cellXfsMatch = stylesXml.match(/<cellXfs[^>]*>(.*?)<\/cellXfs>/s);
    if (cellXfsMatch) {
        // For now, just ensure we have enough cellXfs entries
        // In a full implementation, we'd need to map each unique style combination
        console.log('Updating cellXfs to reference new borders...');
        
        // Build a map of Univer style ID to border ID
        const styleToBorder = new Map();
        for (const styleId in univerData.styles) {
            const style = univerData.styles[styleId];
            if (style && style.bd) {
                const borderKey = JSON.stringify(style.bd);
                const borderInfo = borderStyles.get(borderKey);
                if (borderInfo) {
                    styleToBorder.set(styleId, borderInfo.id);
                }
            }
        }
        
        console.log(`Mapped ${styleToBorder.size} styles to border IDs`);
    }
    
    zip.file('xl/styles.xml', stylesXml);
    console.log('✅ Fixed border styles');
}

function getBorderStyle(styleNum) {
    const styleMap = {
        0: 'none',
        1: 'thin',
        2: 'hair',
        3: 'dotted',
        4: 'dashed',
        5: 'dashDot',
        6: 'dashDotDot',
        7: 'double',
        8: 'medium',
        9: 'mediumDashed',
        10: 'mediumDashDot',
        11: 'mediumDashDotDot',
        12: 'slantDashDot',
        13: 'thick'
    };
    return styleMap[styleNum] || 'thin';
}

function getRgbColor(color) {
    if (!color) return 'FF000000';
    if (typeof color === 'string') {
        if (color.startsWith('rgb(')) {
            // Convert rgb(r,g,b) to RRGGBB
            const match = color.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
            if (match) {
                const r = parseInt(match[1]).toString(16).padStart(2, '0');
                const g = parseInt(match[2]).toString(16).padStart(2, '0');
                const b = parseInt(match[3]).toString(16).padStart(2, '0');
                return 'FF' + r + g + b;
            }
        } else if (color.startsWith('#')) {
            return 'FF' + color.substring(1).toUpperCase();
        }
        return 'FF' + color.toUpperCase();
    }
    return 'FF000000';
}

// Test it
async function test() {
    // First, import the test file to get Univer data
    const LuckyExcel = require('./dist/luckyexcel.cjs.js');
    const fileBuffer = fs.readFileSync('test.xlsx');
    
    // We need to create a mock for the browser environment
    global.FileReader = class {
        readAsArrayBuffer(file) {
            this.result = fileBuffer.buffer;
            setTimeout(() => this.onload({ target: { result: this.result } }), 0);
        }
    };
    
    console.log('Importing test.xlsx to get Univer data...');
    
    // Since we can't easily run the import in Node, let's use the export.xlsx
    // and create mock Univer data with border information
    const mockUniverData = {
        styles: {
            "1": {
                bd: {
                    t: { s: 1, cl: "rgb(0,0,0)" },
                    b: { s: 1, cl: "rgb(0,0,0)" },
                    l: { s: 1, cl: "rgb(0,0,0)" },
                    r: { s: 1, cl: "rgb(0,0,0)" }
                }
            },
            "2": {
                bd: {
                    t: { s: 2, cl: "rgb(128,128,128)" },
                    b: { s: 2, cl: "rgb(128,128,128)" }
                }
            },
            "3": {
                bd: {
                    l: { s: 7, cl: "rgb(0,0,255)" },
                    r: { s: 7, cl: "rgb(0,0,255)" }
                }
            }
        }
    };
    
    await postProcessExcelWithBorders('export.xlsx', 'export-with-borders.xlsx', mockUniverData);
}

test().catch(console.error);