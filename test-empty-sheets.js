const fs = require('fs');
const path = require('path');

// Read the built library
const { LuckyFile } = require('./lib/ToLuckySheet/LuckyFile.js');

// Create a mock file structure for testing
const mockFiles = {
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
        <sheet name="EmptySheet" sheetId="2" r:id="rId2"/>
        <sheet name="Sheet3" sheetId="3" r:id="rId3"/>
    </sheets>
</workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="str"><v>Hello</v></c>
        </row>
    </sheetData>
</worksheet>`,
    // sheet2.xml is intentionally missing to simulate empty sheet
    'xl/worksheets/sheet3.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="str"><v>World</v></c>
        </row>
    </sheetData>
</worksheet>`,
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
    'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellXfs>
</styleSheet>`,
    'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">
</sst>`,
    'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
    <Application>Test</Application>
    <AppVersion>1.0</AppVersion>
</Properties>`,
    'docProps/core.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
    <dc:creator>Test</dc:creator>
    <cp:lastModifiedBy>Test</cp:lastModifiedBy>
    <dcterms:created>2024-01-01T00:00:00Z</dcterms:created>
    <dcterms:modified>2024-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`
};

// Test the import
try {
    console.log('Testing empty sheet import...\n');
    
    const luckyFile = new LuckyFile(mockFiles, 'test.xlsx');
    luckyFile.getWorkBookInfo();
    luckyFile.getSheetsFull();
    
    console.log('Number of sheets found:', luckyFile.sheets.length);
    console.log('\nSheet details:');
    luckyFile.sheets.forEach(sheet => {
        console.log(`- ${sheet.name} (index: ${sheet.index}, order: ${sheet.order})`);
        console.log(`  Has celldata: ${sheet.celldata && sheet.celldata.length > 0}`);
        console.log(`  Row count: ${sheet.row}`);
        console.log(`  Column count: ${sheet.column}`);
    });
    
    // Test conversion to JSON
    const result = JSON.parse(luckyFile.Parse());
    console.log('\nFinal output sheets:', result.sheets.length);
    console.log('Sheet names:', result.sheets.map(s => s.name));
    
} catch (error) {
    console.error('Test failed:', error);
}