// Test the XML parsing directly
const xmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Revenue" sheetId="2" r:id="rId2"/>
    <sheet name="Financial Model>>>" sheetId="3" r:id="rId3"/>
    <sheet name="DCF>>>" sheetId="4" r:id="rId4"/>
    <sheet name="LBO>>>" sheetId="5" r:id="rId5"/>
</sheets>`;

// Test the regex patterns
const originalRegex = /<sheet [^>]+?[^/]>[\s\S]*?<\/sheet>|<sheet [^>]+?\/>/g;
const newRegex = /<sheet\s+(?:[^>"']|"[^"]*"|'[^']*')*?(?:>[\s\S]*?<\/sheet>|\/>)/g;

console.log('Testing with original regex:');
const originalMatches = xmlContent.match(originalRegex);
console.log('Matches:', originalMatches ? originalMatches.length : 0);
if (originalMatches) {
    originalMatches.forEach(m => {
        const nameMatch = m.match(/name="([^"]*)"/);
        if (nameMatch) {
            console.log('  -', nameMatch[1]);
        }
    });
}

console.log('\nTesting with new regex:');
const newMatches = xmlContent.match(newRegex);
console.log('Matches:', newMatches ? newMatches.length : 0);
if (newMatches) {
    newMatches.forEach(m => {
        const nameMatch = m.match(/name="([^"]*)"/);
        if (nameMatch) {
            console.log('  -', nameMatch[1]);
        }
    });
}

// Test attribute parsing
console.log('\nTesting attribute parsing:');
const attrRegex = /[a-zA-Z0-9_:]+="[^"]*"/g;
const testTag = '<sheet name="DCF>>>" sheetId="5" r:id="rId5"/>';
const attrs = testTag.match(attrRegex);
console.log('Attributes found:', attrs);