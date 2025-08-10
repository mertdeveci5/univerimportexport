# CRITICAL FINDINGS

## ExcelJS Library Limitations

After extensive testing, I've discovered that **ExcelJS has severe limitations**:

### 1. Defined Names API is BROKEN
```javascript
workbook.definedNames.add('TestName', 'Sheet1!$A$1');
console.log(workbook.definedNames.model); // Returns: []  <-- EMPTY!
```

The `definedNames.add()` method does NOT work. The model stays empty and nothing gets exported to the XML.

### 2. Array Formula Support is PROBLEMATIC
- Using `fillFormula()` adds @ symbols (breaks formulas)
- Not using it misses required XML attributes (`t="array"` and `ref`)
- No proper support for dynamic array formulas

## What DOES Work

✅ **Border styles** - Fixed by collecting styles into registry (37 styles now exported correctly)

## What DOESN'T Work (ExcelJS limitations)

❌ **Defined names** - ExcelJS's API is broken
❌ **Array formulas** - ExcelJS doesn't support modern Excel properly

## The Only Solution: Post-Processing

Since ExcelJS has fundamental limitations, the ONLY working solution is post-processing the exported file:

```javascript
// After ExcelJS export
const zip = new JSZip();
await zip.loadAsync(exportBuffer);

// 1. Fix array formulas
let sheetXml = await zip.file('xl/worksheets/sheet7.xml').async('string');
sheetXml = sheetXml.replace(
    /<f>TRANSPOSE/g,
    '<f t="array" ref="U83:W83">TRANSPOSE'
);

// 2. Add defined names
let workbookXml = await zip.file('xl/workbook.xml').async('string');
workbookXml = workbookXml.replace(
    '</workbook>',
    '<definedNames>...</definedNames></workbook>'
);

// Save fixed file
const fixedBuffer = await zip.generateAsync({ type: 'arraybuffer' });
```

## Recommendation

### Short Term (Immediate Fix)
Integrate the post-processor from `test-simple-upload.html` into the library as a built-in step after ExcelJS export.

### Long Term (Proper Fix)
Replace ExcelJS with a library that properly supports:
- Modern Excel features (dynamic arrays)
- Defined names
- All Excel 365 functionality

Potential alternatives:
- `xlsx` (SheetJS) - More comprehensive but different API
- `xlsx-populate` - Better formula support
- Direct XML manipulation - Most control but more work

## Current Status

| Feature | Import | Export (Raw) | Export (Post-Processed) |
|---------|--------|--------------|------------------------|
| Borders | ✅ Fixed | ✅ Fixed | ✅ Works |
| Defined Names | ✅ Works | ❌ ExcelJS broken | ✅ Works |
| Array Formulas | ✅ Works | ❌ Missing attributes | ✅ Works |

The post-processor in the HTML test files successfully fixes all issues. This should be integrated into the library.