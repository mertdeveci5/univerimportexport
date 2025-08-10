# Investigation Report: Excel Import/Export Issues

## Executive Summary
Three critical issues were identified in the Excel import/export pipeline:
1. **TRANSPOSE formulas get @ symbols** when opened in Excel (causing #VALUE! errors)
2. **Defined names are completely missing** (0 out of 9 imported/exported)
3. **Border styles mostly disappear** (1 out of 37 border definitions exported)

## Issue 1: TRANSPOSE Formula Corruption

### Root Cause
TRANSPOSE formulas are not being marked as array formulas in the exported XML.

### Original XML (Working)
```xml
<c r="T83" s="100" cm="1">
  <f t="array" ref="T83:V83">+TRANSPOSE($N$43:$N$45)</f>
  <v>32.878632753636502</v>
</c>
```

### Exported XML (Broken)
```xml
<c r="T83" s="75">
  <f>TRANSPOSE($N$43:$N$45)</f>
  <v>32.8786327536365</v>
</c>
```

### Missing Attributes
- `t="array"` - Tells Excel this is an array formula
- `ref="T83:V83"` - Defines the spill range

### Why This Matters
Without `t="array"`, modern Excel (365/2021) treats TRANSPOSE as a regular formula and adds @ symbols for implicit intersection, resulting in: `=@TRANSPOSE(@$N$43:$N$45)`

### Additional Problem
Spill cells (U83, V83) in the export contain gibberish formulas like `FJS3OS` instead of just values. This was partially fixed in v0.1.36 but the array formula issue remains.

## Issue 2: Defined Names Not Imported

### Investigation Results
- **Original file**: 10 defined names in workbook.xml
  - Examples: `capexswitch`, `circ`, `Costswitch`, etc.
- **After import**: 0 defined names in Univer data
- **After export**: 0 defined names in workbook.xml

### Code Analysis
```
✗ LuckyFile.ts does NOT mention definedName
✗ ReadXml.ts does NOT mention definedName
```

The import code doesn't even attempt to read defined names from the Excel file.

## Issue 3: Border Styles Disappearing

### Investigation Results
- **Original file**: 37 border definitions in styles.xml
- **Exported file**: 1 border definition in styles.xml
- **Border retention rate**: ~3%

### Code Analysis
The code has border handling functions:
- `getBorderInfo()` in style.ts
- `handleBorder()` in style.ts
- `convertBorderToExcel()` in univerUtils.ts

However, borders are not being properly preserved through the import→Univer→export pipeline.

## Data Flow Analysis

### Import Flow (Excel → LuckySheet → Univer)
1. **Defined Names**: Not read from workbook.xml
2. **TRANSPOSE**: Imported with `t="array"` but not preserved as array formula in Univer
3. **Borders**: Read but not fully converted to Univer format

### Export Flow (Univer → Excel)
1. **Defined Names**: No code to export them
2. **TRANSPOSE**: Not marked with `t="array"` and `ref` attributes
3. **Borders**: Only minimal borders exported

## File Comparison

| Metric | Original | After Import | After Export |
|--------|----------|--------------|--------------|
| Defined Names | 10 | 0 | 0 |
| DCF Formulas | 526 | 526 | 564 (+38 due to array formula handling) |
| Border Definitions | 37 | ? | 1 |
| TRANSPOSE with t="array" | Yes | No | No |

## Recommendations

### For TRANSPOSE Fix
1. When exporting array formulas, ExcelJS needs to set `t="array"` and `ref` attributes
2. The ArrayFormulaHandler needs to properly mark TRANSPOSE as array formula
3. Spill cells should only contain values, not formulas

### For Defined Names
1. Add code to read `<definedNames>` from workbook.xml during import
2. Store in Univer data structure
3. Add code to export defined names back to workbook.xml

### For Borders
1. Investigate why border styles are not being fully preserved
2. Ensure all border definitions are imported from styles.xml
3. Properly export all border styles back to styles.xml

## Test Files
- **Original**: test.xlsx (working Excel file)
- **Exported**: export.xlsx (broken when opened in Excel)
- **Version**: v0.1.37 (latest published version)

## Next Steps
These issues need to be fixed in the core import/export logic, not just for this specific file but for ALL Excel files to ensure the product works as a robust spreadsheet conversion tool.