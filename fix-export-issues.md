# Fix for Remaining Export Issues

## Issue 1: Defined Names Not Exported

The code in `src/UniverToExcel/Workbook.ts` IS trying to add defined names (lines 28-91), but ExcelJS might not be writing them to the file properly.

**Potential issues:**
1. ExcelJS's `definedNames.add()` might not work correctly
2. The names might be added after the workbook is already serialized

## Issue 2: Array Formula Attributes Missing

The code in `src/UniverToExcel/ArrayFormulaHandler.ts` (lines 157-162) is deliberately NOT using array formula syntax to avoid @ symbols:

```typescript
// Set as a regular formula, not array formula
// This avoids ExcelJS adding @ symbols
masterCell.value = {
    formula: finalFormula,
    result: startCellValue
};
```

But this means it's missing the `t="array"` and `ref` attributes that Excel needs.

## The Problem

ExcelJS doesn't support modern Excel dynamic array formulas properly. When you try to use its array formula methods, it adds @ symbols. When you don't use them, you miss the required XML attributes.

## Solutions

### Option 1: Fix in ArrayFormulaHandler (Quick)

Change lines 157-162 in `ArrayFormulaHandler.ts` to use ExcelJS's array formula method properly:

```typescript
// Use fillFormula for array formulas (required for proper XML attributes)
worksheet.fillFormula(
    rangeStr,  // e.g., "U83:W83"
    finalFormula,  // e.g., "TRANSPOSE($N$43:$N$45)"
    startCellValue  // Initial value
);
```

### Option 2: Post-Processing (Already Works)

Since the post-processor in the HTML files already fixes both issues, we could integrate it into the library.

### Option 3: Replace ExcelJS

Consider using `xlsx` library which has better support for modern Excel features.

## Recommended Fix

Since ExcelJS has limitations, the best approach is to:

1. Let ExcelJS export as it does now
2. Add a post-processing step that fixes the XML

This is what the HTML test files already do successfully.