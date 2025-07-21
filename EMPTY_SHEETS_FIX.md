# Fix for Empty Sheets Not Being Imported

## Problem
When importing Excel files that contain empty sheets (sheets defined in the workbook but without corresponding worksheet files), these sheets were being filtered out during the import process.

## Root Cause
In `src/ToLuckySheet/LuckyFile.ts`, the `getSheetsFull` method was checking if `sheetFile != null` before creating a sheet object. Empty sheets in Excel don't have corresponding worksheet files, so `sheetFile` would be null for these sheets, causing them to be skipped.

## Solution
1. **Modified `LuckyFile.ts`**: Removed the conditional check that skipped sheets when `sheetFile` was null. Now all sheets defined in the workbook are processed, regardless of whether they have a corresponding worksheet file.

2. **Modified `LuckySheet.ts`**: Added a check at the beginning of the constructor to handle null `sheetFile` gracefully. When `sheetFile` is null, the sheet is initialized with minimal default values and returns early, avoiding null reference errors.

## Changes Made

### File: `src/ToLuckySheet/LuckyFile.ts`
- Line 154: Removed the `if(sheetFile!=null)` condition
- Line 155: Renamed variable from `sheet` to `luckySheet` to avoid TypeScript naming conflict
- Lines 173, 175: Updated references to use the renamed variable

### File: `src/ToLuckySheet/LuckySheet.ts`
- Lines 53-69: Added early return check for null `sheetFile` with default values:
  - Sets default grid lines, zoom ratio, column/row dimensions
  - Initializes empty celldata and calcChain arrays
  - Returns early to avoid processing non-existent worksheet data

## Result
Empty sheets are now properly imported and displayed in the Univer/Luckysheet output with default dimensions and settings, matching the behavior of standard spreadsheet applications.