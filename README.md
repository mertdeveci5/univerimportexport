# @mertdeveci55/univer-import-export

A robust Excel/CSV import and export library for [Univer](https://github.com/dream-num/univer) spreadsheets with full format preservation, including formulas, styling, charts, and conditional formatting.

[![npm version](https://img.shields.io/npm/v/@mertdeveci55/univer-import-export.svg)](https://www.npmjs.com/package/@mertdeveci55/univer-import-export)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

✅ **Import Support**
- Excel files (.xlsx, .xls)
- CSV files (.csv)
- Preserves ALL sheets (including empty ones)
- Handles special characters in sheet names (>>>, etc.)
- Maintains exact sheet order
- Full styling preservation (fonts, colors, borders, alignment)
- Formula and calculated value retention (including TRANSPOSE and array formulas)
- Merged cells support
- Images and charts
- Conditional formatting
- Data validation
- Hyperlinks and rich text

✅ **Export Support**
- Excel files (.xlsx)
- CSV files (.csv)
- Full formatting preservation
- Formula export
- Named ranges
- Multiple sheets

## Installation

```bash
npm install @mertdeveci55/univer-import-export
```

or

```bash
yarn add @mertdeveci55/univer-import-export
```

## Usage

### Import Excel to Univer

```javascript
import { LuckyExcel } from '@mertdeveci55/univer-import-export';

// Handle file input
const fileInput = document.getElementById('file-input');
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    
    LuckyExcel.transformExcelToUniver(
        file,
        (univerData) => {
            // Use the Univer data
            console.log('Converted data:', univerData);
            
            // Create Univer instance with the data
            univer.createUnit(UniverInstanceType.UNIVER_SHEET, univerData);
        },
        (error) => {
            console.error('Import error:', error);
        }
    );
});
```

### Import CSV to Univer

```javascript
import { LuckyExcel } from '@mertdeveci55/univer-import-export';

LuckyExcel.transformCsvToUniver(
    csvFile,
    (univerData) => {
        // Use the converted CSV data
        univer.createUnit(UniverInstanceType.UNIVER_SHEET, univerData);
    },
    (error) => {
        console.error('CSV import error:', error);
    }
);
```

### Export Univer to Excel

```javascript
import { LuckyExcel } from '@mertdeveci55/univer-import-export';

// Get Univer snapshot
const snapshot = univer.getActiveWorkbook().save();

LuckyExcel.transformUniverToExcel({
    snapshot: snapshot,
    fileName: 'my-spreadsheet.xlsx',
    success: () => {
        console.log('Export successful');
    },
    error: (err) => {
        console.error('Export error:', err);
    }
});
```

### Export Univer to CSV

```javascript
import { LuckyExcel } from '@mertdeveci55/univer-import-export';

const snapshot = univer.getActiveWorkbook().save();

LuckyExcel.transformUniverToCsv({
    snapshot: snapshot,
    fileName: 'my-data.csv',
    sheetName: 'Sheet1', // Optional: specific sheet to export
    success: () => {
        console.log('CSV export successful');
    },
    error: (err) => {
        console.error('CSV export error:', err);
    }
});
```

## API Reference

### `LuckyExcel.transformExcelToUniver(file, callback, errorHandler)`

Converts Excel file to Univer format.

- **file**: `File` - The Excel file (.xlsx or .xls)
- **callback**: `(data: IWorkbookData) => void` - Success callback with converted data
- **errorHandler**: `(error: Error) => void` - Error callback

### `LuckyExcel.transformCsvToUniver(file, callback, errorHandler)`

Converts CSV file to Univer format.

- **file**: `File` - The CSV file
- **callback**: `(data: IWorkbookData) => void` - Success callback
- **errorHandler**: `(error: Error) => void` - Error callback

### `LuckyExcel.transformUniverToExcel(params)`

Exports Univer data to Excel file.

**Parameters object:**
- **snapshot**: `any` - Univer workbook snapshot
- **fileName**: `string` - Output filename (optional, default: `excel_[timestamp].xlsx`)
- **getBuffer**: `boolean` - Return buffer instead of downloading (optional, default: false)
- **success**: `(buffer?: Buffer) => void` - Success callback
- **error**: `(err: Error) => void` - Error callback

### `LuckyExcel.transformUniverToCsv(params)`

Exports Univer data to CSV file.

**Parameters object:**
- **snapshot**: `any` - Univer workbook snapshot
- **fileName**: `string` - Output filename (optional, default: `csv_[timestamp].csv`)
- **sheetName**: `string` - Specific sheet to export (optional, exports all if not specified)
- **getBuffer**: `boolean` - Return content instead of downloading (optional, default: false)
- **success**: `(content?: string) => void` - Success callback
- **error**: `(err: Error) => void` - Error callback

## Browser Support

The library works in all modern browsers that support:
- ES6+
- File API
- Blob API

## Project Structure

```
src/
├── ToLuckySheet/         # Excel → LuckySheet conversion
│   ├── LuckyFile.ts      # Main file orchestrator - handles sheet discovery
│   ├── LuckySheet.ts     # Individual sheet processor
│   ├── ReadXml.ts        # XML parser with special character escaping
│   ├── LuckyCell.ts      # Cell data processor
│   └── ...
├── LuckyToUniver/        # LuckySheet → Univer conversion
│   ├── UniverWorkBook.ts # Workbook structure converter
│   ├── UniverSheet.ts    # Sheet data converter
│   └── ...
├── UniverToExcel/        # Univer → Excel export
│   ├── Workbook.ts       # Excel workbook builder using ExcelJS
│   └── ...
├── HandleZip.ts          # ZIP file operations using JSZip
└── main.ts               # Entry point with public API

dist/                     # Built output (ESM, CJS, UMD formats)
publish.sh                # Automated publishing script
gulpfile.js              # Build configuration
CLAUDE.md                # Detailed technical documentation
```

## Key Implementation Details

### Special Character Handling
Sheet names with special characters (like `>>>`) are handled through an escape/unescape mechanism in `src/ToLuckySheet/ReadXml.ts`:

```javascript
// Escapes ">" to "__GT__" before XML parsing
escapeXmlAttributes(xmlString)
// Restores "__GT__" back to ">" after parsing
unescapeXmlAttributes(xmlString)
```

### Empty Sheet Preservation
All sheets are preserved during import, even if completely empty. This maintains Excel file structure integrity.

### Formula Support
Comprehensive formula support including:
- Standard formulas (SUM, AVERAGE, IF, VLOOKUP, etc.)
- Array formulas
- TRANSPOSE formulas with proper array handling
- Shared formulas
- Named range references

## Development

### Building

```bash
npm install
npm run build   # Uses gulp to compile TypeScript and bundle
```

### Publishing

Always use the publish script for releases:

```bash
./publish.sh
```

This script:
1. Builds the project
2. Increments version
3. Commits changes
4. Pushes to GitHub
5. Publishes to npm

### Testing

```bash
npm test
```

## Key Improvements (v0.1.24)

1. **Special Character Support**: Handles sheet names with `>>>` and other special characters via escape/unescape mechanism
2. **Empty Sheet Preservation**: Empty sheets are never skipped during import
3. **No Hardcoded Solutions**: Removed all hardcoded sheet additions - all solutions are generic
4. **Sheet Order**: Maintains exact sheet order from original file
5. **Style Preservation**: Complete style mapping including bold, italic, colors, borders
6. **Formula Handling**: Preserves both formulas and calculated values, including TRANSPOSE
7. **XLS Support**: Automatic conversion of .xls files to .xlsx format
8. **Better Error Handling**: Comprehensive error messages and detailed logging

## Dependencies

### Core Dependencies
- [`@progress/jszip-esm`](https://www.npmjs.com/package/@progress/jszip-esm) - ZIP file handling for Excel files
- [`@zwight/exceljs`](https://www.npmjs.com/package/@zwight/exceljs) - Excel file structure (export)
- [`@univerjs/core`](https://github.com/dream-num/univer) - Univer core types and interfaces
- [`dayjs`](https://day.js.org/) - Date manipulation for Excel dates
- [`papaparse`](https://www.papaparse.com/) - CSV parsing
- [`xlsx`](https://sheetjs.com/) - Additional Excel format support

### Build Dependencies
- `gulp` - Build orchestration
- `rollup` - Module bundling
- `typescript` - Type safety
- `terser` - Minification (configured to preserve console.logs)

## Related Projects & References

### Core Dependencies
- **[Univer](https://github.com/dream-num/univer)** - The spreadsheet engine this library supports
- **[LuckySheet](https://github.com/mengshukeji/Luckysheet)** - Intermediate format inspiration
- **[LuckyExcel](https://github.com/dream-num/Luckyexcel)** - Original codebase this fork is based on

### Implementation Examples
- **[alphafrontend](https://github.com/mertdeveci/alphafrontend)** - Production implementation
  - See: `src/utils/excel-import.ts` for usage example
  - See: `src/pages/SpreadsheetsPage.tsx` for UI integration

### Documentation
- **[CLAUDE.md](./CLAUDE.md)** - Detailed technical documentation for AI assistants
- **[publish.sh](./publish.sh)** - Automated publishing script
- **[gulpfile.js](./gulpfile.js)** - Build configuration

## Known Issues & Solutions

### ✅ Resolved Issues

| Issue | Solution | Version Fixed |
|-------|----------|---------------|
| Sheets with special characters (>>>) not importing | Escape/unescape mechanism in ReadXml.ts | v0.1.23+ |
| AttributeList undefined errors | Defensive initialization | v0.1.21+ |
| Duplicate sheets appearing | Removed hardcoded sheet additions | v0.1.24 |
| TRANSPOSE formulas not working | Array formula support | v0.1.18+ |
| Border styles not importing | Added style collection in UniverWorkBook | v0.1.38 |

### ⚠️ Current Export Limitations (ExcelJS Library Issues)

Due to limitations in the underlying ExcelJS library, the following features have known issues during **export**:

| Issue | Root Cause | Impact | Workaround |
|-------|-----------|---------|------------|
| **Defined Names Missing** | ExcelJS `definedNames.add()` API is broken | Named ranges don't work in exported Excel files | Backend post-processing recommended |
| **Array Formula Attributes** | Missing `t="array"` and `ref="range"` XML attributes | TRANSPOSE and other array formulas may not spill correctly | Use `fillFormula()` (adds @ symbols) or backend fix |

**Import functionality works perfectly** - these limitations only affect export operations.

#### Recommended Solutions
1. **Backend Post-Processing**: Use Python/openpyxl to fix Excel files after ExcelJS export
2. **Client-Side XML Manipulation**: Direct ZIP/XML modification (performance overhead)
3. **Alternative Library**: Consider replacing ExcelJS in future versions

> 📋 **Note**: We're actively working on backend integration to resolve these export limitations while maintaining all current functionality.

## Contributing

Contributions are welcome! Please ensure:

1. **No hardcoded solutions** - All fixes must be generic
2. **Extensive logging** - Add console.log for debugging
3. **Use publish.sh** - Never manually publish to npm
4. **Test edge cases** - Including special characters, empty sheets
5. **Follow existing patterns** - Check CLAUDE.md for architecture

## License

MIT © mertdeveci

## Credits

- Original [LuckyExcel](https://github.com/dream-num/Luckyexcel) by DreamNum
- [Univer](https://github.com/dream-num/univer) spreadsheet engine
- All contributors and issue reporters

## Support

For issues and feature requests:
- [GitHub Issues](https://github.com/mertdeveci/univerjs-import-export/issues)
- Check [CLAUDE.md](./CLAUDE.md) for technical details
- Review closed issues for solutions