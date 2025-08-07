# @mertdeveci55/univer-import-export

A comprehensive Excel/CSV import and export library for [Univer](https://github.com/dream-num/univer) spreadsheets with full format preservation.

## Features

✅ **Import Support**
- Excel files (.xlsx, .xls)
- CSV files (.csv)
- Preserves all sheets (including empty ones)
- Maintains exact sheet order
- Full styling preservation (fonts, colors, borders, alignment)
- Formula and calculated value retention
- Merged cells support
- Images and charts
- Conditional formatting
- Data validation

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

## Development

### Building

```bash
npm install
npm run build
```

### Testing

```bash
npm test
```

## Key Improvements in This Version

1. **Empty Sheet Preservation**: Empty sheets are no longer skipped during import
2. **Sheet Order**: Maintains exact sheet order from original file
3. **Style Preservation**: Complete style mapping including bold, italic, colors, borders
4. **Formula Handling**: Preserves both formulas and calculated values
5. **XLS Support**: Automatic conversion of .xls files to .xlsx format
6. **Better Error Handling**: Comprehensive error messages and handling

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT

## Credits

This project is based on the original [Luckyexcel](https://github.com/dream-num/Luckyexcel) and adapted specifically for Univer spreadsheets with enhanced functionality.

## Support

For issues and feature requests, please visit the [GitHub repository](https://github.com/mertdeveci/univerjs-import-export/issues).