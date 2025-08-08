# Univer Import/Export Library - Export Implementation Plan

## Executive Summary

This plan outlines the complete redesign of the export functionality to mirror the proven import process. Instead of using ExcelJS, we will reverse-engineer our existing import pipeline to create a robust, feature-complete export solution.

**Core Principle:** If we can import it, we can export it by reversing the process.

## Current State Analysis

### Import Pipeline (Working Well)
```
Excel File (.xlsx) 
    ↓ [HandleZip.ts - JSZip]
XML Files (workbook.xml, worksheets/*.xml, etc.)
    ↓ [ReadXml.ts - Custom Parser with Special Character Handling]
Parsed XML Elements
    ↓ [LuckyFile.ts/LuckySheet.ts - Data Extraction]
LuckySheet Format (Intermediate Representation)
    ↓ [UniverWorkBook.ts/UniverSheet.ts - Format Conversion]
Univer Format (IWorkbookData)
```

### Export Pipeline (Current - Problematic)
```
Univer Format → ExcelJS → Excel File
```
**Issues:**
- Different library (ExcelJS vs JSZip)
- No intermediate LuckySheet format
- Feature gaps (array formulas, special characters)
- Inconsistent data flow

### Proposed Export Pipeline (New)
```
Univer Format (IWorkbookData)
    ↓ [UniverToLuckySheet - Reverse Conversion]
LuckySheet Format (Intermediate Representation)
    ↓ [LuckySheetToXml - XML Generation]
XML Files (workbook.xml, worksheets/*.xml, etc.)
    ↓ [ExcelBuilder.ts - JSZip]
Excel File (.xlsx)
```

## Implementation Phases

### Phase 1: UniverToLuckySheet Conversion (Week 1)

#### 1.1 Create Folder Structure
```
src/
└── UniverToLuckySheet/
    ├── index.ts                 # Main export interface
    ├── LuckyWorkBook.ts         # IWorkbookData → ILuckyFile
    ├── LuckySheetData.ts        # IWorksheetData → IluckySheet
    ├── LuckyCell.ts             # ICellData → IluckySheetCelldata
    ├── LuckyFormula.ts          # Formula handling (arrays, shared)
    ├── LuckyStyle.ts            # IStyleData → LuckySheet styles
    └── LuckyResources.ts        # Resources → charts, images, etc.
```

#### 1.2 LuckyWorkBook.ts - Core Conversion
```typescript
interface ConversionTasks {
    // Basic workbook info
    - Extract workbook name, appVersion
    - Map locale to LuckySheet format
    - Convert sheet order
    
    // Sheet conversion
    - Iterate through sheets
    - Preserve sheet order
    - Handle empty sheets
    
    // Resources
    - Extract defined names
    - Extract hyperlinks
    - Extract images/charts
    - Extract conditional formatting
    - Extract data validation
    - Extract filters
}
```

#### 1.3 LuckySheetData.ts - Sheet Level Conversion
```typescript
interface SheetConversionTasks {
    // Basic properties
    - name, index, order
    - tabColor → color
    - hidden → hide
    - zoomRatio
    - showGridlines → showGridLines
    - defaultColumnWidth → defaultColWidth
    - defaultRowHeight → defaultRowHeight
    
    // Cell data
    - Convert cellData matrix to array format
    - Preserve all formulas
    - Maintain cell references
    
    // Config object
    - mergeData → config.merge
    - rowData → config.rowlen/rowhidden
    - columnData → config.columnlen/colhidden
    - freeze → freezen
    
    // Special handling
    - Array formulas (TRANSPOSE)
    - Shared formulas
    - Rich text cells
}
```

#### 1.4 LuckyCell.ts - Cell Level Conversion
```typescript
interface CellConversionTasks {
    // Position
    - Extract row/column from matrix position
    
    // Value handling
    - Simple values (string, number, boolean)
    - Formula preservation (f, si fields)
    - Array formula detection (check for range)
    
    // Style conversion
    - Font properties (bl, it, fs, ff, cl)
    - Alignment (ht, vt, tb, tr)
    - Background (bg)
    - Borders (complex mapping)
    - Number format (ct)
    
    // Special cases
    - Rich text (p field)
    - Hyperlinks
    - Cell images
    - Comments/notes
}
```

#### 1.5 LuckyFormula.ts - Formula Specialization
```typescript
interface FormulaHandlingTasks {
    // Array formulas
    - Detect array formulas from Univer
    - Restore ref property for TRANSPOSE
    - Set ft: "array" flag
    
    // Shared formulas
    - Identify shared formula patterns
    - Restore si and ref properties
    
    // Formula cleaning
    - Ensure proper = prefix
    - Handle special functions
    - Preserve cell references
}
```

### Phase 2: LuckySheetToXml Generation (Week 2)

#### 2.1 Create XML Generation Infrastructure
```
src/
└── LuckySheetToXml/
    ├── index.ts                 # Main XML generation interface
    ├── XmlBase.ts               # Base XML utilities
    ├── ContentTypesXml.ts       # [Content_Types].xml
    ├── WorkbookXml.ts           # xl/workbook.xml
    ├── WorksheetXml.ts          # xl/worksheets/sheet{n}.xml
    ├── SharedStringsXml.ts      # xl/sharedStrings.xml
    ├── StylesXml.ts             # xl/styles.xml
    ├── RelationshipsXml.ts      # xl/_rels/*.rels
    ├── ThemeXml.ts              # xl/theme/theme1.xml
    ├── CalcChainXml.ts          # xl/calcChain.xml
    └── DrawingXml.ts            # xl/drawings/*.xml
```

#### 2.2 XmlBase.ts - Core XML Utilities
```typescript
class XmlBase {
    // Character escaping (reuse from ReadXml.ts)
    escapeXmlContent(content: string): string
    escapeXmlAttribute(attr: string): string
    escapeSheetName(name: string): string  // Handle ">>>"
    
    // Element creation
    createElement(tag: string, attrs?: {}, content?: string): string
    createSelfClosingElement(tag: string, attrs?: {}): string
    
    // Document creation
    createXmlDocument(root: string, xmlns: {}): string
    
    // Namespace handling
    addNamespace(tag: string, namespace: string): string
}
```

#### 2.3 WorkbookXml.ts - Workbook Structure
```typescript
interface WorkbookXmlTasks {
    // XML structure
    - Create workbook root with namespaces
    - Generate fileVersion element
    - Generate workbookPr element
    
    // Sheets
    - Create sheets element
    - For each sheet:
        - name attribute (with escaping!)
        - sheetId attribute
        - r:id relationship
        - state (hidden/visible)
    
    // Other elements
    - definedNames
    - calcPr
    - workbookProtection
}
```

#### 2.4 WorksheetXml.ts - Sheet Data (Most Complex)
```typescript
interface WorksheetXmlTasks {
    // Document structure
    - Create worksheet root with namespaces
    - sheetPr (tab color, etc.)
    - dimension (data range)
    - sheetViews (zoom, selection, freeze)
    - sheetFormatPr (default sizes)
    - cols (column widths)
    
    // Sheet data (critical!)
    - Create sheetData element
    - Group cells by row
    - For each row:
        - r attribute (row number)
        - ht, hidden, customHeight
        - For each cell:
            - r attribute (A1 notation)
            - s attribute (style index)
            - t attribute (type)
            - Handle formulas:
                * Regular: <f>formula</f>
                * Array: <f t="array" ref="A1:C3">formula</f>
                * Shared: <f t="shared" si="0" ref="A1:A10">formula</f>
            - Handle values:
                * Direct: <v>value</v>
                * Shared string: <v>index</v>
                * Inline string: <is><t>text</t></is>
    
    // Additional elements
    - mergeCells
    - conditionalFormatting
    - dataValidations
    - hyperlinks
    - drawing relationships
}
```

#### 2.5 StylesXml.ts - Styling Information
```typescript
interface StylesXmlTasks {
    // Style components
    - numFmts (number formats)
    - fonts collection
    - fills collection
    - borders collection
    - cellStyleXfs
    - cellXfs (main styles)
    - cellStyles
    - dxfs (conditional formatting styles)
    
    // Style indexing
    - Build style registry
    - Deduplicate styles
    - Assign indices
}
```

#### 2.6 SharedStringsXml.ts - String Optimization
```typescript
interface SharedStringsXmlTasks {
    // String collection
    - Collect all unique strings
    - Build string index map
    - Generate sst element
    - For each unique string:
        - Create si element
        - Handle rich text (multiple r elements)
        - Handle plain text (t element)
    
    // Update cell references
    - Replace string values with indices
    - Set cell type to shared string
}
```

### Phase 3: Excel File Assembly (Week 3)

#### 3.1 Create Excel Builder
```
src/
└── LuckySheetToExcel/
    ├── ExcelBuilder.ts          # Main orchestrator
    ├── FileStructure.ts         # Excel file structure
    └── ZipGenerator.ts          # JSZip wrapper
```

#### 3.2 ExcelBuilder.ts - Main Orchestration
```typescript
interface BuilderTasks {
    // Preparation
    - Convert Univer to LuckySheet
    - Initialize XML generators
    - Prepare file structure
    
    // XML Generation
    - Generate all XML files
    - Validate XML structure
    - Handle relationships
    
    // File Assembly
    - Create ZIP structure
    - Add all files in correct paths
    - Generate final buffer
    
    // Error handling
    - Validate required files
    - Check for missing relationships
    - Ensure valid Excel structure
}
```

#### 3.3 FileStructure.ts - Excel ZIP Structure
```typescript
interface ExcelStructure {
    root: {
        "[Content_Types].xml": string,
        "_rels/": {
            ".rels": string
        },
        "docProps/": {
            "app.xml": string,
            "core.xml": string
        },
        "xl/": {
            "workbook.xml": string,
            "styles.xml": string,
            "sharedStrings.xml": string,
            "_rels/": {
                "workbook.xml.rels": string
            },
            "worksheets/": {
                "sheet1.xml": string,
                "sheet2.xml": string,
                // ... more sheets
                "_rels/": {
                    "sheet1.xml.rels": string,
                    // ... relationships
                }
            },
            "theme/": {
                "theme1.xml": string
            },
            "drawings/": {
                "drawing1.xml": string,
                // ... more drawings
            },
            "charts/": {
                "chart1.xml": string,
                // ... more charts
            },
            "media/": {
                "image1.png": Buffer,
                // ... more images
            }
        }
    }
}
```

### Phase 4: Integration & Testing (Week 4)

#### 4.1 Update Main API
```typescript
// main.ts modifications
interface ApiChanges {
    // New internal method
    transformUniverToLuckySheet(snapshot: any): ILuckyFile
    
    // New internal method
    transformLuckySheetToExcel(luckyFile: ILuckyFile): Promise<ArrayBuffer>
    
    // Updated public method
    transformUniverToExcel(params: ExportParams): Promise<void>
    
    // Deprecate but keep for compatibility
    transformUniverToExcelLegacy(params: ExportParams): Promise<void>
}
```

#### 4.2 Testing Strategy

##### Unit Tests
```typescript
describe('UniverToLuckySheet', () => {
    test('converts basic cell values')
    test('preserves formulas')
    test('handles array formulas (TRANSPOSE)')
    test('converts styles correctly')
    test('preserves empty sheets')
    test('handles special characters in sheet names')
})

describe('LuckySheetToXml', () => {
    test('generates valid XML')
    test('escapes special characters')
    test('creates proper formula elements')
    test('handles merged cells')
    test('generates correct relationships')
})

describe('ExcelBuilder', () => {
    test('creates valid ZIP structure')
    test('includes all required files')
    test('generates downloadable Excel file')
})
```

##### Integration Tests
```typescript
describe('Round-trip Testing', () => {
    test('Import → Export → Import produces same result')
    test('Complex formulas survive round-trip')
    test('Styling is preserved')
    test('Charts and images are maintained')
    test('Special characters handled correctly')
})
```

##### Edge Cases
```typescript
describe('Edge Cases', () => {
    test('Empty workbook')
    test('Single empty sheet')
    test('Sheet name with ">>>"')
    test('10000+ cells performance')
    test('Complex nested formulas')
    test('Maximum style combinations')
})
```

## Implementation Details

### Critical Success Factors

1. **Character Escaping**
   - Must handle: `< > & " ' >>>` in all contexts
   - Sheet names need special handling
   - Formula content needs different escaping than values

2. **Formula Preservation**
   - Array formulas: Must preserve `ref` attribute
   - Shared formulas: Must maintain `si` and cell references
   - Named ranges: Must resolve correctly

3. **Style Deduplication**
   - Excel requires styles to be deduplicated
   - Each unique combination gets an index
   - Cells reference style by index

4. **Relationship Management**
   - Every reference needs a relationship
   - Relationships must have unique IDs
   - Must maintain parent-child relationships

5. **Empty Sheet Handling**
   - Even empty sheets need worksheet XML
   - Must have minimum structure
   - Default dimensions and properties

### Data Structures

#### Style Registry
```typescript
class StyleRegistry {
    private fonts: Map<string, number> = new Map()
    private fills: Map<string, number> = new Map()
    private borders: Map<string, number> = new Map()
    private cellXfs: Map<string, number> = new Map()
    
    registerFont(font: FontStyle): number
    registerFill(fill: FillStyle): number
    registerBorder(border: BorderStyle): number
    getCellStyleIndex(style: CellStyle): number
}
```

#### String Registry
```typescript
class SharedStringRegistry {
    private strings: Map<string, number> = new Map()
    private count: number = 0
    
    addString(str: string): number
    getStrings(): string[]
    getCount(): number
    getUniqueCount(): number
}
```

#### Relationship Manager
```typescript
class RelationshipManager {
    private relationships: Map<string, Relationship[]> = new Map()
    private nextId: number = 1
    
    addRelationship(parent: string, type: string, target: string): string
    getRelationships(parent: string): Relationship[]
    generateRelationshipXml(parent: string): string
}
```

### Performance Considerations

1. **Memory Management**
   - Stream large files if possible
   - Clear intermediate data after use
   - Use string builders for XML

2. **Optimization Points**
   - Style deduplication (major impact)
   - Shared string optimization
   - Batch cell processing by row

3. **Benchmarks**
   - Target: 10,000 cells < 1 second
   - Target: 100,000 cells < 10 seconds
   - Memory: < 100MB for typical files

## Migration Plan

### Version 0.1.26 - Foundation
- Implement UniverToLuckySheet
- Basic XML generation (workbook, worksheet)
- Simple cell values and formulas

### Version 0.1.27 - Styles & Formatting
- Complete StylesXml implementation
- SharedStrings optimization
- Merged cells, borders, fills

### Version 0.1.28 - Advanced Features
- Array formulas (TRANSPOSE)
- Conditional formatting
- Data validation
- Hyperlinks

### Version 0.1.29 - Media & Charts
- Image export
- Chart export
- Drawing relationships

### Version 0.2.0 - Complete Parity
- All features from import
- Performance optimizations
- Deprecate ExcelJS approach

## Risk Mitigation

### Technical Risks
1. **XML Complexity**
   - Mitigation: Study Excel XML schemas
   - Fallback: Validate against Excel specs

2. **Performance Issues**
   - Mitigation: Profile and optimize hotspots
   - Fallback: Streaming for large files

3. **Compatibility**
   - Mitigation: Test with multiple Excel versions
   - Fallback: Target Excel 2016+ minimum

### Schedule Risks
1. **Scope Creep**
   - Mitigation: Strict phase boundaries
   - Fallback: Ship core features first

2. **Testing Time**
   - Mitigation: Automated test suite
   - Fallback: Beta release for testing

## Success Metrics

1. **Feature Parity**
   - 100% of imported features can be exported
   - Round-trip test success rate > 99%

2. **Performance**
   - Export speed within 2x of import speed
   - Memory usage < current ExcelJS approach

3. **Code Quality**
   - 80%+ code reuse from import
   - 90%+ test coverage
   - Zero critical bugs in production

## Appendix: File Mapping

### Import Process Files (To Study)
- `src/ToLuckySheet/ReadXml.ts` - XML parsing patterns
- `src/ToLuckySheet/LuckyFile.ts` - Workbook structure
- `src/ToLuckySheet/LuckySheet.ts` - Sheet processing
- `src/ToLuckySheet/LuckyCell.ts` - Cell handling
- `src/common/constant.ts` - Excel constants
- `src/common/method.ts` - Utility functions

### Export Process Files (To Create)
- `src/UniverToLuckySheet/*` - Reverse conversion
- `src/LuckySheetToXml/*` - XML generation
- `src/LuckySheetToExcel/*` - File assembly

### Shared Components
- `src/common/XmlBase.ts` - New shared XML utilities
- `src/common/constant.ts` - Existing Excel constants
- `src/HandleZip.ts` - Existing ZIP handling

## Next Steps

1. **Review & Approve Plan**
   - Technical review
   - Timeline approval
   - Resource allocation

2. **Setup Development**
   - Create folder structure
   - Setup test framework
   - Create development branch

3. **Begin Phase 1**
   - Start with UniverToLuckySheet
   - Focus on data integrity
   - Build test suite alongside

---

**Document Version:** 1.0  
**Last Updated:** 2025-08-08  
**Author:** Development Team  
**Status:** DRAFT - Awaiting Approval