# Export Alignment Analysis - Import/Export Symmetry

## Executive Summary

After reviewing the Univer documentation and import code, I've identified critical gaps and misalignments in the export functionality. The export MUST properly reverse-engineer the import process to ensure data integrity and feature preservation.

## Critical Findings

### 1. Data Structure Alignment

#### Univer IWorkbookData Structure (From Documentation)
```typescript
interface IWorkbookData {
    id: string;                                         // Unique identifier
    name: string;                                       // Workbook name
    appVersion: string;                                 // Version string
    locale: LocaleType;                                 // Locale setting
    styles: Record<string, Nullable<IStyleData>>;      // Style registry
    sheetOrder: string[];                               // Sheet ordering
    sheets: { [sheetId: string]: IWorksheetData };     // Sheet data
    resources?: IResources;                             // Plugin data storage
}
```

#### Univer IWorksheetData Structure
```typescript
interface IWorksheetData {
    id: string;                                         // Sheet ID
    name: string;                                       // Sheet name
    tabColor?: string;                                  // Tab color
    hidden: BooleanNumber;                              // 0 or 1
    freeze?: IFreeze;                                   // Freeze panes
    rowCount: number;                                   // Total rows
    columnCount: number;                                // Total columns
    defaultColumnWidth: number;                         // Default width (px)
    defaultRowHeight: number;                           // Default height (px)
    mergeData: IRange[];                                // Merged cells
    cellData: IObjectMatrixPrimitiveType<ICellData>;   // Cell matrix
    rowData: IObjectArrayPrimitiveType<IRowData>;      // Row properties
    columnData: IObjectArrayPrimitiveType<IColumnData>; // Column properties
    showGridlines: BooleanNumber;                       // 0 or 1
    rightToLeft: BooleanNumber;                         // 0 or 1
}
```

### 2. Import Process Flow (What We Must Reverse)

```
Excel → ZIP Extract → XML Parse → LuckySheet → Univer
         ↓             ↓           ↓             ↓
      Files        Elements    Intermediate   Final
                               Structure      Format
```

#### Import Transformations We Need to Reverse:

1. **Cell Data**:
   - Import: Excel XML → LuckySheet celldata → Univer cellData matrix
   - Export: Univer cellData matrix → Excel cells

2. **Formulas**:
   - Import: Excel formulas → LuckySheet f/si → Univer f/si with array tracking
   - Export: Univer f/si → Excel formulas (with proper array formula handling)

3. **Styles**:
   - Import: Excel styles.xml → LuckySheet styles → Univer style registry
   - Export: Univer style registry → Excel styles

4. **Resources (Plugin Data)**:
   - Import: Creates resource entries for each plugin
   - Export: Must properly extract and convert resource data

### 3. Missing Export Features (Based on Import Analysis)

#### Currently Missing in Export:

1. **Filter Export** ❌
   - Import creates: `SHEET_FILTER_PLUGIN` resource
   - Export: Not properly handling filter data

2. **Data Validation Export** ❌
   - Import creates: `SHEET_DATA_VALIDATION_PLUGIN` resource
   - Export: Partial implementation, needs completion

3. **Conditional Formatting Export** ❌
   - Import creates: `SHEET_CONDITIONAL_FORMATTING_PLUGIN` resource
   - Export: Partial implementation, needs enhancement

4. **Drawing/Image Export** ⚠️
   - Import creates: `SHEET_DRAWING_PLUGIN` resource
   - Export: Basic implementation, needs proper coordinate handling

5. **Array Formula Proper Handling** ⚠️
   - Import: Tracks array formulas with range and master cell
   - Export: Implemented but needs validation

### 4. BooleanNumber Handling Issue

#### Problem:
Univer uses `BooleanNumber` (0 or 1) throughout, but export is inconsistent:

```typescript
// INCORRECT in current export:
hidden: hidden === 1 ? 'hidden' : 'visible'  // ❌ String values

// CORRECT per Univer spec:
hidden: BooleanNumber.TRUE   // or 1
hidden: BooleanNumber.FALSE  // or 0
```

### 5. Resource Plugin Mapping

| Plugin Name | Import Creates | Export Status | Action Needed |
|------------|---------------|---------------|---------------|
| SHEET_HYPER_LINK_PLUGIN | ✅ Full | ✅ Implemented | Validate |
| SHEET_DRAWING_PLUGIN | ✅ Full | ⚠️ Basic | Enhance coordinates |
| SHEET_CHART_PLUGIN | ✅ Full | ✅ Implemented | Validate types |
| SHEET_DEFINED_NAME_PLUGIN | ✅ Full | ✅ Implemented | Validate |
| SHEET_CONDITIONAL_FORMATTING_PLUGIN | ✅ Full | ⚠️ Partial | Complete |
| SHEET_DATA_VALIDATION_PLUGIN | ✅ Full | ⚠️ Partial | Complete |
| SHEET_FILTER_PLUGIN | ✅ Full | ❌ Missing | Implement |

### 6. Cell Data Structure Alignment

#### Import Creates (Univer ICellData):
```typescript
interface ICellData {
    v?: string | number;        // Cell value
    s?: string | IStyleData;    // Style (ID or object)
    t?: CellValueType;          // Type (1=string, 2=number, 3=boolean, 4=force text)
    p?: IDocumentData;          // Rich text
    f?: string;                 // Formula
    si?: string;                // Shared formula ID
    custom?: any;               // Custom data
}
```

#### Export Must Handle:
- Boolean values: v=0/1 when t=3
- Style references vs inline styles
- Rich text in p field
- Formula with si for shared formulas
- Array formulas tracked separately

### 7. Formula Handling Alignment

#### Import Process:
1. Reads Excel formulas from XML
2. Creates shared formula IDs (si)
3. Tracks array formulas separately with range info
4. Stores in cellData with f and si fields

#### Export Must:
1. Extract formulas from cellData
2. Handle shared formulas via si field
3. Process array formulas from arrayFormulas array
4. Clean formulas (remove @ symbols, fix syntax)
5. Apply to correct ranges in Excel

### 8. Style Registry Alignment

#### Import Creates:
- Centralized style registry with IDs
- Cells reference styles by ID (string)
- Reduces duplication

#### Export Must:
- Extract styles from registry
- Map Univer style properties to Excel
- Handle both referenced and inline styles
- Preserve all style attributes

## Required Actions for Proper Export

### 1. Complete Missing Resource Exports

```typescript
// Add to Resource.ts
private setFilter() {
    const filters = this.getSheetResource(RESOURCE_PLUGINS.FILTER);
    if (!filters) return;
    // Properly set autoFilter with range
    this.worksheet.autoFilter = this.convertUniverRangeToExcel(filters.ref);
}

private setConditionalFormatting() {
    const conditionals = this.getSheetResource(RESOURCE_PLUGINS.CONDITIONAL_FORMAT);
    // Complete implementation with all rule types
}

private setDataValidation() {
    const validations = this.getSheetResource(RESOURCE_PLUGINS.DATA_VALIDATION);
    // Complete implementation with all validation types
}
```

### 2. Fix BooleanNumber Handling

```typescript
// Create utility in utils.ts
export function exportBooleanNumber(value: BooleanNumber | number | undefined): boolean {
    return value === BooleanNumber.TRUE || value === 1;
}

// Use consistently throughout export
worksheet.hidden = exportBooleanNumber(sheet.hidden) ? 'hidden' : 'visible';
worksheet.showGridlines = exportBooleanNumber(sheet.showGridlines);
```

### 3. Ensure Proper Formula Export

```typescript
// Enhanced formula handling
if (cell.f) {
    // Check if it's array formula first
    const arrayFormula = sheet.arrayFormulas?.find(af => 
        af.formulaId === cell.si
    );
    
    if (arrayFormula) {
        // Handle as array formula
        applyArrayFormula(worksheet, arrayFormula);
    } else if (cell.si) {
        // Handle as shared formula
        applySharedFormula(worksheet, cell.f, cell.si);
    } else {
        // Handle as regular formula
        worksheet.getCell(row, col).formula = cleanFormula(cell.f);
    }
}
```

### 4. Validate Cell Type Handling

```typescript
// Proper cell type export
switch (cell.t) {
    case CellValueType.STRING:    // 1
        target.value = String(cell.v);
        break;
    case CellValueType.NUMBER:    // 2
        target.value = Number(cell.v);
        break;
    case CellValueType.BOOLEAN:   // 3
        target.value = cell.v === 1 || cell.v === '1' || cell.v === true;
        break;
    case CellValueType.FORCE_TEXT: // 4
        target.value = String(cell.v);
        target.numFmt = '@'; // Text format
        break;
}
```

### 5. Rich Text Handling

```typescript
// Handle rich text from p field
if (cell.p && cell.p.body) {
    const richText = convertUniverDocToExcelRichText(cell.p);
    target.value = { richText };
}
```

## Testing Requirements

### Round-Trip Test Cases:

1. **Formula Preservation**:
   - Regular formulas: `=SUM(A1:A10)`
   - Array formulas: `=TRANSPOSE(A1:C3)`
   - Shared formulas: Multiple cells with same formula pattern

2. **Style Preservation**:
   - Font, size, color, background
   - Borders (all sides)
   - Number formats
   - Alignment

3. **Resource Preservation**:
   - Hyperlinks
   - Images with positioning
   - Charts with data ranges
   - Filters and sorting
   - Conditional formatting rules
   - Data validation rules

4. **Structure Preservation**:
   - Sheet order
   - Hidden sheets
   - Frozen panes
   - Merged cells
   - Row/column dimensions

## Conclusion

The export functionality needs significant alignment with the import process to ensure proper data preservation. Key areas requiring immediate attention:

1. Complete missing resource exports (filters, validation, conditional formatting)
2. Fix BooleanNumber handling throughout
3. Ensure formula export matches import structure
4. Validate cell type handling
5. Implement proper rich text support

The export should be a true reverse of the import process, preserving all data and features that Univer supports.