/**
 * Enhanced WorkSheet Export with Complete Feature Support
 * 
 * This file properly handles all Univer data types and ensures
 * complete reverse-engineering of the import process.
 */

import { Workbook, Worksheet } from "@zwight/exceljs";
import { debug } from '../utils/debug';
import { 
    BooleanNumber,
    CellValueType,
    fromBooleanNumber, 
    convertCellValue, 
    resolveStyle, 
    hasFormula,
    isArrayFormulaCell,
    getCellNumberFormat,
    convertBorderToExcel,
    convertAlignmentToExcel,
    convertUniverColorToExcel,
    convertUniverDocToExcelRichText
} from './univerUtils';
import { exportAllResources } from './ResourceHandlers';
import { ArrayFormulaHandler } from "./ArrayFormulaHandler";
import { FormulaCleaner } from "./FormulaCleaner";
import { SheetNameHandler } from "./SheetNameHandler";
import { toOneBased } from "./utils";
import { RESOURCE_PLUGINS, BOOLEAN_VALUES } from "./constants";
import { jsonParse } from "../common/method";

/**
 * Enhanced Excel worksheet creation from Univer data
 */
export function createEnhancedWorksheet(workbook: Workbook, snapshot: any): void {
    const { sheetOrder, sheets, styles, resources } = snapshot;
    
    if (!sheetOrder || !Array.isArray(sheetOrder)) {
        debug.warn('[EnhancedWorkSheet] No sheetOrder found in snapshot');
        return;
    }
    
    debug.log('📊 [EnhancedWorkSheet] Processing workbook with sheets:', sheetOrder);
    
    sheetOrder.forEach((sheetId: string) => {
        const sheet = sheets[sheetId];
        
        if (!sheet) {
            debug.warn('⚠️ [EnhancedWorkSheet] Missing sheet data for ID:', sheetId);
            return;
        }
        
        debug.log('📋 [EnhancedWorkSheet] Creating sheet:', {
            id: sheetId,
            name: sheet.name,
            hasData: !!(sheet.cellData && Object.keys(sheet.cellData).length > 0)
        });
        
        // Create worksheet with proper settings
        const worksheet = createWorksheetWithSettings(workbook, sheet);
        
        // Process sheet data
        processSheetStructure(worksheet, sheet);
        processCellData(worksheet, sheet, styles, snapshot);
        processMergedCells(worksheet, sheet.mergeData);
        
        // Export all resources (filters, validation, etc.)
        exportAllResources(worksheet, sheetId, resources);
    });
}

/**
 * Create worksheet with proper Univer settings
 */
function createWorksheetWithSettings(workbook: Workbook, sheet: any): Worksheet {
    const { 
        name, 
        tabColor, 
        hidden,
        rightToLeft,
        showGridlines,
        freeze,
        defaultRowHeight,
        defaultColumnWidth
    } = sheet;
    
    // Sanitize sheet name for Excel
    const safeName = SheetNameHandler.createExcelSafeSheetName(name);
    
    // Create views with proper boolean conversion
    const views: any[] = [{
        rightToLeft: fromBooleanNumber(rightToLeft),
        showGridLines: fromBooleanNumber(showGridlines)
    }];
    
    // Add freeze pane if present
    if (freeze && (freeze.xSplit > 0 || freeze.ySplit > 0)) {
        views[0].state = 'frozen';
        views[0].xSplit = freeze.xSplit;
        views[0].ySplit = freeze.ySplit;
        if (freeze.startRow > 0 || freeze.startColumn > 0) {
            views[0].topLeftCell = `${columnIndexToLetter(freeze.startColumn)}${freeze.startRow + 1}`;
        }
    }
    
    // Create worksheet with settings
    const worksheet = workbook.addWorksheet(safeName, {
        views: views,
        state: fromBooleanNumber(hidden) ? 'hidden' : 'visible',
        properties: {
            tabColor: tabColor ? { argb: hex2argb(tabColor) } : undefined,
            defaultColWidth: defaultColumnWidth ? defaultColumnWidth / 7.5 : undefined,
            defaultRowHeight: defaultRowHeight ? defaultRowHeight * 0.75 : undefined
        }
    });
    
    debug.log('✅ [EnhancedWorkSheet] Created worksheet:', safeName);
    
    return worksheet;
}

/**
 * Process sheet structure (rows and columns)
 */
function processSheetStructure(worksheet: Worksheet, sheet: any): void {
    const { rowData, columnData, defaultRowHeight, defaultColumnWidth } = sheet;
    
    // Process column data
    if (columnData) {
        for (const colIndex in columnData) {
            const colData = columnData[colIndex];
            const col = worksheet.getColumn(toOneBased(Number(colIndex)));
            
            if (colData.w !== undefined) {
                col.width = colData.w / 7.5; // Convert pixels to Excel units
            } else if (defaultColumnWidth) {
                col.width = defaultColumnWidth / 7.5;
            }
            
            col.hidden = fromBooleanNumber(colData.hd);
        }
    }
    
    // Process row data
    if (rowData) {
        for (const rowIndex in rowData) {
            const rowDataItem = rowData[rowIndex];
            const row = worksheet.getRow(toOneBased(Number(rowIndex)));
            
            if (rowDataItem.h !== undefined) {
                row.height = rowDataItem.h * 0.75; // Convert pixels to Excel units
            } else if (defaultRowHeight) {
                row.height = defaultRowHeight * 0.75;
            }
            
            row.hidden = fromBooleanNumber(rowDataItem.hd);
        }
    }
}

/**
 * Process all cell data with proper type handling
 */
function processCellData(worksheet: Worksheet, sheet: any, styles: any, snapshot: any): void {
    const { cellData, id: sheetId, arrayFormulas } = sheet;
    const arrayHandler = new ArrayFormulaHandler();
    
    // Handle empty sheets
    if (!cellData || Object.keys(cellData).length === 0) {
        debug.log('📄 [EnhancedWorkSheet] Processing empty sheet:', sheet.name);
        return;
    }
    
    debug.log('📊 [EnhancedWorkSheet] Processing cell data:', {
        sheetName: sheet.name,
        rowCount: Object.keys(cellData).length,
        hasArrayFormulas: !!(arrayFormulas && arrayFormulas.length > 0)
    });
    
    // Track shared formulas to handle them properly
    const sharedFormulas = new Map<string, { masterCell: string, range: any }>();
    
    // First pass: identify shared formula master cells
    for (const rowIndex in cellData) {
        const row = cellData[rowIndex];
        if (!row) continue;
        
        for (const colIndex in row) {
            const cell = row[colIndex];
            if (!cell) continue;
            
            // If cell has both si (shared formula ID) and f (formula), it's a master cell
            if (cell.si && cell.f) {
                const cellAddr = `${columnIndexToLetter(Number(colIndex))}${Number(rowIndex) + 1}`;
                sharedFormulas.set(cell.si, { 
                    masterCell: cellAddr,
                    range: null // We'll need to determine the range if needed
                });
                debug.log(`📐 [EnhancedWorkSheet] Found shared formula master: ${cellAddr}, ID: ${cell.si}`);
            }
        }
    }
    
    // Second pass: process cells
    for (const rowIndex in cellData) {
        const row = cellData[rowIndex];
        if (!row) continue;
        
        // Process each cell in the row
        for (const colIndex in row) {
            const cell = row[colIndex];
            if (!cell) continue;
            
            const rowNum = Number(rowIndex);
            const colNum = Number(colIndex);
            
            // Check for array formula
            if (arrayFormulas && isArrayFormulaCell(cell, arrayFormulas)) {
                const arrayFormula = arrayFormulas.find((af: any) => af.formulaId === cell.si);
                
                if (arrayFormula && 
                    arrayFormula.masterRow === rowNum && 
                    arrayFormula.masterCol === colNum &&
                    !arrayHandler.isRangeProcessed(arrayFormula.range)) {
                    
                    debug.log('🔢 [EnhancedWorkSheet] Processing array formula master cell');
                    arrayHandler.applyArrayFormula(worksheet, arrayFormula, cell.v);
                    
                    // Apply style to master cell
                    applyCellStyle(worksheet.getCell(toOneBased(rowNum), toOneBased(colNum)), cell, styles);
                    continue;
                } else if (arrayFormula) {
                    // Skip dependent cells of array formula
                    continue;
                }
            }
            
            // Check if this is a dependent cell in a shared formula
            if (cell.si && !cell.f) {
                // This is a dependent cell - ExcelJS will handle it automatically
                // Just set the value, not the formula
                const target = worksheet.getCell(toOneBased(rowNum), toOneBased(colNum));
                target.value = convertCellValue(cell);
                applyCellStyle(target, cell, styles);
                continue;
            }
            
            // Process regular cell (or shared formula master cell)
            processCell(worksheet, rowNum, colNum, cell, styles, snapshot);
        }
    }
}

/**
 * Process individual cell with complete data handling
 */
function processCell(worksheet: Worksheet, row: number, col: number, cell: any, styles: any, snapshot: any): void {
    const target = worksheet.getCell(toOneBased(row), toOneBased(col));
    
    // Handle cell value based on type
    if (cell.p && cell.p.body) {
        // Rich text takes precedence
        target.value = convertUniverDocToExcelRichText(cell.p);
    } else if (hasFormula(cell)) {
        // Handle formula
        processFormula(target, cell, snapshot);
    } else {
        // Handle regular value with proper type conversion
        target.value = convertCellValue(cell);
    }
    
    // Apply cell style
    applyCellStyle(target, cell, styles);
    
    // Apply number format if needed
    if (cell.t === CellValueType.FORCE_STRING) {
        target.numFmt = '@'; // Text format
    }
}

/**
 * Process formula for a cell
 */
function processFormula(target: any, cell: any, snapshot: any): void {
    // IMPORTANT: cell.si is a shared formula ID, not a formula string!
    // Only cell.f contains the actual formula text
    
    if (!cell.f) {
        // No formula in this cell
        // If it has si but no f, it's a dependent cell in a shared formula range
        // ExcelJS will handle these automatically when we set the master cell's formula
        return;
    }
    
    // Log original formula
    debug.log('🔢 [Formula] Processing formula:', {
        original: cell.f,
        hasSi: !!cell.si,
        value: cell.v
    });
    
    // Clean the formula
    const formula = FormulaCleaner.cleanFormula(cell.f);
    
    if (formula) {
        debug.log('✅ [Formula] Setting formula:', {
            cleaned: formula,
            result: cell.v
        });
        target.value = {
            formula: formula,
            result: cell.v !== undefined ? convertCellValue(cell) : undefined
        };
    } else {
        // If formula is invalid, use the cached value
        debug.warn('⚠️ [Formula] Formula cleaned to empty, using cached value:', {
            original: cell.f,
            value: cell.v
        });
        target.value = convertCellValue(cell);
    }
}

/**
 * Apply cell style with complete style resolution
 */
function applyCellStyle(target: any, cell: any, styles: any): void {
    if (!cell.s) return;
    
    // Resolve style from registry or inline
    const style = resolveStyle(cell.s, styles);
    if (!style) return;
    
    // Font styles
    if (style.ff || style.fs || style.cl || style.bl !== undefined || style.it !== undefined) {
        target.font = {
            name: style.ff,
            size: style.fs,
            color: style.cl ? convertUniverColorToExcel(style.cl) : undefined,
            bold: fromBooleanNumber(style.bl),
            italic: fromBooleanNumber(style.it),
            underline: style.ul ? true : undefined,
            strike: style.st ? true : undefined
        };
    }
    
    // Background color
    if (style.bg) {
        target.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: convertUniverColorToExcel(style.bg)
        };
    }
    
    // Borders
    if (style.bd) {
        target.border = convertBorderToExcel(style.bd);
    }
    
    // Alignment
    const alignment = convertAlignmentToExcel(style);
    if (alignment) {
        target.alignment = alignment;
    }
    
    // Number format
    const numFmt = getCellNumberFormat(cell, style);
    if (numFmt) {
        target.numFmt = numFmt;
    }
}

/**
 * Process merged cells
 */
function processMergedCells(worksheet: Worksheet, mergeData: any[]): void {
    if (!mergeData || !Array.isArray(mergeData)) return;
    
    for (const merge of mergeData) {
        if (merge.startRow !== undefined && 
            merge.startColumn !== undefined && 
            merge.endRow !== undefined && 
            merge.endColumn !== undefined) {
            
            worksheet.mergeCells(
                toOneBased(merge.startRow),
                toOneBased(merge.startColumn),
                toOneBased(merge.endRow),
                toOneBased(merge.endColumn)
            );
        }
    }
}

/**
 * Helper function to convert column index to letter
 */
function columnIndexToLetter(index: number): string {
    let result = '';
    let num = index;
    
    while (num >= 0) {
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26) - 1;
        if (num < 0) break;
    }
    
    return result;
}

/**
 * Helper to convert hex color to ARGB
 */
function hex2argb(hex: string): string {
    if (!hex) return 'FF000000';
    hex = hex.replace('#', '');
    if (hex.length === 6) {
        return 'FF' + hex.toUpperCase();
    } else if (hex.length === 8) {
        return hex.toUpperCase();
    }
    return 'FF000000';
}