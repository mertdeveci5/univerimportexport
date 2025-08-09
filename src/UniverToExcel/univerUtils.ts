/**
 * Univer-specific utility functions for proper data conversion
 * 
 * This file ensures proper handling of Univer data types during export
 * to maintain perfect symmetry with the import process.
 */

import { BooleanNumber, CellValueType } from '@univerjs/core';
import { debug } from '../utils/debug';

/**
 * Convert Univer BooleanNumber to JavaScript boolean
 * Univer uses 0 and 1 to represent false and true
 * 
 * @param value - BooleanNumber value (0, 1, or undefined)
 * @returns JavaScript boolean
 */
export function fromBooleanNumber(value: BooleanNumber | number | undefined): boolean {
    return value === BooleanNumber.TRUE || value === 1;
}

/**
 * Convert JavaScript boolean to Univer BooleanNumber
 * 
 * @param value - JavaScript boolean
 * @returns BooleanNumber (0 or 1)
 */
export function toBooleanNumber(value: boolean | undefined): BooleanNumber {
    return value ? BooleanNumber.TRUE : BooleanNumber.FALSE;
}

/**
 * Convert Univer cell value based on cell type
 * Properly handles the different CellValueType cases
 * 
 * @param cell - Cell data from Univer
 * @returns Properly typed value for Excel
 */
export function convertCellValue(cell: any): any {
    if (!cell) return undefined;
    
    const { v, t, p } = cell;
    
    // Rich text takes precedence
    if (p && p.body) {
        return convertUniverDocToExcelRichText(p);
    }
    
    // Handle based on cell type
    switch (t) {
        case CellValueType.STRING: // 1
            return String(v ?? '');
            
        case CellValueType.NUMBER: // 2
            return typeof v === 'number' ? v : Number(v);
            
        case CellValueType.BOOLEAN: // 3
            // In Univer, boolean values are stored as 0 or 1
            return v === 1 || v === '1' || v === true;
            
        case CellValueType.FORCE_STRING: // 4
            // Force string means treat number as string
            return String(v ?? '');
            
        default:
            // Auto-detect type if not specified
            if (v === null || v === undefined) {
                return '';
            }
            if (typeof v === 'boolean') {
                return v;
            }
            if (typeof v === 'number' || !isNaN(Number(v))) {
                return Number(v);
            }
            return String(v);
    }
}

/**
 * Convert Univer IDocumentData (rich text) to Excel rich text format
 * 
 * @param doc - Univer document data
 * @returns Excel rich text array
 */
export function convertUniverDocToExcelRichText(doc: any): any {
    if (!doc || !doc.body) {
        return [];
    }
    
    const { body } = doc;
    const { dataStream, textRuns, paragraphs } = body;
    
    if (!dataStream || !textRuns) {
        return [];
    }
    
    const richTextParts: any[] = [];
    
    // Process each text run
    for (const textRun of textRuns) {
        const { st, ed, ts } = textRun;
        const text = dataStream.substring(st, ed);
        
        if (!text) continue;
        
        const part: any = { text };
        
        // Convert text styles to Excel format
        if (ts) {
            const font: any = {};
            
            if (ts.ff) font.name = ts.ff;        // Font family
            if (ts.fs) font.size = ts.fs;        // Font size
            if (ts.cl) font.color = convertUniverColorToExcel(ts.cl); // Color
            if (ts.bl === 1) font.bold = true;   // Bold
            if (ts.it === 1) font.italic = true; // Italic
            if (ts.ul) font.underline = true;    // Underline
            if (ts.st) font.strike = true;       // Strikethrough
            
            // Superscript/subscript
            if (ts.va === 2) font.vertAlign = 'subscript';
            if (ts.va === 3) font.vertAlign = 'superscript';
            
            if (Object.keys(font).length > 0) {
                part.font = font;
            }
        }
        
        richTextParts.push(part);
    }
    
    return richTextParts.length > 0 ? { richText: richTextParts } : '';
}

/**
 * Convert Univer color object to Excel color format
 * 
 * @param color - Univer color object
 * @returns Excel ARGB color string
 */
export function convertUniverColorToExcel(color: any): any {
    if (!color) return undefined;
    
    if (typeof color === 'string') {
        // Already a color string
        return color.startsWith('#') ? color.replace('#', 'FF') : color;
    }
    
    if (color.rgb) {
        // Convert RGB to ARGB
        const rgb = color.rgb.replace('#', '');
        return { argb: 'FF' + rgb };
    }
    
    if (color.theme !== undefined) {
        // Theme color
        return { theme: color.theme };
    }
    
    return undefined;
}

/**
 * Convert Univer range to Excel range string
 * 
 * @param range - Univer range object
 * @returns Excel range string (e.g., "A1:C3")
 */
export function convertUniverRangeToExcel(range: any): string | undefined {
    if (!range) return undefined;
    
    // Handle string range
    if (typeof range === 'string') {
        return range;
    }
    
    // Handle object range
    if (range.startRow !== undefined && 
        range.startColumn !== undefined && 
        range.endRow !== undefined && 
        range.endColumn !== undefined) {
        
        const startCol = columnIndexToLetter(range.startColumn);
        const endCol = columnIndexToLetter(range.endColumn);
        const startRow = range.startRow + 1;
        const endRow = range.endRow + 1;
        
        return `${startCol}${startRow}:${endCol}${endRow}`;
    }
    
    return undefined;
}

/**
 * Convert column index to Excel letter
 * 
 * @param index - 0-based column index
 * @returns Excel column letter
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
 * Resolve style from Univer style registry
 * 
 * @param styleRef - Style reference (string ID or inline object)
 * @param styles - Style registry
 * @returns Resolved style object
 */
export function resolveStyle(styleRef: string | any, styles: Record<string, any>): any {
    if (!styleRef) return null;
    
    // If it's a string, look up in registry
    if (typeof styleRef === 'string') {
        const resolved = styles[styleRef];
        if (!resolved) {
            debug.warn('Style ID not found in registry:', styleRef);
        }
        return resolved || null;
    }
    
    // Otherwise it's an inline style object
    return styleRef;
}

/**
 * Check if a cell has a formula
 * 
 * @param cell - Cell data
 * @returns True if cell has a formula
 */
export function hasFormula(cell: any): boolean {
    return !!(cell && (cell.f || cell.si));
}

/**
 * Check if a cell is part of an array formula
 * 
 * @param cell - Cell data
 * @param arrayFormulas - Array formulas for the sheet
 * @returns True if cell is part of an array formula
 */
export function isArrayFormulaCell(cell: any, arrayFormulas: any[]): boolean {
    if (!cell || !cell.si || !arrayFormulas) return false;
    
    return arrayFormulas.some(af => af.formulaId === cell.si);
}

/**
 * Get number format for cell
 * 
 * @param cell - Cell data
 * @param style - Resolved style object
 * @returns Number format string
 */
export function getCellNumberFormat(cell: any, style: any): string | undefined {
    // Force string type
    if (cell.t === CellValueType.FORCE_STRING) {
        return '@'; // Excel text format
    }
    
    // From style
    if (style && style.n && style.n.pattern) {
        return style.n.pattern;
    }
    
    return undefined;
}

/**
 * Convert Univer border style to Excel
 * 
 * @param border - Univer border object
 * @returns Excel border object
 */
export function convertBorderToExcel(border: any): any {
    if (!border) return undefined;
    
    const excelBorder: any = {};
    
    // Map border positions
    const positions = {
        t: 'top',
        b: 'bottom',
        l: 'left',
        r: 'right',
        tl_br: 'diagonal',
        bl_tr: 'diagonal'
    };
    
    for (const [univerPos, excelPos] of Object.entries(positions)) {
        if (border[univerPos]) {
            const b = border[univerPos];
            excelBorder[excelPos] = {
                style: convertBorderStyle(b.s),
                color: convertUniverColorToExcel(b.cl)
            };
        }
    }
    
    // Handle diagonal direction
    if (border.tl_br || border.bl_tr) {
        excelBorder.diagonalUp = !!border.bl_tr;
        excelBorder.diagonalDown = !!border.tl_br;
    }
    
    return Object.keys(excelBorder).length > 0 ? excelBorder : undefined;
}

/**
 * Convert Univer border style number to Excel style string
 * 
 * @param styleNum - Univer border style number
 * @returns Excel border style string
 */
function convertBorderStyle(styleNum: number): string {
    const styleMap: { [key: number]: string } = {
        0: 'thin',
        1: 'thin',
        2: 'hair',
        3: 'dotted',
        4: 'dashed',
        5: 'dashDot',
        6: 'dashDotDot',
        7: 'double',
        8: 'medium',
        9: 'mediumDashed',
        10: 'mediumDashDot',
        11: 'mediumDashDotDot',
        12: 'slantDashDot',
        13: 'thick'
    };
    
    return styleMap[styleNum] || 'thin';
}

/**
 * Convert Univer alignment to Excel
 * 
 * @param style - Univer style object
 * @returns Excel alignment object
 */
export function convertAlignmentToExcel(style: any): any {
    if (!style) return undefined;
    
    const alignment: any = {};
    
    // Horizontal alignment (ht)
    if (style.ht !== undefined) {
        const hMap: { [key: number]: string } = {
            0: 'center',
            1: 'left',
            2: 'right',
            3: 'justify'
        };
        alignment.horizontal = hMap[style.ht] || 'left';
    }
    
    // Vertical alignment (vt)
    if (style.vt !== undefined) {
        const vMap: { [key: number]: string } = {
            0: 'middle',
            1: 'top',
            2: 'bottom'
        };
        alignment.vertical = vMap[style.vt] || 'top';
    }
    
    // Text wrap (tb)
    if (style.tb === 3) {
        alignment.wrapText = true;
    }
    
    // Text rotation (tr)
    if (style.tr) {
        if (style.tr.v === 1) {
            alignment.textRotation = 90; // Vertical text
        } else if (style.tr.a !== undefined) {
            alignment.textRotation = style.tr.a;
        }
    }
    
    // Indent (pd)
    if (style.pd && style.pd.l) {
        alignment.indent = style.pd.l;
    }
    
    return Object.keys(alignment).length > 0 ? alignment : undefined;
}