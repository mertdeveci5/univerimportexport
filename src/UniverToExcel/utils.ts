/**
 * Utility functions for Excel Export
 * 
 * This file contains shared utility functions to avoid code duplication
 * and ensure consistent behavior across the export functionality.
 */

import { EXCEL_COLUMN_CONSTANTS, INDEX_OFFSET } from './constants';
import { debug } from '../utils/debug';

/**
 * Convert a 0-based column number to Excel column letter(s)
 * Examples: 0 -> 'A', 1 -> 'B', 25 -> 'Z', 26 -> 'AA', 27 -> 'AB'
 * 
 * @param colNumber - 0-based column index
 * @returns Excel column letter(s)
 * @throws Error if column number is negative or exceeds Excel's maximum
 */
export function columnNumberToLetter(colNumber: number): string {
    if (colNumber < 0) {
        throw new Error(`Invalid column number: ${colNumber}. Column numbers must be non-negative.`);
    }
    
    if (colNumber >= EXCEL_COLUMN_CONSTANTS.MAX_COLUMNS) {
        throw new Error(`Column number ${colNumber} exceeds Excel's maximum of ${EXCEL_COLUMN_CONSTANTS.MAX_COLUMNS - 1}`);
    }
    
    let result = '';
    let num = colNumber;
    
    while (num >= 0) {
        result = String.fromCharCode(
            EXCEL_COLUMN_CONSTANTS.ASCII_UPPERCASE_A + (num % EXCEL_COLUMN_CONSTANTS.ALPHABET_SIZE)
        ) + result;
        num = Math.floor(num / EXCEL_COLUMN_CONSTANTS.ALPHABET_SIZE) - 1;
        if (num < 0) break;
    }
    
    return result;
}

/**
 * Convert Excel column letter(s) to 0-based column number
 * Examples: 'A' -> 0, 'B' -> 1, 'Z' -> 25, 'AA' -> 26, 'AB' -> 27
 * 
 * @param columnLetter - Excel column letter(s)
 * @returns 0-based column index
 * @throws Error if column letter is invalid
 */
export function columnLetterToNumber(columnLetter: string): number {
    if (!columnLetter || !/^[A-Z]+$/i.test(columnLetter)) {
        throw new Error(`Invalid column letter: "${columnLetter}". Must be one or more letters A-Z.`);
    }
    
    const upper = columnLetter.toUpperCase();
    let result = 0;
    
    for (let i = 0; i < upper.length; i++) {
        const charValue = upper.charCodeAt(i) - EXCEL_COLUMN_CONSTANTS.ASCII_UPPERCASE_A + 1;
        result = result * EXCEL_COLUMN_CONSTANTS.ALPHABET_SIZE + charValue;
    }
    
    return result - 1; // Convert to 0-based
}

/**
 * Convert 0-based row/column indices to 1-based for Excel
 * 
 * @param zeroBasedIndex - 0-based index
 * @returns 1-based index for Excel
 */
export function toOneBased(zeroBasedIndex: number): number {
    return zeroBasedIndex + INDEX_OFFSET.ZERO_TO_ONE_BASED;
}

/**
 * Convert 1-based Excel indices to 0-based
 * 
 * @param oneBasedIndex - 1-based Excel index
 * @returns 0-based index
 */
export function toZeroBased(oneBasedIndex: number): number {
    return oneBasedIndex - INDEX_OFFSET.ZERO_TO_ONE_BASED;
}

/**
 * Create an Excel cell reference from row and column indices
 * 
 * @param row - 0-based row index
 * @param col - 0-based column index
 * @param absolute - Whether to use absolute references ($)
 * @returns Excel cell reference (e.g., 'A1', '$A$1')
 */
export function createCellReference(
    row: number, 
    col: number, 
    absolute: { row?: boolean; col?: boolean } = {}
): string {
    const colLetter = columnNumberToLetter(col);
    const rowNumber = toOneBased(row);
    
    const colRef = absolute.col ? `$${colLetter}` : colLetter;
    const rowRef = absolute.row ? `$${rowNumber}` : `${rowNumber}`;
    
    return `${colRef}${rowRef}`;
}

/**
 * Create an Excel range reference from start and end coordinates
 * 
 * @param startRow - 0-based start row
 * @param startCol - 0-based start column
 * @param endRow - 0-based end row
 * @param endCol - 0-based end column
 * @returns Excel range reference (e.g., 'A1:C3')
 */
export function createRangeReference(
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
): string {
    const startRef = createCellReference(startRow, startCol);
    const endRef = createCellReference(endRow, endCol);
    return `${startRef}:${endRef}`;
}

/**
 * Parse an Excel cell reference to get row and column indices
 * 
 * @param cellRef - Excel cell reference (e.g., 'A1', '$B$2')
 * @returns Object with 0-based row and column indices and absolute flags
 */
export function parseCellReference(cellRef: string): {
    row: number;
    col: number;
    absoluteRow: boolean;
    absoluteCol: boolean;
} | null {
    const match = cellRef.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
    
    if (!match) {
        debug.warn(`Invalid cell reference: ${cellRef}`);
        return null;
    }
    
    const [, dollarCol, colLetters, dollarRow, rowStr] = match;
    
    return {
        row: toZeroBased(parseInt(rowStr, 10)),
        col: columnLetterToNumber(colLetters),
        absoluteRow: dollarRow === '$',
        absoluteCol: dollarCol === '$'
    };
}

/**
 * Parse an Excel range reference to get start and end coordinates
 * 
 * @param rangeRef - Excel range reference (e.g., 'A1:C3')
 * @returns Object with start and end coordinates (0-based)
 */
export function parseRangeReference(rangeRef: string): {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
} | null {
    const parts = rangeRef.split(':');
    
    if (parts.length !== 2) {
        debug.warn(`Invalid range reference: ${rangeRef}`);
        return null;
    }
    
    const start = parseCellReference(parts[0]);
    const end = parseCellReference(parts[1]);
    
    if (!start || !end) {
        return null;
    }
    
    return {
        startRow: start.row,
        startCol: start.col,
        endRow: end.row,
        endCol: end.col
    };
}

/**
 * Validate if a value represents a boolean in Excel terms
 * 
 * @param value - Value to check
 * @returns Boolean representation or undefined if not boolean
 */
export function toBooleanValue(value: any): boolean | undefined {
    if (value === 1 || value === '1' || value === true) return true;
    if (value === 0 || value === '0' || value === false) return false;
    return undefined;
}

/**
 * Check if a row index is valid for Excel
 * 
 * @param row - 0-based row index
 * @returns True if valid
 */
export function isValidRow(row: number): boolean {
    return row >= 0 && row < EXCEL_COLUMN_CONSTANTS.MAX_ROWS;
}

/**
 * Check if a column index is valid for Excel
 * 
 * @param col - 0-based column index
 * @returns True if valid
 */
export function isValidColumn(col: number): boolean {
    return col >= 0 && col < EXCEL_COLUMN_CONSTANTS.MAX_COLUMNS;
}

/**
 * Check if a cell reference is valid
 * 
 * @param row - 0-based row index
 * @param col - 0-based column index
 * @returns True if valid cell reference
 */
export function isValidCell(row: number, col: number): boolean {
    return isValidRow(row) && isValidColumn(col);
}

/**
 * Safely get a value from a nested object structure
 * 
 * @param obj - Object to get value from
 * @param path - Path to the value (e.g., 'a.b.c')
 * @param defaultValue - Default value if path doesn't exist
 * @returns Value at path or default
 */
export function safeGet<T = any>(obj: any, path: string, defaultValue?: T): T {
    const keys = path.split('.');
    let current = obj;
    
    for (const key of keys) {
        if (current == null || typeof current !== 'object') {
            return defaultValue as T;
        }
        current = current[key];
    }
    
    return current ?? defaultValue;
}

/**
 * Create a unique identifier for a range
 * 
 * @param startRow - Start row (0-based)
 * @param startCol - Start column (0-based)
 * @param endRow - End row (0-based)
 * @param endCol - End column (0-based)
 * @returns Unique range identifier
 */
export function createRangeId(
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
): string {
    return `${startRow}_${startCol}_${endRow}_${endCol}`;
}

/**
 * Normalize a sheet name for comparison
 * 
 * @param name - Sheet name
 * @returns Normalized sheet name
 */
export function normalizeSheetName(name: string): string {
    return name.trim().toLowerCase();
}

/**
 * Check if two ranges overlap
 * 
 * @param range1 - First range
 * @param range2 - Second range
 * @returns True if ranges overlap
 */
export function doRangesOverlap(
    range1: { startRow: number; startCol: number; endRow: number; endCol: number },
    range2: { startRow: number; startCol: number; endRow: number; endCol: number }
): boolean {
    return !(
        range1.endRow < range2.startRow ||
        range2.endRow < range1.startRow ||
        range1.endCol < range2.startCol ||
        range2.endCol < range1.startCol
    );
}

/**
 * Deep clone an object (simple implementation for plain objects)
 * 
 * @param obj - Object to clone
 * @returns Cloned object
 */
export function deepClone<T>(obj: T): T {
    if (obj === null || typeof obj !== 'object') {
        return obj;
    }
    
    if (obj instanceof Date) {
        return new Date(obj.getTime()) as any;
    }
    
    if (Array.isArray(obj)) {
        return obj.map(item => deepClone(item)) as any;
    }
    
    const cloned = {} as T;
    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            cloned[key] = deepClone(obj[key]);
        }
    }
    
    return cloned;
}