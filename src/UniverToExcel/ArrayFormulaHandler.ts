import { debug } from '../utils/debug';
import { 
    columnNumberToLetter, 
    createRangeReference,
    createRangeId,
    isValidCell 
} from './utils';
import { FORMULA_CONSTANTS } from './constants';

export interface ArrayFormulaInfo {
    formula: string;
    range: { startRow: number; endRow: number; startCol: number; endCol: number };
    masterRow: number;
    masterCol: number;
    formulaId: string;
}

/**
 * ArrayFormulaHandler - Handles array formulas (like TRANSPOSE) for Excel export
 * 
 * This class detects array formulas from Univer data and applies them correctly
 * to ExcelJS worksheets using fillFormula method.
 */
export class ArrayFormulaHandler {
    private arrayFormulas: Map<string, ArrayFormulaInfo> = new Map();
    private processedRanges: Set<string> = new Set();

    /**
     * Check if a cell is part of an array formula
     */
    isArrayFormula(cell: any, sheetData: any, rowIndex: number, colIndex: number): boolean {
        // Validate inputs
        if (!cell || !sheetData || !isValidCell(rowIndex, colIndex)) {
            return false;
        }
        
        // Check if this cell has a shared formula ID that corresponds to an array formula
        if (!cell.si) {
            return false;
        }

        // Look for array formulas in sheet data
        const arrayFormulas = sheetData.arrayFormulas || [];
        
        for (const arrayFormula of arrayFormulas) {
            if (arrayFormula.formulaId === cell.si) {
                const { range } = arrayFormula;
                
                // Validate range object
                if (!this.isValidRange(range)) {
                    debug.warn('🔢 [ArrayFormula] Invalid range in array formula:', arrayFormula);
                    continue;
                }
                
                const isInRange = rowIndex >= range.startRow && 
                                 rowIndex <= range.endRow &&
                                 colIndex >= range.startCol &&
                                 colIndex <= range.endCol;
                
                if (isInRange) {
                    debug.log('🔢 [ArrayFormula] Found array formula cell:', {
                        formula: arrayFormula.formula,
                        range: createRangeReference(range.startRow, range.startCol, range.endRow, range.endCol),
                        cell: `${columnNumberToLetter(colIndex)}${rowIndex + 1}`,
                        formulaId: cell.si
                    });
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * Get array formula info for a cell
     */
    getArrayFormulaInfo(cell: any, sheetData: any): ArrayFormulaInfo | null {
        if (!cell?.si || !sheetData) {
            return null;
        }

        const arrayFormulas = sheetData.arrayFormulas || [];
        
        for (const arrayFormula of arrayFormulas) {
            if (arrayFormula.formulaId === cell.si) {
                // Validate the array formula structure
                if (this.isValidArrayFormula(arrayFormula)) {
                    return arrayFormula;
                }
            }
        }

        return null;
    }

    /**
     * Apply array formula to ExcelJS worksheet
     * This should be called for the master cell of the array formula
     */
    applyArrayFormula(worksheet: any, arrayFormula: ArrayFormulaInfo, startCellValue: any): boolean {
        // Validate inputs
        if (!worksheet || !arrayFormula || !this.isValidArrayFormula(arrayFormula)) {
            debug.warn('🔢 [ArrayFormula] Invalid parameters for applyArrayFormula');
            return false;
        }
        
        const { formula, range } = arrayFormula;
        
        // Create range string (e.g., "A1:C3")
        const rangeStr = createRangeReference(range.startRow, range.startCol, range.endRow, range.endCol);
        
        // Create unique range ID for tracking
        const rangeId = createRangeId(range.startRow, range.startCol, range.endRow, range.endCol);
        
        // Avoid processing the same range twice
        if (this.processedRanges.has(rangeId)) {
            debug.log('🔢 [ArrayFormula] Range already processed:', rangeStr);
            return false;
        }
        
        this.processedRanges.add(rangeId);

        // Clean the formula
        const cleanFormula = this.cleanFormula(formula);
        
        if (!cleanFormula) {
            debug.warn('🔢 [ArrayFormula] Formula cleaning resulted in empty formula');
            return false;
        }

        try {
            debug.log('🔢 [ArrayFormula] Applying array formula:', {
                range: rangeStr,
                formula: cleanFormula,
                masterCell: createRangeReference(
                    arrayFormula.masterRow, 
                    arrayFormula.masterCol,
                    arrayFormula.masterRow,
                    arrayFormula.masterCol
                )
            });

            // CRITICAL FIX: Don't use fillFormula for TRANSPOSE as it adds @ symbols
            // Set the formula only on the master cell, letting Excel handle the spill
            
            const masterRow = arrayFormula.masterRow + 1;  // Convert to 1-based
            const masterCol = arrayFormula.masterCol + 1;  // Convert to 1-based
            const masterCell = worksheet.getCell(masterRow, masterCol);
            
            // Remove any @ symbols from the formula before setting
            const finalFormula = cleanFormula
                .replace(/@TRANSPOSE/gi, 'TRANSPOSE')
                .replace(/@(\$?[A-Z]+\$?\d+)/g, '$1')
                .replace(/@(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)/g, '$1');
            
            // Set as a regular formula, not array formula
            // This avoids ExcelJS adding @ symbols
            masterCell.value = {
                formula: finalFormula,
                result: startCellValue
            };
            
            debug.log('✅ [ArrayFormula] Applied TRANSPOSE without @ symbols:', {
                formula: finalFormula,
                cell: `${columnNumberToLetter(masterCol - 1)}${masterRow}`
            });
            return true;

        } catch (error: any) {
            debug.error('❌ [ArrayFormula] Error applying array formula:', {
                range: rangeStr,
                formula: cleanFormula,
                error: error.message,
                stack: error.stack
            });
            return false;
        }
    }

    /**
     * Check if a range has already been processed
     */
    isRangeProcessed(range: { startRow: number; endRow: number; startCol: number; endCol: number }): boolean {
        if (!this.isValidRange(range)) {
            return false;
        }
        
        const rangeId = createRangeId(range.startRow, range.startCol, range.endRow, range.endCol);
        return this.processedRanges.has(rangeId);
    }

    /**
     * Reset handler for new worksheet
     */
    reset(): void {
        this.arrayFormulas.clear();
        this.processedRanges.clear();
        debug.log('🔢 [ArrayFormula] Handler reset');
    }

    /**
     * Debug: Log current state
     */
    debugState(): void {
        debug.log('🔍 [ArrayFormula] Handler State:', {
            arrayFormulasCount: this.arrayFormulas.size,
            processedRangesCount: this.processedRanges.size,
            processedRanges: Array.from(this.processedRanges)
        });
    }
    
    /**
     * Clean formula for Excel compatibility
     */
    private cleanFormula(formula: string): string {
        if (!formula || typeof formula !== 'string') {
            return '';
        }
        
        let cleaned = formula.trim();
        
        // Remove leading formula start character if present
        if (cleaned.startsWith(FORMULA_CONSTANTS.SPECIAL_CHARS.FORMULA_START)) {
            cleaned = cleaned.substring(1);
        }
        
        // Remove any Excel error codes
        for (const errorCode of Object.values(FORMULA_CONSTANTS.ERROR_CODES)) {
            if (cleaned.includes(errorCode)) {
                debug.warn('🔢 [ArrayFormula] Formula contains error code:', errorCode);
                return '';
            }
        }
        
        return cleaned;
    }
    
    /**
     * Validate range object structure
     */
    private isValidRange(range: any): boolean {
        return range &&
               typeof range === 'object' &&
               typeof range.startRow === 'number' &&
               typeof range.endRow === 'number' &&
               typeof range.startCol === 'number' &&
               typeof range.endCol === 'number' &&
               range.startRow >= 0 &&
               range.endRow >= range.startRow &&
               range.startCol >= 0 &&
               range.endCol >= range.startCol &&
               isValidCell(range.startRow, range.startCol) &&
               isValidCell(range.endRow, range.endCol);
    }
    
    /**
     * Validate array formula structure
     */
    private isValidArrayFormula(arrayFormula: any): boolean {
        return arrayFormula &&
               typeof arrayFormula === 'object' &&
               typeof arrayFormula.formula === 'string' &&
               arrayFormula.formula.length > 0 &&
               this.isValidRange(arrayFormula.range) &&
               typeof arrayFormula.masterRow === 'number' &&
               typeof arrayFormula.masterCol === 'number' &&
               typeof arrayFormula.formulaId === 'string';
    }
    
    /**
     * Get statistics about processed array formulas
     */
    getStatistics(): {
        totalProcessed: number;
        ranges: string[];
        formulas: string[];
    } {
        return {
            totalProcessed: this.processedRanges.size,
            ranges: Array.from(this.processedRanges),
            formulas: Array.from(this.arrayFormulas.values()).map(af => af.formula)
        };
    }
}