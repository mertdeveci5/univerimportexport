import { debug } from '../utils/debug';

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
        // Check if this cell has a shared formula ID that corresponds to an array formula
        if (!cell.si) return false;

        // Look for array formulas in sheet data
        const arrayFormulas = sheetData.arrayFormulas || [];
        for (const arrayFormula of arrayFormulas) {
            if (arrayFormula.formulaId === cell.si) {
                const { range } = arrayFormula;
                const isInRange = rowIndex >= range.startRow && 
                                 rowIndex <= range.endRow &&
                                 colIndex >= range.startCol &&
                                 colIndex <= range.endCol;
                
                if (isInRange) {
                    debug.log('🔢 [ArrayFormula] Found array formula cell:', {
                        formula: arrayFormula.formula,
                        range: `${this.numberToColumnLetter(range.startCol)}${range.startRow + 1}:${this.numberToColumnLetter(range.endCol)}${range.endRow + 1}`,
                        cell: `${this.numberToColumnLetter(colIndex)}${rowIndex + 1}`,
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
        if (!cell.si) return null;

        const arrayFormulas = sheetData.arrayFormulas || [];
        for (const arrayFormula of arrayFormulas) {
            if (arrayFormula.formulaId === cell.si) {
                return arrayFormula;
            }
        }

        return null;
    }

    /**
     * Apply array formula to ExcelJS worksheet
     * This should be called for the master cell of the array formula
     */
    applyArrayFormula(worksheet: any, arrayFormula: ArrayFormulaInfo, startCellValue: any): boolean {
        const { formula, range, masterRow, masterCol } = arrayFormula;
        
        // Create range string (e.g., "A1:C3")
        const rangeStr = `${this.numberToColumnLetter(range.startCol)}${range.startRow + 1}:${this.numberToColumnLetter(range.endCol)}${range.endRow + 1}`;
        
        // Avoid processing the same range twice
        if (this.processedRanges.has(rangeStr)) {
            return false;
        }
        this.processedRanges.add(rangeStr);

        // Clean the formula (remove leading =)
        let cleanFormula = formula;
        if (cleanFormula.startsWith('=')) {
            cleanFormula = cleanFormula.substring(1);
        }

        try {
            debug.log('🔢 [ArrayFormula] Applying array formula:', {
                range: rangeStr,
                formula: cleanFormula,
                masterCell: `${this.numberToColumnLetter(masterCol)}${masterRow + 1}`
            });

            // Use ExcelJS fillFormula for array formulas
            // This creates the array formula across the entire range
            worksheet.fillFormula(rangeStr, cleanFormula, startCellValue);
            
            debug.log('✅ [ArrayFormula] Successfully applied array formula');
            return true;

        } catch (error: any) {
            debug.log('❌ [ArrayFormula] Error applying array formula:', error);
            debug.log('❌ [ArrayFormula] Failed formula details:', {
                range: rangeStr,
                formula: cleanFormula,
                error: error.message
            });
            return false;
        }
    }

    /**
     * Check if a range has already been processed
     */
    isRangeProcessed(range: { startRow: number; endRow: number; startCol: number; endCol: number }): boolean {
        const rangeStr = `${this.numberToColumnLetter(range.startCol)}${range.startRow + 1}:${this.numberToColumnLetter(range.endCol)}${range.endRow + 1}`;
        return this.processedRanges.has(rangeStr);
    }

    /**
     * Convert column number to Excel column letter (0-based)
     * 0 -> A, 1 -> B, 25 -> Z, 26 -> AA, etc.
     */
    private numberToColumnLetter(colNumber: number): string {
        let result = '';
        let num = colNumber;
        
        while (num >= 0) {
            result = String.fromCharCode(65 + (num % 26)) + result;
            num = Math.floor(num / 26) - 1;
            if (num < 0) break;
        }
        
        return result;
    }

    /**
     * Reset handler for new worksheet
     */
    reset(): void {
        this.arrayFormulas.clear();
        this.processedRanges.clear();
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
}