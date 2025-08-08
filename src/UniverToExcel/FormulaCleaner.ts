import { debug } from '../utils/debug';
import { SheetNameHandler } from './SheetNameHandler';

/**
 * FormulaCleaner - Handles formula cleaning for Excel export
 * 
 * This class fixes common issues with formulas when exporting to Excel:
 * 1. @ symbols incorrectly added by ExcelJS
 * 2. Double equals (==) issues
 * 3. Invalid formula syntax
 * 4. Special character handling
 */
export class FormulaCleaner {
    
    /**
     * Clean formula for Excel export
     * This is the main method that applies all cleaning rules
     */
    static cleanFormula(formula: string): string {
        if (!formula || typeof formula !== 'string') {
            return '';
        }

        let cleaned = formula;
        
        // Step 1: Remove leading = if present (ExcelJS adds its own)
        if (cleaned.startsWith('=')) {
            cleaned = cleaned.substring(1);
            debug.log('[FormulaCleaner] Removed leading =');
        }

        // Step 2: Remove @ symbols that ExcelJS incorrectly adds
        cleaned = this.removeIncorrectAtSymbols(cleaned);
        
        // Step 3: Fix Excel function names
        cleaned = this.fixFunctionNames(cleaned);
        
        // Step 4: Validate formula syntax
        cleaned = this.validateSyntax(cleaned);
        
        // Step 5: Handle special characters
        cleaned = this.handleSpecialCharacters(cleaned);

        if (formula !== cleaned) {
            debug.log('[FormulaCleaner] Formula cleaned:', {
                original: formula,
                cleaned: cleaned,
                changes: this.getChanges(formula, cleaned)
            });
        }

        return cleaned;
    }

    /**
     * Remove @ symbols that ExcelJS incorrectly places
     */
    private static removeIncorrectAtSymbols(formula: string): string {
        let cleaned = formula;
        
        // Fix 1: Remove @ from function names (like @TRANSPOSE -> TRANSPOSE)
        cleaned = cleaned.replace(/@([A-Z][A-Z0-9_]*)\s*\(/g, '$1(');
        
        // Fix 2: Remove @ from cell references (like @$N$43 or @A1)
        cleaned = cleaned.replace(/@(\$?[A-Z]+\$?\d+)/g, '$1');
        
        // Fix 3: Remove @ from named ranges (but keep structured references like @[Column])
        cleaned = cleaned.replace(/@([A-Za-z_][A-Za-z0-9_]*)\b(?![\[])/g, '$1');
        
        // Fix 4: Remove @ from range references (like @A1:C3)
        cleaned = cleaned.replace(/@(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)/g, '$1');
        
        return cleaned;
    }

    /**
     * Fix Excel function names and syntax
     */
    private static fixFunctionNames(formula: string): string {
        let cleaned = formula;
        
        // Remove Excel namespace prefixes that might cause issues
        cleaned = cleaned.replace(/=_xlfn\./g, '=');
        cleaned = cleaned.replace(/_xlfn\./g, '');
        
        // Fix TRANSPOSE function specifically (common array formula)
        cleaned = cleaned.replace(/TRANSPOSE\s*\(\s*@/g, 'TRANSPOSE(');
        
        return cleaned;
    }

    /**
     * Validate and fix basic formula syntax
     */
    private static validateSyntax(formula: string): string {
        if (!formula) return '';
        
        let cleaned = formula.trim();
        
        // Check for basic syntax issues
        if (cleaned.includes('#REF!')) {
            debug.log('[FormulaCleaner] Formula contains #REF! error:', cleaned);
            return ''; // Invalid formula
        }
        
        if (cleaned.includes('#NAME?')) {
            debug.log('[FormulaCleaner] Formula contains #NAME? error:', cleaned);
            return ''; // Invalid formula
        }
        
        // Check for balanced parentheses
        const openParens = (cleaned.match(/\(/g) || []).length;
        const closeParens = (cleaned.match(/\)/g) || []).length;
        
        if (openParens !== closeParens) {
            debug.log('[FormulaCleaner] Unbalanced parentheses in formula:', {
                formula: cleaned,
                openParens,
                closeParens
            });
        }
        
        return cleaned;
    }

    /**
     * Handle special characters in formulas
     */
    private static handleSpecialCharacters(formula: string): string {
        let cleaned = formula;
        
        // Handle sheet names with special characters using SheetNameHandler
        // Find sheet references (format: SheetName!CellRef)
        const sheetRefPattern = /([A-Za-z0-9_>>>\s<>!@#$%^&()+={}|;",.\-]+)!([A-Z]+\d+)/g;
        cleaned = cleaned.replace(sheetRefPattern, (match, sheetName, cellRef) => {
            // Use SheetNameHandler to format sheet name properly for formulas
            const formattedSheetName = SheetNameHandler.formatSheetNameForFormula(sheetName);
            return `${formattedSheetName}!${cellRef}`;
        });
        
        return cleaned;
    }

    /**
     * Get summary of changes made during cleaning
     */
    private static getChanges(original: string, cleaned: string): string[] {
        const changes: string[] = [];
        
        if (original.startsWith('=') && !cleaned.startsWith('=')) {
            changes.push('Removed leading =');
        }
        
        if (original.includes('@') && !cleaned.includes('@')) {
            changes.push('Removed @ symbols');
        }
        
        if (original.includes('_xlfn.') && !cleaned.includes('_xlfn.')) {
            changes.push('Removed _xlfn. prefix');
        }
        
        if (original !== cleaned && changes.length === 0) {
            changes.push('Other syntax fixes');
        }
        
        return changes;
    }

    /**
     * Check if a formula is valid for Excel export
     */
    static isValidFormula(formula: string): boolean {
        if (!formula || typeof formula !== 'string') {
            return false;
        }
        
        // Check for obvious errors
        if (formula.includes('#REF!') || formula.includes('#NAME?')) {
            return false;
        }
        
        // Formula should not be empty after cleaning
        const cleaned = this.cleanFormula(formula);
        return cleaned.length > 0;
    }

    /**
     * Special handling for array formulas
     */
    static cleanArrayFormula(formula: string, range?: string): string {
        let cleaned = this.cleanFormula(formula);
        
        // For array formulas, ensure they're properly formatted
        if (range && cleaned && !cleaned.includes(range)) {
            debug.log('[FormulaCleaner] Array formula cleaned:', {
                formula: cleaned,
                range: range
            });
        }
        
        return cleaned;
    }
}