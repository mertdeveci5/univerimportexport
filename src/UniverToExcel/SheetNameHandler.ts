import { debug } from '../utils/debug';

/**
 * SheetNameHandler - Handles sheet name validation and sanitization for Excel export
 * 
 * Excel has specific rules for sheet names:
 * 1. Cannot exceed 31 characters
 * 2. Cannot contain: \ / ? * [ ] : 
 * 3. Cannot start or end with single quotes
 * 4. Special characters like >>> need careful handling
 */
export class SheetNameHandler {
    
    // Characters that are invalid in Excel sheet names
    private static readonly INVALID_CHARS = /[\\\/\?\*\[\]:]/g;
    private static readonly MAX_LENGTH = 31;
    
    /**
     * Sanitize sheet name for Excel export
     * This ensures the sheet name is valid for Excel while preserving readability
     */
    static sanitizeSheetName(name: string): string {
        if (!name || typeof name !== 'string') {
            return 'Sheet1'; // Default fallback
        }

        let sanitized = name;
        
        // Step 1: Handle special characters like >>> 
        // These are valid in Univer but need to be preserved carefully in Excel
        const originalName = sanitized;
        
        // Step 2: Replace invalid characters with underscores
        sanitized = sanitized.replace(this.INVALID_CHARS, '_');
        
        // Step 3: Remove leading/trailing single quotes if present
        sanitized = sanitized.replace(/^'|'$/g, '');
        
        // Step 4: Ensure maximum length
        if (sanitized.length > this.MAX_LENGTH) {
            sanitized = sanitized.substring(0, this.MAX_LENGTH);
        }
        
        // Step 5: Ensure name is not empty after sanitization
        if (!sanitized.trim()) {
            sanitized = 'Sheet1';
        }
        
        // Step 6: Handle reserved names or duplicates (handled by ExcelJS, but log it)
        const reservedNames = ['History', 'Print_Titles', 'Print_Area', 'Sheet_Title'];
        if (reservedNames.includes(sanitized)) {
            debug.log('⚠️ [SheetName] Reserved sheet name detected:', sanitized);
        }
        
        if (originalName !== sanitized) {
            debug.log('🔄 [SheetName] Sheet name sanitized:', {
                original: originalName,
                sanitized: sanitized,
                changes: this.getChanges(originalName, sanitized)
            });
        }
        
        return sanitized;
    }

    /**
     * Check if a sheet name is valid for Excel
     */
    static isValidSheetName(name: string): boolean {
        if (!name || typeof name !== 'string') {
            return false;
        }
        
        // Check length
        if (name.length > this.MAX_LENGTH) {
            return false;
        }
        
        // Check for invalid characters
        if (this.INVALID_CHARS.test(name)) {
            return false;
        }
        
        // Check for leading/trailing quotes
        if (name.startsWith("'") || name.endsWith("'")) {
            return false;
        }
        
        // Check for empty after trim
        if (!name.trim()) {
            return false;
        }
        
        return true;
    }

    /**
     * Create Excel-safe sheet name while preserving special characters when possible
     * This method is more sophisticated than sanitizeSheetName and tries to keep
     * special characters like >>> intact when they won't cause Excel issues
     */
    static createExcelSafeSheetName(name: string): string {
        if (!name || typeof name !== 'string') {
            return 'Sheet1';
        }

        let safeName = name;
        
        // Special handling for >>> characters (common in our test cases)
        // These are actually valid in Excel sheet names, just need to be quoted in formulas
        
        // Replace only truly problematic characters
        const problemChars = {
            '\\': '_',
            '/': '_',
            '?': '_',
            '*': '_',
            '[': '(',
            ']': ')',
            ':': '_'
        };
        
        for (const [invalid, replacement] of Object.entries(problemChars)) {
            safeName = safeName.replace(new RegExp(`\\${invalid}`, 'g'), replacement);
        }
        
        // Handle length
        if (safeName.length > this.MAX_LENGTH) {
            safeName = safeName.substring(0, this.MAX_LENGTH);
        }
        
        // Ensure not empty
        if (!safeName.trim()) {
            safeName = 'Sheet1';
        }
        
        return safeName;
    }

    /**
     * Get summary of changes made during sanitization
     */
    private static getChanges(original: string, sanitized: string): string[] {
        const changes: string[] = [];
        
        if (original.length > this.MAX_LENGTH && sanitized.length === this.MAX_LENGTH) {
            changes.push(`Truncated from ${original.length} to ${this.MAX_LENGTH} chars`);
        }
        
        if (this.INVALID_CHARS.test(original) && !this.INVALID_CHARS.test(sanitized)) {
            changes.push('Replaced invalid characters');
        }
        
        if ((original.startsWith("'") || original.endsWith("'")) && 
            (!sanitized.startsWith("'") && !sanitized.endsWith("'"))) {
            changes.push('Removed leading/trailing quotes');
        }
        
        if (original !== sanitized && changes.length === 0) {
            changes.push('Other sanitization');
        }
        
        return changes;
    }

    /**
     * Handle sheet name references in formulas
     * Ensures that sheet names with special characters are properly quoted in formulas
     */
    static formatSheetNameForFormula(sheetName: string): string {
        if (!sheetName) return sheetName;
        
        // If sheet name contains spaces or special characters, wrap in single quotes
        const needsQuoting = /[ >>><!@#$%^&()+={}|;",.<>]/.test(sheetName);
        
        if (needsQuoting && !sheetName.startsWith("'") && !sheetName.endsWith("'")) {
            return `'${sheetName}'`;
        }
        
        return sheetName;
    }

    /**
     * Debug method to analyze sheet names
     */
    static analyzeSheetName(name: string): any {
        return {
            original: name,
            isValid: this.isValidSheetName(name),
            sanitized: this.sanitizeSheetName(name),
            excelSafe: this.createExcelSafeSheetName(name),
            length: name?.length || 0,
            hasInvalidChars: name ? this.INVALID_CHARS.test(name) : false,
            hasSpecialChars: name ? /[>>><!@#$%^&()+={}|;",.<>]/.test(name) : false
        };
    }
}