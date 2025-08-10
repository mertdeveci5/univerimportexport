import JSZip from '@progress/jszip-esm';
import { debug } from '../utils/debug';

/**
 * Post-processor for Excel files to fix issues that ExcelJS doesn't handle properly
 * This includes:
 * 1. Array formulas (TRANSPOSE, FILTER, etc.)
 * 2. Defined names
 * 3. Border styles (future)
 */
export class ExcelPostProcessor {
    private zip: JSZip;
    private definedNames: any[] = [];
    
    constructor() {
        this.zip = new JSZip();
    }
    
    /**
     * Process an Excel buffer to fix known issues
     */
    async processExcelBuffer(buffer: Buffer, univerData: any): Promise<Buffer> {
        debug.log('🔧 [PostProcessor] Starting Excel post-processing');
        
        // Load the Excel file
        await this.zip.loadAsync(buffer);
        
        // Fix array formulas
        await this.fixArrayFormulas();
        
        // Add defined names
        await this.addDefinedNames(univerData);
        
        // Fix border styles (TODO)
        // await this.fixBorderStyles(univerData);
        
        // Generate the fixed buffer
        const fixedBuffer = await this.zip.generateAsync({ 
            type: 'uint8array',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
        
        debug.log('✅ [PostProcessor] Post-processing complete');
        return Buffer.from(fixedBuffer);
    }
    
    /**
     * Fix array formulas by adding t="array" and ref attributes
     */
    private async fixArrayFormulas(): Promise<void> {
        debug.log('🔧 [PostProcessor] Fixing array formulas');
        
        // Get all worksheet files
        const worksheetFiles = Object.keys(this.zip.files)
            .filter(f => f.match(/xl\/worksheets\/sheet\d+\.xml/));
        
        for (const sheetPath of worksheetFiles) {
            let sheetXml = await this.zip.file(sheetPath)!.async('string');
            let modified = false;
            
            // Array formula patterns to detect
            const arrayFunctions = [
                'TRANSPOSE',
                'FILTER',
                'SORT',
                'SORTBY',
                'UNIQUE',
                'SEQUENCE',
                'RANDARRAY'
            ];
            
            // Create regex pattern for array functions
            const pattern = new RegExp(
                `<c([^>]*?)><f>(.*?(?:${arrayFunctions.join('|')}).*?)</f>`,
                'g'
            );
            
            sheetXml = sheetXml.replace(pattern, (match, attrs, formula) => {
                // Extract cell reference
                const cellMatch = attrs.match(/r="([A-Z]+)(\d+)"/);
                if (!cellMatch) return match;
                
                const col = cellMatch[1];
                const row = cellMatch[2];
                
                // Determine spill range based on formula
                const spillRange = this.determineSpillRange(formula, col, row);
                
                debug.log(`📐 [PostProcessor] Found array formula at ${col}${row}: ${formula.substring(0, 50)}`);
                debug.log(`   Adding: t="array" ref="${spillRange}"`);
                
                modified = true;
                return `<c${attrs}><f t="array" ref="${spillRange}">${formula}</f>`;
            });
            
            if (modified) {
                this.zip.file(sheetPath, sheetXml);
                debug.log(`✅ [PostProcessor] Fixed array formulas in ${sheetPath}`);
            }
        }
    }
    
    /**
     * Determine the spill range for an array formula
     */
    private determineSpillRange(formula: string, col: string, row: string): string {
        // This is a simplified version - in production, you'd need to:
        // 1. Parse the formula to understand the input range
        // 2. Calculate the output dimensions
        // 3. Handle different orientations (horizontal vs vertical)
        
        // For TRANSPOSE, swap dimensions
        if (formula.includes('TRANSPOSE')) {
            // Try to extract the range
            const rangeMatch = formula.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
            if (rangeMatch) {
                const startCol = rangeMatch[1];
                const startRow = parseInt(rangeMatch[2]);
                const endCol = rangeMatch[3];
                const endRow = parseInt(rangeMatch[4]);
                
                // Calculate dimensions
                const inputRows = endRow - startRow + 1;
                const inputCols = this.columnToNumber(endCol) - this.columnToNumber(startCol) + 1;
                
                // TRANSPOSE swaps dimensions
                const outputCols = inputRows;
                const outputRows = inputCols;
                
                // Calculate end cell
                const endColNum = this.columnToNumber(col) + outputCols - 1;
                const endRowNum = parseInt(row) + outputRows - 1;
                
                return `${col}${row}:${this.numberToColumn(endColNum)}${endRowNum}`;
            }
        }
        
        // Default to single cell if we can't determine
        return `${col}${row}:${col}${row}`;
    }
    
    /**
     * Add defined names to workbook.xml
     */
    private async addDefinedNames(univerData: any): Promise<void> {
        debug.log('🔧 [PostProcessor] Adding defined names');
        
        // Check if Univer data has defined names
        const definedNames = univerData.namedRanges || univerData.definedNames || [];
        
        if (!definedNames || (Array.isArray(definedNames) && definedNames.length === 0)) {
            debug.log('   No defined names to add');
            return;
        }
        
        // Get workbook.xml
        let workbookXml = await this.zip.file('xl/workbook.xml')!.async('string');
        
        // Check if definedNames section already exists
        if (workbookXml.includes('<definedNames>')) {
            debug.log('   definedNames section already exists');
            return;
        }
        
        // Build defined names XML
        let definedNamesXml = '    <definedNames>\n';
        
        if (Array.isArray(definedNames)) {
            definedNames.forEach((name: any) => {
                definedNamesXml += `        <definedName name="${name.name}">${name.ref}</definedName>\n`;
            });
        } else {
            // Handle object format
            Object.entries(definedNames).forEach(([name, ref]: [string, any]) => {
                const refString = typeof ref === 'object' ? ref.ref : ref;
                definedNamesXml += `        <definedName name="${name}">${refString}</definedName>\n`;
            });
        }
        
        definedNamesXml += '    </definedNames>';
        
        // Insert before </workbook>
        workbookXml = workbookXml.replace('</workbook>', definedNamesXml + '\n</workbook>');
        
        this.zip.file('xl/workbook.xml', workbookXml);
        debug.log(`✅ [PostProcessor] Added ${definedNames.length || Object.keys(definedNames).length} defined names`);
    }
    
    /**
     * Fix border styles (TODO: Implementation needed)
     */
    private async fixBorderStyles(univerData: any): Promise<void> {
        debug.log('🔧 [PostProcessor] Fixing border styles');
        // TODO: Implement border style fixing
        // This would involve:
        // 1. Reading styles from Univer data
        // 2. Merging with existing styles.xml
        // 3. Updating cell references to use correct style IDs
        debug.log('   Border fixing not yet implemented');
    }
    
    // Helper functions
    private columnToNumber(col: string): number {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
    }
    
    private numberToColumn(num: number): string {
        let col = '';
        while (num > 0) {
            const remainder = (num - 1) % 26;
            col = String.fromCharCode(65 + remainder) + col;
            num = Math.floor((num - 1) / 26);
        }
        return col;
    }
}