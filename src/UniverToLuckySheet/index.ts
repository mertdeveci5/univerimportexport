/**
 * UniverToLuckySheet - Phase 1 of the Export Implementation
 * 
 * This module converts Univer format back to LuckySheet format,
 * which is the exact reverse of LuckyToUniver.
 * 
 * This enables us to reuse the existing XML generation patterns
 * from the import process but in reverse.
 */

import type { IWorkbookData } from '@univerjs/core';
import { ILuckyFile } from '../ToLuckySheet/ILuck';
// import { UniverToLuckyWorkBook } from './LuckyWorkBook';
import { debug } from '../utils/debug';

export class UniverToLuckySheet {
    constructor(private univerData: IWorkbookData) {}

    /**
     * Convert Univer format to LuckySheet format
     * This is the main entry point for Phase 1
     */
    convert(): ILuckyFile {
        debug.log('🔄 [UniverToLuckySheet] Starting conversion from Univer to LuckySheet');
        debug.log('🔄 [UniverToLuckySheet] Input data:', {
            id: this.univerData.id,
            name: this.univerData.name,
            sheetCount: this.univerData.sheetOrder?.length || 0,
            hasStyles: !!this.univerData.styles,
            hasResources: !!this.univerData.resources
        });

        // const converter = new UniverToLuckyWorkBook(this.univerData);
        // const luckyFile = converter.convert();
        throw new Error('UniverToLuckySheet conversion not implemented - using direct ExcelJS approach instead');

        // This code is unreachable due to the throw above
        // debug.log('✅ [UniverToLuckySheet] Conversion complete:', {
        //     fileName: luckyFile.info.name,
        //     sheetCount: luckyFile.sheets.length,
        //     sheetNames: luckyFile.sheets.map(s => s.name)
        // });

        // return luckyFile;
    }
}

// Re-export main classes for external use
// NOTE: These imports are commented out because we chose the direct ExcelJS approach
// export { UniverToLuckyWorkBook } from './LuckyWorkBook';
// export { UniverToLuckySheetData } from './LuckySheetData';
// export { UniverToLuckyCell } from './LuckyCell';
// export { UniverToLuckyFormula } from './LuckyFormula';
// export { UniverToLuckyStyle } from './LuckyStyle';
// export { UniverToLuckyResources } from './LuckyResources';