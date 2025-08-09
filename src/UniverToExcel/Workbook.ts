import exceljs from "@zwight/exceljs";
import { debug } from '../utils/debug';
import { jsonParse } from "../common/method";
import { createEnhancedWorksheet } from "./EnhancedWorkSheet";
// SHEET_HYPER_LINK_PLUGIN
// SHEET_DRAWING_PLUGIN
// SHEET_DEFINED_NAME_PLUGIN
// SHEET_CONDITIONAL_FORMATTING_PLUGIN
// SHEET_DATA_VALIDATION_PLUGIN
// SHEET_FILTER_PLUGIN
const Workbook = exceljs.Workbook;
export class WorkBook extends Workbook {
    constructor(snapshot: any) {
        super()
        this.init(snapshot);
    }

    private init(snapshot: any) {
        // this.properties.date1904 = true;
        this.calcProperties.fullCalcOnLoad = true;

        debug.log('🚀 [WorkBook] Using enhanced export');
        createEnhancedWorksheet(this, snapshot);
        
        this.setDefineNames(snapshot.resources);
    }

    private setDefineNames(resources: any[]) {
        if (!resources || !Array.isArray(resources)) {
            return;
        }
        
        const definedNamesResource = resources.find(d => d.name === 'SHEET_DEFINED_NAME_PLUGIN');
        if (!definedNamesResource) {
            return;
        }
        
        const definedNames = jsonParse(definedNamesResource?.data);
        if (!definedNames || typeof definedNames !== 'object') {
            return;
        }
        
        // DEBUG: Excel export error
        debug.log('[DEBUG] Export - Processing defined names:', Object.keys(definedNames).length);
        
        for (const key in definedNames) {
            const element = definedNames[key];
            if (!element || !element.name || !element.formulaOrRefString) {
                debug.log('[DEBUG] Export - Skipping invalid defined name:', element);
                continue;
            }
            
            try {
                // Validate the formula/reference string
                const formula = element.formulaOrRefString;
                
                debug.log('[DEBUG] Export - Processing defined name:', element.name, '=', formula);
                
                // Only skip if formula contains actual Excel errors (not just suspicious strings)
                if (formula.includes('#REF!') || formula.includes('#NAME?')) {
                    debug.log('[DEBUG] Export - Skipping invalid defined name with Excel error:', element.name, formula);
                    continue;
                }
                
                // Validate that name and formula are not empty
                if (!element.name.trim() || !formula.trim()) {
                    debug.log('[DEBUG] Export - Skipping empty defined name:', element.name, formula);
                    continue;
                }
                
                // Fix defined names: ExcelJS expects formulas WITHOUT the leading =
                let cleanFormula = formula;
                if (cleanFormula.startsWith('=')) {
                    cleanFormula = cleanFormula.substring(1);
                }
                
                // Add the defined name (ExcelJS: add(name, formula))
                debug.log('[DEBUG] Export - About to add defined name, definedNames type:', typeof this.definedNames);
                debug.log('[DEBUG] Export - definedNames object:', this.definedNames);
                debug.log('[DEBUG] Export - Calling add with:', element.name, cleanFormula);
                
                this.definedNames.add(element.name, cleanFormula);
                
                debug.log('[DEBUG] Export - Successfully added defined name:', element.name);
                debug.log('[DEBUG] Export - definedNames.model after add:', this.definedNames.model);
            } catch (error) {
                debug.log('[DEBUG] Export - Error adding defined name:', element.name, error);
                // Don't throw, continue with other defined names
            }
        }
    }
    // private setSheetProtection(snapshot: any) {
    //     const { resources, id } = snapshot
    //     const sheetProtections = jsonParse(resources.find((d: any) => d.name === 'SHEET_WORKSHEET_PROTECTION_PLUGIN')?.data);
    //     const protection = sheetProtections?.[id];
    //     if (!protection) return;

    // }
}