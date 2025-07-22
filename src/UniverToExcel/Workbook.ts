import exceljs from "@zwight/exceljs";
import { jsonParse } from "../common/method";
import { ExcelWorkSheet } from "./WorkSheet";
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

        this.setDefineNames(snapshot.resources);
        ExcelWorkSheet(this, snapshot)
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
        console.log('[DEBUG] Export - Processing defined names:', Object.keys(definedNames).length);
        
        for (const key in definedNames) {
            const element = definedNames[key];
            if (!element || !element.name || !element.formulaOrRefString) {
                console.log('[DEBUG] Export - Skipping invalid defined name:', element);
                continue;
            }
            
            try {
                // Validate the formula/reference string
                const formula = element.formulaOrRefString;
                
                // Skip if it contains problematic patterns
                if (formula.includes('#REF!') || formula.includes('#NAME?') || formula.includes('SHEET_DEFINED_NAME')) {
                    console.log('[DEBUG] Export - Skipping problematic defined name:', element.name, formula);
                    continue;
                }
                
                this.definedNames.add(formula, element.name);
            } catch (error) {
                console.log('[DEBUG] Export - Error adding defined name:', element.name, error);
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