/**
 * Complete Resource Handlers for Excel Export
 * 
 * This file implements all resource plugin exports to ensure
 * complete feature parity with the import process.
 */

import { Worksheet } from '@zwight/exceljs';
import { debug } from '../utils/debug';
import { jsonParse } from '../common/method';
import { RESOURCE_PLUGINS } from './constants';
import { convertUniverRangeToExcel, convertUniverColorToExcel, fromBooleanNumber } from './univerUtils';

/**
 * Export filter data to Excel worksheet
 * Handles SHEET_FILTER_PLUGIN resource
 */
export function exportFilter(worksheet: Worksheet, sheetId: string, resources: any[]): void {
    if (!resources || !Array.isArray(resources)) return;
    
    const filterResource = resources.find(r => r.name === RESOURCE_PLUGINS.FILTER);
    if (!filterResource) {
        debug.log('🔍 [Filter] No filter data found');
        return;
    }
    
    const filterData = jsonParse(filterResource.data);
    if (!filterData || !filterData[sheetId]) {
        debug.log('🔍 [Filter] No filter for sheet:', sheetId);
        return;
    }
    
    const sheetFilter = filterData[sheetId];
    debug.log('🔍 [Filter] Applying filter:', sheetFilter);
    
    try {
        // Convert Univer range to Excel format
        if (sheetFilter.ref) {
            const range = convertUniverRangeToExcel(sheetFilter.ref);
            if (range) {
                worksheet.autoFilter = range;
                debug.log('✅ [Filter] Applied autoFilter to range:', range);
            }
        }
        
        // Apply column filters if present
        if (sheetFilter.filters && Array.isArray(sheetFilter.filters)) {
            for (const filter of sheetFilter.filters) {
                applyColumnFilter(worksheet, filter);
            }
        }
    } catch (error) {
        debug.error('❌ [Filter] Error applying filter:', error);
    }
}

/**
 * Apply column-specific filter
 */
function applyColumnFilter(worksheet: Worksheet, filter: any): void {
    if (!filter || !worksheet.autoFilter) return;
    
    // ExcelJS autoFilter columns are accessible after setting autoFilter
    // This is a simplified implementation - ExcelJS has limited filter API
    debug.log('🔍 [Filter] Column filter data:', filter);
    
    // Note: ExcelJS doesn't fully support complex filters
    // This would need enhancement based on specific requirements
}

/**
 * Export conditional formatting to Excel worksheet
 * Handles SHEET_CONDITIONAL_FORMATTING_PLUGIN resource
 */
export function exportConditionalFormatting(worksheet: Worksheet, sheetId: string, resources: any[]): void {
    if (!resources || !Array.isArray(resources)) return;
    
    const cfResource = resources.find(r => r.name === RESOURCE_PLUGINS.CONDITIONAL_FORMAT);
    if (!cfResource) {
        debug.log('🎨 [ConditionalFormat] No conditional formatting data found');
        return;
    }
    
    const cfData = jsonParse(cfResource.data);
    if (!cfData || !cfData[sheetId]) {
        debug.log('🎨 [ConditionalFormat] No conditional formatting for sheet:', sheetId);
        return;
    }
    
    const sheetCF = cfData[sheetId];
    debug.log('🎨 [ConditionalFormat] Processing rules:', sheetCF);
    
    if (!Array.isArray(sheetCF)) return;
    
    for (const rule of sheetCF) {
        try {
            const excelRule = convertConditionalRule(rule);
            if (excelRule && excelRule.ref) {
                worksheet.addConditionalFormatting(excelRule);
                debug.log('✅ [ConditionalFormat] Applied rule:', excelRule);
            }
        } catch (error) {
            debug.error('❌ [ConditionalFormat] Error applying rule:', error);
        }
    }
}

/**
 * Convert Univer conditional formatting rule to Excel format
 */
function convertConditionalRule(rule: any): any {
    if (!rule || !rule.ranges) return null;
    
    const excelRule: any = {
        ref: convertUniverRangeToExcel(rule.ranges),
        rules: []
    };
    
    // Process each condition
    if (rule.cfList && Array.isArray(rule.cfList)) {
        for (const cf of rule.cfList) {
            const convertedRule = convertSingleCondition(cf);
            if (convertedRule) {
                excelRule.rules.push(convertedRule);
            }
        }
    }
    
    return excelRule.rules.length > 0 ? excelRule : null;
}

/**
 * Convert single conditional formatting condition
 */
function convertSingleCondition(cf: any): any {
    if (!cf) return null;
    
    const rule: any = {
        type: mapConditionType(cf.type),
        priority: cf.priority || 1
    };
    
    // Style formatting
    if (cf.style) {
        rule.style = {
            fill: cf.style.bg ? {
                type: 'pattern',
                pattern: 'solid',
                fgColor: convertUniverColorToExcel(cf.style.bg)
            } : undefined,
            font: {
                color: cf.style.cl ? convertUniverColorToExcel(cf.style.cl) : undefined,
                bold: fromBooleanNumber(cf.style.bl),
                italic: fromBooleanNumber(cf.style.it)
            },
            border: cf.style.bd ? convertBorderForCF(cf.style.bd) : undefined
        };
    }
    
    // Condition-specific parameters
    switch (cf.type) {
        case 'cellIs':
            rule.operator = cf.operator; // between, equal, greaterThan, etc.
            rule.formulae = cf.value ? [cf.value] : [];
            break;
            
        case 'text':
            rule.operator = cf.operator; // containsText, notContains, etc.
            rule.text = cf.value;
            break;
            
        case 'colorScale':
            rule.cfvo = cf.colorScale?.cfvo || [];
            rule.color = cf.colorScale?.colors || [];
            break;
            
        case 'dataBar':
            rule.cfvo = cf.dataBar?.cfvo || [];
            rule.color = cf.dataBar?.color;
            break;
            
        case 'iconSet':
            rule.iconSet = cf.iconSet?.iconSet;
            rule.cfvo = cf.iconSet?.cfvo || [];
            break;
            
        case 'expression':
            rule.formulae = cf.formula ? [cf.formula] : [];
            break;
    }
    
    return rule;
}

/**
 * Map Univer condition type to Excel
 */
function mapConditionType(type: string): string {
    const typeMap: { [key: string]: string } = {
        'cellIs': 'cellIs',
        'text': 'containsText',
        'colorScale': 'colorScale',
        'dataBar': 'dataBar',
        'iconSet': 'iconSet',
        'expression': 'expression',
        'duplicateValues': 'duplicateValues',
        'top10': 'top10',
        'aboveAverage': 'aboveAverage'
    };
    
    return typeMap[type] || 'expression';
}

/**
 * Convert border for conditional formatting
 */
function convertBorderForCF(border: any): any {
    if (!border) return undefined;
    
    const result: any = {};
    
    if (border.t) result.top = { style: 'thin', color: convertUniverColorToExcel(border.t.cl) };
    if (border.b) result.bottom = { style: 'thin', color: convertUniverColorToExcel(border.b.cl) };
    if (border.l) result.left = { style: 'thin', color: convertUniverColorToExcel(border.l.cl) };
    if (border.r) result.right = { style: 'thin', color: convertUniverColorToExcel(border.r.cl) };
    
    return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * Export data validation to Excel worksheet
 * Handles SHEET_DATA_VALIDATION_PLUGIN resource
 */
export function exportDataValidation(worksheet: Worksheet, sheetId: string, resources: any[]): void {
    if (!resources || !Array.isArray(resources)) return;
    
    const dvResource = resources.find(r => r.name === RESOURCE_PLUGINS.DATA_VALIDATION);
    if (!dvResource) {
        debug.log('✓ [DataValidation] No data validation found');
        return;
    }
    
    const dvData = jsonParse(dvResource.data);
    if (!dvData || !dvData[sheetId]) {
        debug.log('✓ [DataValidation] No data validation for sheet:', sheetId);
        return;
    }

    const sheetDV = dvData[sheetId];
    debug.log('✓ [DataValidation] Processing validations:', sheetDV);
    
    if (!Array.isArray(sheetDV)) return;
    
    for (const validation of sheetDV) {
        try {
            applyDataValidation(worksheet, validation);
        } catch (error) {
            debug.error('❌ [DataValidation] Error applying validation:', error);
        }
    }
}

/**
 * Apply data validation to worksheet
 */
function applyDataValidation(worksheet: Worksheet, validation: any): void {
    if (!validation || !validation.ranges) return;

    let validationRanges = validation.ranges;
    if(!Array.isArray(validationRanges)) { validationRanges = [validationRanges]; }

    for (const validationRange of validationRanges) {

        const range = convertUniverRangeToExcel(validationRange);
        if (!range) continue;

        const excelValidation: any = {
            type: mapValidationType(validation.type),
            showErrorMessage: validation.showErrorMessage !== false,
            showInputMessage: validation.showInputMessage !== false,
            error: validation.errorMessage || 'Invalid value',
            errorTitle: validation.errorTitle || 'Error',
            prompt: validation.inputMessage || '',
            promptTitle: validation.inputTitle || ''
        };

        // Set validation-specific properties
        switch (validation.type) {
            case 'list':
                debugger
                if (validation.formula1) {
                    // Formula reference
                    if(validation.formula1.startsWith('=')){
                        excelValidation.formulae = [validation.formula1];
                    }
                    else {
                        excelValidation.formulae = [`"${validation}"`];
                    }
                } else if (validation.list) {
                    // Direct list
                    excelValidation.formulae = [`"${validation.list.join(',')}"`];
                }
                excelValidation.allowBlank = validation.allowBlank !== false;
                break;

            case 'whole':
            case 'decimal':
                excelValidation.operator = validation.operator || 'between';
                excelValidation.formulae = [];
                if (validation.formula1) excelValidation.formulae.push(validation.formula1);
                if (validation.formula2) excelValidation.formulae.push(validation.formula2);
                excelValidation.allowBlank = validation.allowBlank !== false;
                break;

            case 'date':
            case 'time':
                excelValidation.operator = validation.operator || 'between';
                excelValidation.formulae = [];
                if (validation.formula1) excelValidation.formulae.push(convertDateFormula(validation.formula1));
                if (validation.formula2) excelValidation.formulae.push(convertDateFormula(validation.formula2));
                break;

            case 'textLength':
                excelValidation.operator = validation.operator || 'between';
                excelValidation.formulae = [];
                if (validation.formula1) excelValidation.formulae.push(validation.formula1);
                if (validation.formula2) excelValidation.formulae.push(validation.formula2);
                break;

            case 'custom':
                excelValidation.formulae = validation.formula1 ? [validation.formula1] : [];
                break;
        }

        // Apply to cells in range
        const cells = worksheet.getCell(range);
        if (cells) {
            cells.dataValidation = excelValidation;
            debug.log('✅ [DataValidation] Applied to range:', range, excelValidation);
        }
    }
}

/**
 * Map Univer validation type to Excel
 */
function mapValidationType(type: string): string {
    const typeMap: { [key: string]: string } = {
        'list': 'list',
        'whole': 'whole',
        'decimal': 'decimal',
        'date': 'date',
        'time': 'time',
        'textLength': 'textLength',
        'custom': 'custom'
    };
    
    return typeMap[type] || 'custom';
}

/**
 * Convert date formula for validation
 */
function convertDateFormula(formula: any): any {
    if (typeof formula === 'number') {
        // Excel date serial number
        const date = new Date((formula - 25569) * 86400 * 1000);
        return date.toISOString().split('T')[0];
    }
    return formula;
}

/**
 * Export comments to Excel worksheet
 * Handles SHEET_COMMENT_PLUGIN resource
 */
export function exportComments(worksheet: Worksheet, sheetId: string, resources: any[]): void {
    if (!resources || !Array.isArray(resources)) return;
    
    const commentResource = resources.find(r => r.name === RESOURCE_PLUGINS.COMMENT);
    if (!commentResource) {
        debug.log('💬 [Comments] No comments found');
        return;
    }
    
    const commentData = jsonParse(commentResource.data);
    if (!commentData || !commentData[sheetId]) {
        debug.log('💬 [Comments] No comments for sheet:', sheetId);
        return;
    }
    
    const sheetComments = commentData[sheetId];
    debug.log('💬 [Comments] Processing comments:', sheetComments);
    
    if (!Array.isArray(sheetComments)) return;
    
    for (const comment of sheetComments) {
        try {
            if (comment.row !== undefined && comment.column !== undefined) {
                const cell = worksheet.getCell(comment.row + 1, comment.column + 1);
                // ExcelJS Comment interface doesn't have author field
                // We'll include author info in the text content
                const authorPrefix = comment.author ? `[${comment.author}]: ` : '';
                const fullContent = `${authorPrefix}${comment.content || ''}`;
                
                cell.note = {
                    texts: [{ text: fullContent }]
                };
                debug.log('✅ [Comments] Added comment to cell:', `${comment.column},${comment.row}`);
            }
        } catch (error) {
            debug.error('❌ [Comments] Error adding comment:', error);
        }
    }
}

/**
 * Export all resources for a worksheet
 * This is the main entry point for resource export
 */
export function exportAllResources(worksheet: Worksheet, sheetId: string, resources: any[]): void {
    debug.log('📦 [Resources] Exporting all resources for sheet:', sheetId);
    
    // Export in order of dependency
    exportFilter(worksheet, sheetId, resources);
    exportConditionalFormatting(worksheet, sheetId, resources);
    exportDataValidation(worksheet, sheetId, resources);
    exportComments(worksheet, sheetId, resources);
    
    debug.log('✅ [Resources] Completed resource export for sheet:', sheetId);
}