import { Workbook, Worksheet, WorksheetViewCommon, WorksheetViewFrozen } from "@zwight/exceljs";
import { debug } from '../utils/debug';
import { convertSheetIdToName, heightConvert, hex2argb, wdithConvert } from "./util";
import { cellStyle, fontConvert } from "./CellStyle";
import { jsonParse, removeEmptyAttr } from "../common/method";
import { Resource } from "./Resource";
import { ArrayFormulaHandler } from "./ArrayFormulaHandler";
import { FormulaCleaner } from "./FormulaCleaner";
import { SheetNameHandler } from "./SheetNameHandler";

export class ViewCommon implements WorksheetViewCommon{
    rightToLeft: boolean;
    activeCell: string;
    showRuler: boolean;
    showRowColHeaders: boolean;
    showGridLines: boolean;
    zoomScale: number;
    zoomScaleNormal: number;
}
export class FrozenView implements WorksheetViewFrozen{
    state: "frozen";
    style?: "pageBreakPreview";
    xSplit?: number;
    ySplit?: number;
    topLeftCell?: string;
}
export function ExcelWorkSheet(workbook: Workbook, snapshot: any) {
    const { sheetOrder, sheets, styles, resources } = snapshot;
    if (!sheetOrder || !Array.isArray(sheetOrder)) {
        debug.warn('[ExcelWorkSheet] No sheetOrder found in snapshot');
        return;
    }
    sheetOrder.forEach((sheetId: string) => {
        const sheet = sheets[sheetId];
        
        // Ensure sheet exists (important for empty sheets)
        if (!sheet) {
            debug.warn('⚠️ [ExcelWorkSheet] Missing sheet data for ID:', sheetId);
            return;
        }
        
        const { 
            id,
            name, 
            tabColor, 
            defaultRowHeight, 
            defaultColumnWidth, 
            hidden,
            rightToLeft,
            showGridlines,
            freeze,
            mergeData
        } = sheet;
        
        // Log sheet processing
        debug.log('📋 [ExcelWorkSheet] Processing sheet:', {
            sheetId: id,
            name: name,
            hidden: hidden === 1,
            hasCellData: !!(sheet.cellData && Object.keys(sheet.cellData).length > 0),
            hasArrayFormulas: !!(sheet.arrayFormulas && sheet.arrayFormulas.length > 0)
        });
        const commonView = new ViewCommon();
        commonView.rightToLeft = rightToLeft === 1;
        commonView.showGridLines = showGridlines === 1;
        const frozenView = new FrozenView();
        if (freeze && (freeze.xSplit > 0 || freeze.ySplit > 0)){
            frozenView.state = 'frozen';
            frozenView.xSplit = freeze.xSplit;
            frozenView.ySplit = freeze.ySplit;
        }
        const views = Object.assign(commonView, frozenView)

        const defaultColWidth = wdithConvert(defaultColumnWidth);
        const defaultRowHeightR = heightConvert(defaultRowHeight);
        // Sanitize sheet name for Excel compatibility
        const excelSafeName = SheetNameHandler.createExcelSafeSheetName(name);
        
        // Log sheet name analysis in debug mode
        if (name !== excelSafeName) {
            debug.log('📋 [SheetName] Analysis:', SheetNameHandler.analyzeSheetName(name));
        }
        
        const worksheet = workbook.addWorksheet(excelSafeName, {
            views: [views],
            state: hidden === 1 ? 'hidden' : 'visible',
            properties: {
                tabColor: tabColor ? { argb: hex2argb(tabColor) } : undefined,
                defaultColWidth: defaultColWidth,
                defaultRowHeight: defaultRowHeightR,
                dyDescent: 0
            }
        })
        setColumns(worksheet, sheet.columnData, defaultColWidth);
        setRows(worksheet, sheet.rowData, defaultRowHeightR)
        setCell(worksheet, sheet, styles, snapshot, workbook);
        setMerges(worksheet, mergeData);
        
        new Resource(id, workbook, worksheet, resources);
    });
}


function setMerges(worksheet: Worksheet, mergeData: any[]) {
    if (!mergeData || !Array.isArray(mergeData)) {
        return;
    }
    mergeData.forEach(d => {
        worksheet.mergeCells(d.startRow + 1, d.startColumn + 1, d.endRow + 1, d.endColumn + 1)
    })
}

function setCell(worksheet: Worksheet, sheet: any, styles: any, snapshot: any, workbook: Workbook) {
    const { resources, sheets } = snapshot;
    const { cellData, id } = sheet;
    const arrayHandler = new ArrayFormulaHandler();
    
    // Handle empty sheets - ensure they are properly created
    if (!cellData || Object.keys(cellData).length === 0) {
        debug.log('📄 [EmptySheet] Processing empty sheet:', {
            sheetId: id,
            sheetName: sheet.name,
            hasCellData: !!cellData,
            cellDataKeys: cellData ? Object.keys(cellData).length : 0
        });
        
        // For empty sheets, just ensure worksheet has minimum structure
        // ExcelJS will handle the empty sheet correctly
        return;
    }
    
    // Log array formulas for debugging
    if (sheet.arrayFormulas && sheet.arrayFormulas.length > 0) {
        debug.log('🔢 [ArrayFormula] Found array formulas in sheet:', {
            sheetId: id,
            sheetName: sheet.name,
            arrayFormulaCount: sheet.arrayFormulas.length,
            formulas: sheet.arrayFormulas.map((af: any) => ({
                formula: af.formula,
                range: `${af.range.startRow},${af.range.startCol}:${af.range.endRow},${af.range.endCol}`,
                formulaId: af.formulaId
            }))
        });
    }
    
    debug.log('📊 [CellData] Processing sheet data:', {
        sheetId: id,
        sheetName: sheet.name,
        rowCount: Object.keys(cellData).length,
        totalCells: Object.values(cellData).reduce((sum: number, row: any) => sum + Object.keys(row || {}).length, 0)
    });
    
    for (const rowid in cellData) {
        const row = cellData[rowid];
        if (!row) continue; // Skip empty rows
        
        for (const columnid in row) {
            const cell = row[columnid];
            if (!cell) continue; // Skip empty cells
            
            const rowIndex = Number(rowid);
            const colIndex = Number(columnid);
            const target = worksheet.getCell(rowIndex + 1, colIndex + 1);

            // Check if this is an array formula cell
            if (arrayHandler.isArrayFormula(cell, sheet, rowIndex, colIndex)) {
                const arrayFormulaInfo = arrayHandler.getArrayFormulaInfo(cell, sheet);
                
                if (arrayFormulaInfo && 
                    arrayFormulaInfo.masterRow === rowIndex && 
                    arrayFormulaInfo.masterCol === colIndex &&
                    !arrayHandler.isRangeProcessed(arrayFormulaInfo.range)) {
                    
                    // This is the master cell of an array formula
                    debug.log('🔢 [ArrayFormula] Processing master cell:', {
                        cell: `${colIndex},${rowIndex}`,
                        formula: arrayFormulaInfo.formula,
                        range: arrayFormulaInfo.range
                    });
                    
                    // Apply array formula to the entire range
                    const success = arrayHandler.applyArrayFormula(worksheet, arrayFormulaInfo, cell.v);
                    
                    if (success) {
                        // Skip normal value handling for array formula cells
                        // Apply styles only
                        let originStyle = cell.s;
                        if (typeof cell.s === 'string') {
                            originStyle = styles[cell.s];
                        }
                        const style = removeEmptyAttr(cellStyle(originStyle, originStyle?.n?.pattern || cell.f));
                        Object.assign(target, style);
                        continue;
                    } else {
                        debug.log('⚠️ [ArrayFormula] Failed to apply array formula, falling back to regular handling');
                    }
                } else {
                    // This is a dependent cell of an array formula - skip individual processing
                    debug.log('🔢 [ArrayFormula] Skipping dependent cell:', {
                        cell: `${colIndex},${rowIndex}`,
                        formulaId: cell.si
                    });
                    continue;
                }
            }

            // Regular cell handling
            const valueFromHandle = handleValue(cell, {
                resources,
                sheetId: id,
                rowId: rowid,
                columnId: columnid,
                sheets
            }, workbook, sheet);
            
            debug.log('[DEBUG] Export - Cell', rowIndex + 1, colIndex + 1, 'value from handleValue:', JSON.stringify(valueFromHandle));
            target.value = valueFromHandle;
            
            // Post-process: Clean formulas using FormulaCleaner
            if (target.value && typeof target.value === 'object' && 'formula' in target.value) {
                const originalFormula = target.value.formula;
                const cleanedFormula = FormulaCleaner.cleanFormula(originalFormula);
                
                if (cleanedFormula && cleanedFormula !== originalFormula) {
                    target.value = {
                        formula: cleanedFormula,
                        result: target.value.result
                    };
                } else if (!cleanedFormula) {
                    // If formula is invalid, fall back to the result value
                    debug.log('[DEBUG] Export - Invalid formula, using result value:', originalFormula);
                    target.value = target.value.result || '';
                }
            }
            
            let originStyle = cell.s;
            if (typeof cell.s === 'string') {
                originStyle = styles[cell.s]
            }
            const style = removeEmptyAttr(cellStyle(originStyle, originStyle?.n?.pattern || cell.f))
            Object.assign(target, style)
            // debug.log(target)
        }
    }
}
function getHyperLink(cellSource: any) {
    const { resources, sheetId, rowId, columnId } = cellSource;
    const hyperlinks = jsonParse(resources.find((d: any) => d.name === 'SHEET_HYPER_LINK_PLUGIN')?.data);
    const list = hyperlinks?.[sheetId] || [];
    const hyperlink = list.find((d: any) => d.row === Number(rowId) && d.column === Number(columnId));
    return hyperlink
}

function handleHyperLink(hyperlink: any, sheets: any) {
    let hyperlinks;
    if (hyperlink) {
        const { payload } = hyperlink;
        let link = '';
        let model = '';
        if (payload.includes('#gid=') || payload.includes('range=')) {
            const str = payload.replace('#', '');
            const arr = str.split('&');
            link += '';
            if (arr.length === 1 && arr[0].includes('range=')) {
                link += arr[0].replace('range=')
            }
            if (arr.length === 2) {
                link += `\'${convertSheetIdToName(sheets, arr[0].replace('gid=', ''))}\'`;
                link += `!${arr[1].replace('range=', '')}`
            }
        } else {
            link = payload
            model = 'External'
        }
        
        if (link) hyperlinks = {
            hyperlink: link,
            hyperlinkModel: model
        } 
    }
    return hyperlinks
}

function handleValue(cell: any, cellSource: any, workbook: Workbook, sheetData?: any) {
    const { sheets } = cellSource
    const hyperlink = getHyperLink(cellSource)
    const hyperlinks = handleHyperLink(hyperlink, sheets)
    let value;
    if (cell.p) {
        const body = cell.p?.body;
        if (cell.p.drawingsOrder?.length) {
            const image = cell.p.drawings[cell.p.drawingsOrder[0]];
            const { id, value: imgId } = workbook.addCellImage({
                base64: image.source,
                extension: 'png',
                descr: image.description,
                ext: {
                    width: image.transform.width,
                    height: image.transform.height,
                }
            })
            value = { id, cellImageId: imgId, ...(hyperlinks || {}) }
            return value
        } else {
            value = {
                richText: body?.textRuns.map((d: any) => {
                    return {
                        text: body.dataStream.substring(d.st, d.ed),
                        font: fontConvert(d.ts)
                    }
                })
            }
        }
    } else if (cell.si) {
        // Check if this is an array formula (handled separately)
        if (sheetData && sheetData.arrayFormulas) {
            const isArray = sheetData.arrayFormulas.some((af: any) => af.formulaId === cell.si);
            if (isArray) {
                debug.log('[DEBUG] Export - Array formula detected in handleValue, using value only:', cell.v);
                value = cell.v || '';
                return value; // Return early for array formulas
            }
        }
        
        // Validate and clean regular shared formula
        if (FormulaCleaner.isValidFormula(cell.si)) {
            const cleanedFormula = FormulaCleaner.cleanFormula(cell.si);
            if (cleanedFormula) {
                debug.log('[DEBUG] Export - Processing shared formula:', cell.si, '->', cleanedFormula);
                value = { formula: cleanedFormula, result: cell.v };
            } else {
                debug.log('[DEBUG] Export - Formula cleaned to empty, using value:', cell.v);
                value = cell.v || '';
            }
        } else {
            debug.log('[DEBUG] Export - Skipping invalid shared formula:', cell.si);
            value = cell.v || '';
        }
    } else if (cell.f) {
        // Handle regular formulas (not just shared formulas)
        // Validate and clean formula before adding
        if (FormulaCleaner.isValidFormula(cell.f)) {
            const cleanedFormula = FormulaCleaner.cleanFormula(cell.f);
            if (cleanedFormula) {
                debug.log('[DEBUG] Export - Processing regular formula:', cell.f, '->', cleanedFormula);
                value = { formula: cleanedFormula, result: cell.v };
            } else {
                debug.log('[DEBUG] Export - Formula cleaned to empty, using value:', cell.v);
                value = cell.v || '';
            }
        } else {
            debug.log('[DEBUG] Export - Skipping invalid formula:', cell.f);
            value = cell.v || '';
        }
    } else {
        value = cell.v
    }
    if (hyperlinks) {
        const text = value?.richText?.map?.((d: any) => d.text)?.join('') || value?.result || value;
        value = {
            text: text,
            ...hyperlinks
        }
    }
    return value
}

function setColumns(worksheet: Worksheet, columnData: any = {}, defaultColumnWidth: number) {
    for (const key in columnData) {
        if (Object.prototype.hasOwnProperty.call(columnData, key)) {
            const element = columnData[key];
            const column = worksheet.getColumn(Number(key) + 1)
            column.width = element.w ? wdithConvert(element.w) : defaultColumnWidth;
            column.hidden = element.hd === 1;
        }
    }
}
function setRows(worksheet: Worksheet, rowData: any = {}, defaultRowHeight: number) {
    for (const key in rowData) {
        if (Object.prototype.hasOwnProperty.call(rowData, key)) {
            const element = rowData[key];
            const row = worksheet.getRow(Number(key) + 1)
            row.height = element.h ? heightConvert(element.h) : defaultRowHeight;
            row.hidden = element.hd === 1;
        }
    }
}