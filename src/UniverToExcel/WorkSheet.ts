import { Workbook, Worksheet, WorksheetViewCommon, WorksheetViewFrozen } from "@zwight/exceljs";
import { convertSheetIdToName, heightConvert, hex2argb, wdithConvert } from "./util";
import { cellStyle, fontConvert } from "./CellStyle";
import { jsonParse, removeEmptyAttr } from "../common/method";
import { Resource } from "./Resource";

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
        console.warn('[ExcelWorkSheet] No sheetOrder found in snapshot');
        return;
    }
    sheetOrder.forEach((sheetId: string) => {
        const sheet = sheets[sheetId];
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
        const worksheet = workbook.addWorksheet(name, {
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
    for (const rowid in cellData) {
        const row = cellData[rowid];
        for (const columnid in row) {
            const cell = row[columnid];
            if (!cell) continue;
            // console.log(rowid + 1, columnid + 1)
            const target = worksheet.getCell(Number(rowid) + 1, Number(columnid) + 1)

            const valueFromHandle = handleValue(cell, {
                resources,
                sheetId: id,
                rowId: rowid,
                columnId: columnid,
                sheets
            }, workbook );
            
            console.log('[DEBUG] Export - Cell', Number(rowid) + 1, Number(columnid) + 1, 'value from handleValue:', JSON.stringify(valueFromHandle));
            target.value = valueFromHandle;
            
            // Post-process: Remove @ symbols that ExcelJS adds to named ranges
            // This happens after ExcelJS processes the formula
            if (target.value && typeof target.value === 'object' && 'formula' in target.value) {
                const originalFormula = target.value.formula;
                
                // Remove incorrectly placed @ symbols
                // ExcelJS sometimes adds @ symbols where they shouldn't be:
                // 1. Before function names: @TRANSPOSE -> TRANSPOSE  
                // 2. Before cell references in functions: (@$N$43:$N$45) -> ($N$43:$N$45)
                // 3. Before named ranges: @circ -> circ
                let cleanedFormula = originalFormula;
                
                // Fix: Remove @ from function names (like @TRANSPOSE)
                cleanedFormula = cleanedFormula.replace(/@([A-Z][A-Z0-9_]*)\s*\(/g, '$1(');
                
                // Fix: Remove @ from cell references (like @$N$43 or @A1)
                cleanedFormula = cleanedFormula.replace(/@(\$?[A-Z]+\$?\d+)/g, '$1');
                
                // Fix: Remove @ from named ranges (but not structured references like @[Column])
                cleanedFormula = cleanedFormula.replace(/@([A-Za-z_][A-Za-z0-9_]*)\b(?![\[])/g, '$1');
                
                // Debug log
                if (originalFormula !== cleanedFormula) {
                    console.log('[DEBUG] Export - Cleaned @ symbols in formula:', originalFormula, '->', cleanedFormula);
                }
                
                // Create new value object with cleaned formula
                // IMPORTANT: Also ensure no leading = to prevent double equals issue
                if (cleanedFormula.startsWith('=')) {
                    cleanedFormula = cleanedFormula.substring(1);
                    console.log('[DEBUG] Export - Also stripped leading = from @ cleaned formula');
                }
                
                target.value = {
                    formula: cleanedFormula,
                    result: target.value.result
                };
            }
            
            let originStyle = cell.s;
            if (typeof cell.s === 'string') {
                originStyle = styles[cell.s]
            }
            const style = removeEmptyAttr(cellStyle(originStyle, originStyle?.n?.pattern || cell.f))
            Object.assign(target, style)
            // console.log(target)
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

function handleValue(cell: any, cellSource: any, workbook: Workbook) {
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
        // Validate shared formula
        if (typeof cell.si === 'string' && !cell.si.includes('#REF!') && !cell.si.includes('#NAME?')) {
            // Fix double equals issue: ExcelJS expects formulas WITHOUT the leading =
            let formula = cell.si;
            if (formula.startsWith('=')) {
                formula = formula.substring(1);
            }
            console.log('[DEBUG] Export - Processing shared formula:', cell.si, '->', formula);
            value = { formula: formula, result: cell.v }
        } else {
            console.log('[DEBUG] Export - Skipping invalid shared formula:', cell.si);
            value = cell.v || '';
        }
    } else if (cell.f) {
        // Handle regular formulas (not just shared formulas)
        // Validate formula before adding
        if (typeof cell.f === 'string' && !cell.f.includes('#REF!') && !cell.f.includes('#NAME?')) {
            // Fix double equals issue: ExcelJS expects formulas WITHOUT the leading =
            let formula = cell.f;
            console.log('[DEBUG] Export - Original formula from Univer:', formula);
            if (formula.startsWith('=')) {
                formula = formula.substring(1);
                console.log('[DEBUG] Export - Stripped leading = from formula');
            }
            console.log('[DEBUG] Export - Final formula to ExcelJS:', formula);
            value = { formula: formula, result: cell.v }
        } else {
            console.log('[DEBUG] Export - Skipping invalid formula:', cell.f);
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