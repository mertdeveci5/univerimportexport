import {
    BooleanNumber,
    ICellData,
    IRange,
    IObjectMatrixPrimitiveType,
    IObjectArrayPrimitiveType,
    IRowData,
    IColumnData,
    CellValueType,
    Nullable,
    IDocumentData,
    PositionedObjectLayoutType,
    DrawingTypeEnum,
} from '@univerjs/core';
import { UniverSheetBase } from './UniverSheetBase';
import { handleStyle, removeEmptyAttr } from './utils';
import { str2num, generateRandomId } from '../common/method';
import { IluckySheet, IluckySheetConfig, IluckySheetCelldata, IluckysheetHyperlink, IluckysheetFrozen, IluckySheetCelldataValue } from '../ToLuckySheet/ILuck';
import { ImageSourceType } from './ILuckInterface';

export interface HyperLink {
    id: string;
    payload: string | { gid: string; range: string };
    row: number;
    column: number;
}

export interface ArrayFormula {
    formula: string;
    range: {startRow: number, endRow: number, startCol: number, endCol: number};
    masterRow: number;
    masterCol: number;
}

export interface UniverSheetMode extends UniverSheetBase {
    hyperLink: HyperLink[];
    arrayFormulas: ArrayFormula[];
    mode: UniverSheetMode;
}
export class UniverSheet extends UniverSheetBase {
    hyperLink: HyperLink[] = [];
    arrayFormulas: ArrayFormula[] = [];
    constructor(sheetData: IluckySheet) {
        super();
        const {
            color,
            zoomRatio,
            celldata,
            config = {} as IluckySheetConfig,
            showGridLines,
            defaultColWidth,
            defaultRowHeight,
            hide,
        } = sheetData || {};
        this.name = sheetData.name;
        this.id = `sheet-${sheetData.index}`;
        if (sheetData) {
            this.tabColor = color;
            this.zoomRatio = zoomRatio;
            this.showGridlines = Number(showGridLines);
            this.defaultColumnWidth = defaultColWidth;
            this.defaultRowHeight = defaultRowHeight;
            this.hidden = hide;
            this.handleSheetLink(sheetData.hyperlink);

            if (config.merge) this.mergeData = this.handleMerge(config);

            // Always process sheet data, even if celldata is empty to preserve empty sheets
            const { cellData, rowCount, colCount } = this.handleCellData(celldata || [], config);
            this.cellData = cellData;
            this.rowCount = this.rowCount > rowCount ? this.rowCount : rowCount + 1;
            this.columnCount = this.columnCount > colCount ? this.columnCount : colCount + 1;
            console.log(this.rowCount, this.columnCount)
            this.handleRowAndColumnData(config);
            if (sheetData.freezen) this.handleFreeze(sheetData.freezen);
        }
    }
    get mode(): Omit<UniverSheetMode, 'mode'> {
        return {
            id: this.id,
            name: this.name,
            type: this.type,
            tabColor: this.tabColor,
            hidden: this.hidden,
            freeze: this.freeze,
            rowCount: this.rowCount,
            columnCount: this.columnCount,
            zoomRatio: this.zoomRatio,
            scrollTop: this.scrollTop,
            scrollLeft: this.scrollLeft,
            defaultColumnWidth: this.defaultColumnWidth,
            defaultRowHeight: this.defaultRowHeight,
            mergeData: this.mergeData,
            cellData: this.cellData,
            rowData: this.rowData,
            columnData: this.columnData,
            rowHeader: this.rowHeader,
            columnHeader: this.columnHeader,
            showGridlines: this.showGridlines,
            rightToLeft: this.rightToLeft,
            selections: this.selections,
            hyperLink: this.hyperLink,
            arrayFormulas: this.arrayFormulas,
        };
    }
    private handleMerge = (config: IluckySheetConfig): IRange[] => {
        const merges = config.merge;
        if (!merges) return [];
        return Object.values(merges).map((merge) => {
            return {
                startRow: merge.r,
                endRow: merge.r + merge.rs - 1,
                startColumn: merge.c,
                endColumn: merge.c + merge.cs - 1,
            };
        });
    };
    private handleCellData = (celldata: IluckySheetCelldata[], config: IluckySheetConfig) => {
        const handleCell = (row: IluckySheetCelldata): ICellData => {
            const { v } = row;
            if (typeof v === 'string' || v === null || v === undefined) {
                return { v: v as string };
            }
            const tMap: any = {
                s: CellValueType.STRING,
                n: CellValueType.NUMBER,
                b: CellValueType.BOOLEAN,
                str: CellValueType.STRING,
            };
            const borderConf = config.borderInfo?.find(
                (d) => d.value.col_index === row.c && d.value.row_index === row.r
            );

            let cellType = v.ct?.t && tMap[v.ct?.t] ? tMap[v.ct?.t] : CellValueType.NUMBER;

            let val = cellType === CellValueType.NUMBER ? str2num(v.v) : v.v;
            if (cellType === CellValueType.BOOLEAN) val = v.v == '1' ? 1 : 0;

            if (Number.isNaN(Number(val)) && cellType === CellValueType.NUMBER)
                cellType = CellValueType.STRING;
            // Check if cell has hyperlink
            const hyperlinkInfo = this.hyperLink.find((d) => d.column === row.c && d.row === row.r);
            if (hyperlinkInfo) {
                cellType = CellValueType.STRING;
            }

            // Handle formulas - preserve formula and calculated value
            const f = v.f?.replace(/=_xlfn./g, '=');
            
            // Check if this is an array formula
            const isArrayFormula = v.ft === 'array' && v.ref;
            let arrayFormulaRange = null;
            if (isArrayFormula && v.ref && f) {
                // Parse the array formula range (e.g., "A1:C3")
                arrayFormulaRange = this.parseRange(v.ref);
                if (arrayFormulaRange) {
                    // Store array formula information for later processing
                    this.arrayFormulas.push({
                        formula: f,
                        range: arrayFormulaRange,
                        masterRow: row.r,
                        masterCol: row.c
                    });
                }
            }
            
            const cell: ICellData = {
                // custom: v., // User stored custom fields
                f,
                // p: , // The unique key, a random string, is used for the plug-in to associate the cell. When the cell information changes, the plug-in does not need to change the data, reducing the pressure on the back-end interface id?: string.
                s: handleStyle(row, borderConf),
                // si: Handle shared formula ID if available from other sources
                t: cellType,
                v: val, // Always preserve the calculated value
            };
            // Handle hyperlinks first - convert to rich text format
            if (hyperlinkInfo) {
                const hyperlinkDoc = this.handleHyperlinkDocument(row, hyperlinkInfo);
                if (hyperlinkDoc) {
                    cell.p = hyperlinkDoc;
                }
            } else {
                const pVal = this.handleDocument(row, config);
                if (pVal) cell.p = pVal;

                const pValImg = this.handleCellImage(row, config);
                if (pValImg) {
                    cell.p = pValImg;
                    cell.f = undefined;
                    cell.v = undefined;
                }
            }
            return removeEmptyAttr(cell);
        };
        let row: number | undefined = undefined;
        let colCount = 0;
        const rowData = celldata.reduce((pre: any, cur) => {
            if (row === cur.r) {
                pre[cur.r].push(cur);
            } else {
                row = cur.r;
                pre[row] = [cur];
            }
            if (cur.c > colCount) colCount = cur.c;
            return pre;
        }, []);
        const cell: IObjectMatrixPrimitiveType<ICellData> = {};
        // console.log(rowData, celldata, colCount)
        rowData.forEach((row: IluckySheetCelldata[], rowIndex: number) => {
            for (let index = 0; index < colCount + 1; index++) {
                const element = row.find((d) => d.c === index) || {
                    r: rowIndex,
                    c: index,
                    v: null,
                };
                if (!cell[element.r]) cell[element.r] = {};
                cell[element.r][element.c] = handleCell(element);
            }
        });
        
        // Apply array formulas to all cells in their ranges
        this.applyArrayFormulas(cell);
        
        return {
            cellData: cell,
            rowCount: rowData.length,
            colCount,
        };
    };
    private handleDocument = (row: IluckySheetCelldata, config: IluckySheetConfig) => {
        const matchArray = (str: string, charToFind: string) => {
            const regex = new RegExp(charToFind, 'g');
            let match;
            const indices = [];

            while ((match = regex.exec(str))) {
                indices.push(match.index);
            }

            return indices;
        };
        const removeLastChar = (str: string, charToRemove: string) => {
            const regex = new RegExp(`${charToRemove}`, 'g');
            return str.replace(regex, '\r');
        };
        let pVlaue: Nullable<IDocumentData> = null;
        const { v } = row;
        if (typeof v === 'string' || v === null || v === undefined) {
            return undefined;
        }
        if (v.ct && v.ct.t === 'inlineStr') {

            v.ct.s = v.ct.s?.map(d => {
                d.v = removeLastChar(d.v || '', '\r\n');
                return d
            }) || []

            let dataStream = v.ct.s.reduce((prev, cur) => {
                return prev + cur.v;
            }, '');
            dataStream = dataStream ? dataStream.replace(/\n/g, '\r') + '\r\n' : '';
            const matchChart = {
                r: '\r', // PARAGRAPH
                n: '\n', // SECTION_BREAK
                v: '\v', // COLUMN_BREAK
                f: '\f', // PAGE_BREAK
                '0': '\0', // DOCS_END
                t: '\t', // TAB
                b: '\b', // customBlock
                x1A: '\x1A', // table start
                x1B: '\x1B', // table row start
                x1C: '\x1C', // table cell start
                x1D: '\x1D', // table cell end
                x1E: '\x1E', // table row end || customRange end
                x1F: '\x1F', // table end || customRange start
            };
            const paragraphs = matchArray(dataStream, matchChart.r).map((d) => {
                return {
                    startIndex: d,
                };
            });
            const sectionBreaks = matchArray(dataStream, matchChart.n).map((d) => {
                return {
                    startIndex: d,
                };
            });
            const textRuns = v.ct.s.map((d, index) => {
                const start = v.ct.s.reduce((prev, cur, curi) => {
                    if (curi < index) return prev + (cur.v?.length || 0);
                    return prev;
                }, 0);
                const end = start + (v.ct!.s?.[index]?.v?.length || 0);
                const borderConf = config.borderInfo?.find(
                    (d) => d.value.col_index === row.c && d.value.row_index === row.r
                );
                return {
                    st: start,
                    ed: end,
                    ts: handleStyle(
                        {
                            v: (v.ct!.s[index] || v.ct!.s[0]) as IluckySheetCelldataValue,
                            r: row.r,
                            c: row.c,
                        },
                        borderConf,
                        true
                    ),
                };
            });
            pVlaue = {
                id: generateRandomId(6),
                documentStyle: {
                    documentFlavor: 0,
                    pageSize: { width: 0, height: 0 },
                    renderConfig: {},
                    textStyle: {},
                },
                body: {
                    dataStream,
                    paragraphs,
                    sectionBreaks,
                    textRuns,
                },
                drawings: {},
            };
        }
        return pVlaue;
    };

    /**
     * Create rich text document for hyperlink cells
     * @param row Cell data
     * @param hyperlinkInfo Hyperlink information 
     */
    private handleHyperlinkDocument = (row: IluckySheetCelldata, hyperlinkInfo: HyperLink): Nullable<IDocumentData> => {
        const { v } = row;
        let cellText = '';
        
        // Get display text - prefer actual cell value over hyperlink address
        if (typeof v === 'string') {
            cellText = v;
        } else if (v && typeof v === 'object' && v.v) {
            cellText = String(v.v);
        } else {
            // Fallback to hyperlink address if no display text
            cellText = typeof hyperlinkInfo.payload === 'string' 
                ? hyperlinkInfo.payload 
                : (hyperlinkInfo.payload?.range || 'Link');
        }

        if (!cellText) return null;

        const linkText = cellText + '\r\n';
        const linkId = generateRandomId(6);
        
        return {
            id: generateRandomId(6),
            documentStyle: {
                documentFlavor: 0,
                pageSize: { width: 0, height: 0 },
                renderConfig: {},
                textStyle: {},
            },
            body: {
                dataStream: linkText,
                paragraphs: [{
                    startIndex: linkText.length - 2, // Before \r\n
                }],
                sectionBreaks: [{
                    startIndex: linkText.length - 1, // Before \n
                }],
                textRuns: [{
                    st: 0,
                    ed: cellText.length,
                    ts: {
                        cl: {
                            rgb: 'rgb(0, 0, 255)' // Blue color for hyperlinks
                        },
                        ul: {
                            s: 1 // Underline
                        }
                    }
                }],
                customRanges: [{
                    startIndex: 0,
                    endIndex: cellText.length,
                    rangeId: linkId,
                    rangeType: 2, // HYPERLINK_RANGE
                    properties: {
                        url: typeof hyperlinkInfo.payload === 'string' ? hyperlinkInfo.payload : undefined,
                        payload: typeof hyperlinkInfo.payload === 'object' ? hyperlinkInfo.payload : undefined
                    }
                }]
            },
            drawings: {},
        };
    };

    private handleCellImage = (row: IluckySheetCelldata, config: IluckySheetConfig) => {
        let pVlaue: Nullable<any> = null;
        const { v } = row;
        if (typeof v === 'string' || v === null || v === undefined) {
            return undefined;
        }
        if (v.ct && v.ct.t === 'str' && v.ct.ci) {
            const blockId = generateRandomId(6);
            const valueId = generateRandomId(6);
            const { default: defaultData,  src, descr } = v.ct.ci || {};
            const borderConf = config.borderInfo?.find(
                (d) => d.value.col_index === row.c && d.value.row_index === row.r
            );
            pVlaue = {
                id: valueId,
                documentStyle: {
                    documentFlavor: 0,
                    pageSize: { width: 0, height: 0 },
                    renderConfig: {},
                    textStyle: {},
                },
                body: {
                    dataStream: '\b\r\n',
                    paragraphs: [{
                        startIndex: 1,
                        paragraphStyle: { horizontalAlign: v.ht }
                    }],
                    sectionBreaks: [{ startIndex: 2 }],
                    textRuns: [{
                        ed: 1,
                        st: 0,
                        ts: handleStyle(
                            {
                                v: v,
                                r: row.r,
                                c: row.c,
                            },
                            borderConf,
                            true
                        ),
                    }],
                    customBlocks: [{ startIndex: 0, blockId }]
                },
                drawings: {
                    [blockId]: {
                        unitId: valueId,
                        subUnitId: valueId,
                        drawingId: blockId,
                        layoutType: PositionedObjectLayoutType.INLINE,
                        title: '',
                        description: descr,
                        docTransform: {
                            size: {
                                width: defaultData.width,
                                height: defaultData.height
                            },
                            positionH: {
                                relativeFrom: 0,
                                posOffset: 0
                            },
                            positionV: {
                                relativeFrom: 1,
                                posOffset: 0
                            },
                            angle: 0
                        },
                        drawingType: DrawingTypeEnum.DRAWING_IMAGE,
                        imageSourceType: ImageSourceType.BASE64,
                        source: src,
                        transform: defaultData
                    }
                },
                drawingsOrder: [blockId]
            };
        }
        return pVlaue;
    }

    private handleRowAndColumnData = (config: IluckySheetConfig) => {
        const columnData: IObjectArrayPrimitiveType<Partial<IColumnData>> = {};
        const rowData: IObjectArrayPrimitiveType<Partial<IRowData>> = {};
        for (let index = 0; index < this.rowCount; index++) {
            rowData[index] = {
                h: config.rowlen?.[index] || this.defaultRowHeight,
                ia: !config.rowlen?.[index] ? BooleanNumber.TRUE : BooleanNumber.FALSE,
                ah: this.defaultRowHeight,
                hd: config.rowhidden?.[index] === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            };
        }

        for (let index = 0; index < this.columnCount; index++) {
            columnData[index] = {
                w: config.columnlen?.[index] || this.defaultColumnWidth,
                hd: config.colhidden?.[index] === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            };
        }
        this.rowData = rowData;
        this.columnData = columnData;
    };

    /**
     * 处理链接
     * @param sheetName IluckysheetHyperlink
     */
    private handleSheetLink = (hyperlinks: IluckysheetHyperlink) => {
        if (!hyperlinks) return;
        const links = Object.keys(hyperlinks).map((d) => {
            const row = d.split('_')[0],
                column = d.split('_')[1];

            const item = hyperlinks[d];
            let payload: any = item.linkAddress;
            if (item.linkType === 'internal') {
                const locationList = item.linkAddress.split('!');
                payload = {};
                if (locationList[0]) payload['gid'] = locationList[0];
                if (locationList[1]) payload['range'] = locationList[1];
            }
            return {
                id: generateRandomId(6),
                row: Number(row),
                column: Number(column),
                payload,
            };
        });
        this.hyperLink = links;
    };
    
    private handleFreeze = (freeze: IluckysheetFrozen) => {
        this.freeze = {
            xSplit: freeze.vertical,
            ySplit: freeze.horizen,
            startColumn: freeze.vertical,
            startRow: freeze.horizen,
        };
    };

    /**
     * Apply array formulas to all cells in their ranges
     */
    private applyArrayFormulas = (cellData: IObjectMatrixPrimitiveType<ICellData>) => {
        for (const arrayFormula of this.arrayFormulas) {
            const { formula, range, masterRow, masterCol } = arrayFormula;
            
            // Apply the formula to all cells in the range
            for (let r = range.startRow; r <= range.endRow; r++) {
                for (let c = range.startCol; c <= range.endCol; c++) {
                    if (!cellData[r]) cellData[r] = {};
                    if (!cellData[r][c]) cellData[r][c] = {};
                    
                    // For array formulas like TRANSPOSE, only the master cell should have the formula
                    // Other cells should reference the master cell or be empty initially
                    if (r === masterRow && c === masterCol) {
                        cellData[r][c].f = formula;
                    } else {
                        // For non-master cells in an array formula, they may contain calculated values
                        // but should not have the formula duplicated
                        // We'll let Univer handle the array formula expansion
                        if (!cellData[r][c].v && !cellData[r][c].f) {
                            // Mark as part of an array formula result
                            cellData[r][c] = {
                                v: null, // Will be calculated by Univer
                                t: CellValueType.STRING
                            };
                        }
                    }
                }
            }
        }
    };

    /**
     * Parse Excel range string like "A1:C3" to row/column indices
     */
    private parseRange = (rangeStr: string): {startRow: number, endRow: number, startCol: number, endCol: number} | null => {
        try {
            const parts = rangeStr.split(':');
            if (parts.length !== 2) return null;
            
            const startCell = this.cellRefToIndices(parts[0]);
            const endCell = this.cellRefToIndices(parts[1]);
            
            if (!startCell || !endCell) return null;
            
            return {
                startRow: startCell.row,
                endRow: endCell.row,
                startCol: startCell.col,
                endCol: endCell.col
            };
        } catch {
            return null;
        }
    };

    /**
     * Convert cell reference like "A1" to row/column indices
     */
    private cellRefToIndices = (cellRef: string): {row: number, col: number} | null => {
        try {
            const match = cellRef.match(/^([A-Z]+)(\d+)$/);
            if (!match) return null;
            
            const colStr = match[1];
            const rowStr = match[2];
            
            // Convert column letters to number (A=0, B=1, etc.)
            let col = 0;
            for (let i = 0; i < colStr.length; i++) {
                col = col * 26 + (colStr.charCodeAt(i) - 64);
            }
            col -= 1; // Convert to 0-based
            
            const row = parseInt(rowStr) - 1; // Convert to 0-based
            
            return { row, col };
        } catch {
            return null;
        }
    };
}
