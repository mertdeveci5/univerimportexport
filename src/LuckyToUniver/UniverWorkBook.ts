import {
    IStyleData,
    IWorkbookData,
    IWorksheetData,
    Nullable,
} from '@univerjs/core';
import { IResources } from '@univerjs/core/lib/types/services/resource-manager/type';
import { DrawingTypeEnum, PresetGeometryType, LocaleType } from '../common/univerEnums';
import { debug } from '../utils/debug';
import { HyperLink, UniverSheet } from './UniverSheet';
// import { ISheetDrawing, SheetDrawingAnchorType } from '@univerjs/sheets-drawing';
import { handleStyle } from './utils';
import { generateRandomId } from '../common/method';
import { IluckySheet, ILuckyFile, IWorkBookInfo } from '../ToLuckySheet/ILuck';
import { ImageSourceType } from './ILuckInterface';
import { handleRanges } from '../ToLuckySheet/style';

interface Sheets {
    [sheetId: string]: Partial<IWorksheetData & { hyperLink: HyperLink[] }>;
}
interface LuckySheetObj {
    [sheetId: string]: Partial<IluckySheet>;
}

export class UniverWorkBook implements IWorkbookData {
    id!: string;
    rev?: number | undefined;
    name!: string;
    appVersion!: string;
    locale!: any; // Changed from LocaleType to any to avoid import issues
    styles!: Record<string, Nullable<IStyleData>>;
    sheetOrder!: string[];
    sheets!: Sheets;
    resources?: IResources | undefined = [];
    constructor(file: ILuckyFile) {
        debug.log('🔍 [UniverWorkBook] Constructor called with file:', {
            hasInfo: !!file.info,
            hasSheets: !!file.sheets,
            hasData: !!(file as any).data,
            hasWorkbook: !!file.workbook,
            fileKeys: Object.keys(file)
        });
        
        // Handle both 'sheets' and 'data' properties for backwards compatibility
        const sheets = file.sheets || (file as any).data || [];
        const { info, workbook } = file;
        
        debug.log('🔍 [UniverWorkBook] Sheets extracted:', {
            sheetsCount: sheets.length,
            sheetNames: sheets.map((s: IluckySheet) => s.name)
        });
        
        this.id = generateRandomId(6);
        this.name = info.name;
        this.appVersion = info.appversion;
        this.locale = LocaleType.ZH_CN; // Using local enum

        const workSheets: Sheets = {},
            order: string[] = [],
            sheetsObj: LuckySheetObj = {};
        debug.log('🔍 [UniverWorkBook] Processing sheets before sort:', sheets.map((s: IluckySheet) => ({
            name: s.name, 
            order: s.order,
            hasCellData: !!(s.celldata && s.celldata.length > 0),
            cellCount: s.celldata ? s.celldata.length : 0
        })));
        
        sheets
            .sort((a: IluckySheet, b: IluckySheet) => Number(a.order) - Number(b.order))
            .forEach((d: IluckySheet) => {
                debug.log('🔍 [DEBUG] Creating UniverSheet for:', {
                    name: d.name,
                    order: d.order,
                    hasCellData: !!(d.celldata && d.celldata.length > 0),
                    cellCount: d.celldata ? d.celldata.length : 0
                });
                const sheet = new UniverSheet(d);
                workSheets[sheet.id] = sheet.mode;
                sheetsObj[sheet.id] = d;
                order.push(sheet.id);
                debug.log('✅ [DEBUG] Sheet created with ID:', sheet.id, 'name:', sheet.name);
            });

        // debug.log(workSheets,sheets)
        this.handleHyperLinks(workSheets);
        this.handleImage(workSheets, sheets);
        this.handleChart(workSheets, sheets);
        this.handleNames(workbook);
        this.handleCondition(sheetsObj);
        this.handleVerification(sheetsObj);
        this.handleFilter(sheetsObj);
        this.sheetOrder = order;

        // Collect all styles from sheets into registry
        this.collectStyles(workSheets);

        this.sheets = workSheets;
    }

    get mode(): IWorkbookData {
        return {
            id: this.id,
            rev: this.rev,
            name: this.name,
            appVersion: this.appVersion,
            locale: this.locale,
            styles: this.styles,
            sheetOrder: this.sheetOrder,
            sheets: this.sheets,
            resources: this.resources,
        };
    }
    private handleHyperLinks = (workSheets: Sheets) => {
        const hyperLinks: { [key: string]: HyperLink[] } = {};
        for (const key in workSheets) {
            const link = workSheets[key].hyperLink;
            if (!link?.length) continue;
            hyperLinks[key] = link.map((d: any) => {
                let payload = d.payload;
                if (typeof d.payload !== 'string') {
                    payload = '#';
                    const gid = d.payload.gid.replace(/'|"/g, '');
                    const sheetId = Object.values(workSheets).find(
                        (sheet) => sheet.name === gid
                    )?.id;

                    if (gid && sheetId) {
                        payload += `gid=${sheetId}`;
                    }
                    if (gid && sheetId && d.payload.range) payload += '&';
                    if (d.payload.range) payload += `range=${d.payload.range}`;
                }
                return {
                    ...d,
                    payload,
                };
            });
        }
        // debug.log(workSheets, hyperLinks)
        this.resources?.push({
            name: 'SHEET_HYPER_LINK_PLUGIN',
            data: JSON.stringify(hyperLinks),
        });
    };
    private handleImage = (workSheets: Sheets, sheets: IluckySheet[]) => {
        const drawerList: {
            [key: string]: { order: string[]; data: { [key: string]: any } };
        } = {};

        Object.values(workSheets).forEach((sheet) => {
            const images = sheets.find((d) => d.name === sheet.name)?.images;
            if (!images) return;
            const order = Object.keys(images);
            const data: { [key: string]: any } = {};
            order.forEach((key) => {
                const image = images[key];
                if (sheet.columnCount < image.toCol) {
                    sheet.columnCount = image.toCol + 1;
                }
                if (sheet.rowCount < image.toRow) {
                    sheet.rowCount = image.toRow + 1;
                }
                let imageObj: any = {
                    unitId: this.id,
                    subUnitId: sheet.id || '',
                    drawingId: key,
                    transform: {
                        width: 0,
                        height: 0,
                        scaleX: 0,
                        scaleY: 0,
                        left: 0,
                        top: 0,
                        angle: 0,
                        skewX: 0,
                        skewY: 0,
                        flipX: false,
                        flipY: false,
                        ...(image.transform || {}),
                    },
                    sheetTransform: {
                        angle: 0,
                        skewX: 0,
                        skewY: 0,
                        flipX: false,
                        flipY: false,
                        from: {
                            column: image.fromCol,
                            columnOffset: image.fromColOff,
                            row: image.fromRow,
                            rowOffset: image.fromRowOff,
                        },
                        to: {
                            column: image.toCol,
                            columnOffset: image.toColOff,
                            row: image.toRow,
                            rowOffset: image.toRowOff,
                        },
                    },
                }
                if (image.type === 'chart') {
                    imageObj = {
                        ...imageObj,
                        drawingType: DrawingTypeEnum.DRAWING_CHART,
                        componentKey: 'Chart',
                        data: {
                            ...(image.data || {}),
                            range: `${sheet.name}!${image.data.range}`
                        },
                        allowTransform: true
                    }
                } else {
                    imageObj = {
                        ...imageObj,
                        drawingType: DrawingTypeEnum.DRAWING_IMAGE,
                        imageSourceType: ImageSourceType.BASE64,
                        source: image.src,
                        prstGeom: PresetGeometryType.Rect as Nullable<PresetGeometryType>,
                        anchorType: '1',
                    }
                }

                data[key] = imageObj;
            });
            drawerList[sheet.id!] = {
                data,
                order,
            };
        });
        this.resources?.push({
            name: 'SHEET_DRAWING_PLUGIN',
            data: JSON.stringify(drawerList),
        });
    };

    private handleChart = (workSheets: Sheets, sheets: IluckySheet[]) => {
        const chartList: {
            [key: string]: any
        } = {};
        Object.values(workSheets).forEach((sheet) => {
            const charts = sheets.find((d) => d.name === sheet.name)?.charts;
            if (!charts) return;
            charts.forEach((chart) => {
                if (!chartList[sheet.id!]) {
                    chartList[sheet.id!] = []
                }
                chartList[sheet.id!].push({
                    rangeInfo: {
                        isRowDirection: chart.isRowDirection,
                        rangeInfo: {
                            unitId: this.id,
                            subUnitId: sheet.id || '',
                            range: handleRanges(chart.range)[0]
                        }
                    },
                    id: chart.id,
                    chartType: chart.chartType,
                    context: chart.context,
                    style: chart.style,
                    dataAggregation: {}
                })
            });
        })
        // debug.log('chartList', chartList)
        this.resources?.push({
            name: 'SHEET_CHART_PLUGIN',
            data: JSON.stringify(chartList),
        });
    }
    private handleNames = (workbook: IWorkBookInfo) => {
        this.resources?.push({
            name: 'SHEET_DEFINED_NAME_PLUGIN',
            data: JSON.stringify(workbook.defineNames),
        });
    };
    private handleCondition = (sheets: LuckySheetObj) => {
        const obj: any = {};
        Object.keys(sheets).forEach((d) => {
            const condition = sheets[d].conditionalFormatting?.map((d: any) => {
                if (d.rule?.style) {
                    d.rule.style = handleStyle(
                        { v: d.rule.style, r: 0, c: 0 },
                        { value: d.rule?.style?.border, rangeType: '' }
                    );
                }
                return d;
            });
            obj[d] = condition;
        });
        this.resources?.push({
            name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
            data: JSON.stringify(obj),
        });
    };

    private handleVerification = (sheets: LuckySheetObj) => {
        const obj: any = {};
        Object.keys(sheets).forEach((d) => {
            obj[d] = sheets[d].dataVerificationList;
        });
        this.resources?.push({
            name: 'SHEET_DATA_VALIDATION_PLUGIN',
            data: JSON.stringify(obj),
        });
    };
    private handleFilter = (sheets: LuckySheetObj) => {
        const obj: any = {};
        Object.keys(sheets).forEach((d) => {
            obj[d] = sheets[d].filter;
        });
        this.resources?.push({
            name: 'SHEET_FILTER_PLUGIN',
            data: JSON.stringify(obj),
        });
    };

    private collectStyles(workSheets: Sheets): void {
        const styleRegistry: Record<string, IStyleData> = {};
        let styleIdCounter = 0;
        
        debug.log('🎨 [UniverWorkBook] Starting style collection from sheets');
        
        // Iterate through all sheets
        for (const sheetId in workSheets) {
            const sheet = workSheets[sheetId];
            if (!sheet.cellData) continue;
            
            let sheetStyleCount = 0;
            
            // Iterate through all cells
            Object.values(sheet.cellData).forEach((rowData: any) => {
                Object.values(rowData).forEach((cell: any) => {
                    if (cell.s && typeof cell.s === 'object') {
                        // This cell has an inline style object
                        const styleKey = JSON.stringify(cell.s);
                        
                        // Check if we've seen this style before
                        let styleId = Object.keys(styleRegistry).find(
                            id => JSON.stringify(styleRegistry[id]) === styleKey
                        );
                        
                        if (!styleId) {
                            // New style, add to registry
                            styleId = `style_${styleIdCounter++}`;
                            styleRegistry[styleId] = cell.s;
                            sheetStyleCount++;
                        }
                        
                        // Replace inline style with style ID reference
                        cell.s = styleId;
                    }
                });
            });
            
            if (sheetStyleCount > 0) {
                debug.log(`  Sheet ${sheet.name || sheetId}: ${sheetStyleCount} unique styles`);
            }
        }
        
        this.styles = styleRegistry;
        
        const borderCount = Object.values(styleRegistry).filter((s: any) => s.bd).length;
        debug.log('📊 [UniverWorkBook] Style collection complete:', {
            totalStyles: Object.keys(styleRegistry).length,
            stylesWithBorders: borderCount
        });
    }
}
