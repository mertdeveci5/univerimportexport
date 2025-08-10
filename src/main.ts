import { LuckyFile } from "./ToLuckySheet/LuckyFile";
// import {SecurityDoor,Car} from './content';

import { HandleZip } from './HandleZip';
import { HandleXls } from './HandleXls';
import { debug } from './utils/debug';

import { IuploadfileList } from "./ICommon";

import { WorkBook } from "./UniverToExcel/Workbook";
import { exportUniverToExcel } from "./UniverToExcel/index";
import exceljs from "@zwight/exceljs";
import { CSV } from "./UniverToCsv/CSV";
import { isObject } from "./common/method";
import { UniverWorkBook } from "./LuckyToUniver/UniverWorkBook";
import { IWorkbookData } from "@univerjs/core";
import { formatSheetData, getDataByFile } from "./common/utils";
import { UniverCsvWorkBook } from "./LuckyToUniver/UniverCsvWorkBook";
export class LuckyExcel {
    constructor() { }
    static transformExcelToLucky(excelFile: File,
        callback?: (files: IuploadfileList, fs?: string) => void,
        errorHandler?: (err: Error) => void) {
        let handleZip: HandleZip = new HandleZip(excelFile);

        // const fileReader = new FileReader();
        // fileReader.onload = async (e) => {
        //     const { result } = e.target as any;
        //     const workbook = new exceljs.Workbook();
        //     const data = await workbook.xlsx.load(result);
        //     // console.log('exceljs', data)
        // }
        // fileReader.readAsArrayBuffer(excelFile)

        handleZip.unzipFile(function (files: IuploadfileList) {
            let luckyFile = new LuckyFile(files, excelFile.name);
            let luckysheetfile = luckyFile.Parse();
            let exportJson = JSON.parse(luckysheetfile);
            // console.log('output---->', exportJson)
            if (callback != undefined) {
                callback(exportJson, luckysheetfile);
            }
        },
            function (err: Error) {
                if (errorHandler) {
                    errorHandler(err);
                } else {
                    debug.error(err);
                }
            });
    }

    static transformExcelToLuckyByUrl(
        url: string,
        name: string,
        callBack?: (files: IuploadfileList, fs?: string) => void,
        errorHandler?: (err: Error) => void) {
        let handleZip: HandleZip = new HandleZip();
        handleZip.unzipFileByUrl(url, function (files: IuploadfileList) {
            let luckyFile = new LuckyFile(files, name);
            let luckysheetfile = luckyFile.Parse();
            let exportJson = JSON.parse(luckysheetfile);
            if (callBack != undefined) {
                callBack(exportJson, luckysheetfile);
            }
        },
            function (err: Error) {
                if (errorHandler) {
                    errorHandler(err);
                } else {
                    debug.error(err);
                }
            });
    }


    static async transformExcelToUniver(
        excelFile: File,
        callback?: (files: IWorkbookData, fs?: string) => void,
        errorHandler?: (err: Error) => void
    ) {
        // Debug logging removed
        
        debug.log('🚀 [PACKAGE] transformExcelToUniver START', {
            fileName: excelFile.name,
            fileSize: `${(excelFile.size / 1024).toFixed(2)} KB`,
            fileType: excelFile.type,
            timestamp: new Date().toISOString()
        });
        
        const startTime = Date.now();
        
        // Return a Promise that resolves when the callback is called
        return new Promise<void>((resolveMain, rejectMain) => {
            try {
                // Handle both XLS and XLSX files
                const processExcelFiles = async (files: IuploadfileList) => {
                    try {
                        debug.log('📦 [PACKAGE] Processing Excel files...', {
                            fileCount: Object.keys(files).length,
                            elapsed: `${Date.now() - startTime}ms`
                        });
                        
                        debug.log('📦 [PACKAGE] Creating LuckyFile...');
                        let luckyFile = new LuckyFile(files, excelFile.name);
                        
                        debug.log('📦 [PACKAGE] Parsing LuckyFile...');
                        let luckysheetfile = luckyFile.Parse();
                        
                        debug.log('📦 [PACKAGE] Parsing JSON output...');
                        let exportJson = JSON.parse(luckysheetfile);
                        
                        debug.log('📦 [PACKAGE] Parsed exportJson structure:', {
                            hasData: !!exportJson?.data,
                            hasSheets: !!exportJson?.sheets,
                            dataLength: exportJson?.data?.length || 0,
                            sheetsLength: exportJson?.sheets?.length || 0,
                            topLevelKeys: Object.keys(exportJson || {}),
                            elapsed: `${Date.now() - startTime}ms`
                        });
                        
                        if (exportJson?.data) {
                            debug.log('📦 [PACKAGE] Sheets in data property:', exportJson.data.map((s: any) => s.name));
                        }
                        if (exportJson?.sheets) {
                            debug.log('📦 [PACKAGE] Sheets in sheets property:', exportJson.sheets.map((s: any) => s.name));
                        }
                        
                        if (callback != undefined) {
                            debug.log('📦 [PACKAGE] Creating UniverWorkBook...');
                            const univerData = new UniverWorkBook(exportJson);
                            
                            // Debug logging removed
                            
                            debug.log('📦 [PACKAGE] Calling callback with data...');
                            callback(univerData.mode, luckysheetfile);
                            debug.log('✅ [PACKAGE] transformExcelToUniver COMPLETE', {
                                totalTime: `${Date.now() - startTime}ms`
                            });
                        }
                        resolveMain();
                    } catch (err) {
                        debug.error('❌ [PACKAGE] Process error:', err);
                        if (errorHandler) {
                            errorHandler(err as Error);
                        }
                        rejectMain(err);
                    }
                };

                // Check if it's an XLS file
                if (HandleXls.isXlsFile(excelFile)) {
                    debug.log('📁 [PACKAGE] XLS file detected, converting to XLSX...');
                    HandleXls.convertXlsToXlsx(excelFile)
                        .then(files => {
                            debug.log('📁 [PACKAGE] XLS conversion complete');
                            return processExcelFiles(files);
                        })
                        .catch(err => {
                            debug.error('❌ [PACKAGE] XLS conversion error:', err);
                            if (errorHandler) {
                                errorHandler(err);
                            }
                            rejectMain(err);
                        });
                } else {
                    // Handle XLSX file normally
                    debug.log('📁 [PACKAGE] XLSX file detected, unzipping...');
                    let handleZip: HandleZip = new HandleZip(excelFile);
                    handleZip.unzipFile(
                        (files: IuploadfileList) => {
                            processExcelFiles(files).catch(err => {
                                debug.error('❌ [PACKAGE] Processing error:', err);
                                if (errorHandler) {
                                    errorHandler(err);
                                }
                                rejectMain(err);
                            });
                        },
                        function (err: Error) {
                            debug.error('❌ [PACKAGE] Unzip error:', {
                                error: err.message,
                                elapsed: `${Date.now() - startTime}ms`
                            });
                            if (errorHandler) {
                                errorHandler(err);
                            }
                            rejectMain(err);
                        }
                    );
                }
            } catch (err) {
                debug.error('❌ [PACKAGE] Transform error:', {
                    error: err instanceof Error ? err.message : String(err),
                    stack: err instanceof Error ? err.stack : undefined,
                    elapsed: `${Date.now() - startTime}ms`
                });
                if (errorHandler) {
                    errorHandler(err as Error);
                }
                rejectMain(err);
            }
        });
    }

    static transformCsvToUniver(
        file: File,
        callback?: (files: IWorkbookData, fs?: string[][]) => void,
        errorHandler?: (err: Error) => void
    ) {
        try {
            getDataByFile({ file }).then((source) => {
                const sheetData = formatSheetData(source, file)!;
                const univerData = new UniverCsvWorkBook(sheetData || [])
                callback?.(univerData.mode, sheetData);
            })
        } catch (error) {
            errorHandler(error);
        }
    }

    static async transformUniverToExcel(params: {
        snapshot: any,
        fileName?: string,
        getBuffer?: boolean,
        success?: (buffer?: exceljs.Buffer) => void,
        error?: (err: Error) => void
    }) {
        const { snapshot, fileName = `excel_${(new Date).getTime()}.xlsx`, getBuffer = false, success, error } = params;
        try {
            debug.log('🚀 [transformUniverToExcel] Starting export with enhanced handler');
            
            // Use enhanced export for better feature support
            const buffer = await exportUniverToExcel(snapshot);
            
            debug.log('✅ [transformUniverToExcel] Export completed, buffer size:', buffer.length);
            if (getBuffer) {
                success?.(buffer);
            } else {
                this.downloadFile(fileName, buffer);
                success?.()
            }

        } catch (err) {
            error?.(err)
        }
    }

    static async transformUniverToCsv(params: {
        snapshot: any,
        fileName?: string,
        getBuffer?: boolean,
        sheetName?: string,
        success?: (csvContent?: string | { [key: string]: string }) => void,
        error?: (err: Error) => void
    }) {
        const { snapshot, fileName = `csv_${(new Date).getTime()}.csv`, getBuffer = false, success, error, sheetName } = params;
        try {
            const csv = new CSV(snapshot);
            debug.log(csv);

            let contents: string | { [key: string]: string };
            if (sheetName) {
                contents = csv.csvContent[sheetName];
            } else {
                contents = csv.csvContent;
            }
            if (getBuffer) {
                success?.(contents);
            } else {
                if (isObject(contents)) {
                    for (const key in contents) {
                        if (Object.prototype.hasOwnProperty.call(contents, key)) {
                            const element = contents[key];
                            this.downloadFile(`${fileName}_${key}`, element);
                        }
                    }
                } else {
                    this.downloadFile(fileName, contents);
                }
                success?.()
            }
        } catch (err) {
            error(err)
        }
    }

    private static downloadFile(fileName: string, buffer: exceljs.Buffer | string) {
        const link = document.createElement('a');
        let blob: Blob;
        if (typeof buffer === 'string') {
            blob = new Blob([buffer], { type: "text/csv;charset=utf-8;" });
        } else {
            blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8' });
        }

        const url = URL.createObjectURL(blob);
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();

        link.addEventListener('click', () => {
            link.remove();
            setTimeout(() => {
                URL.revokeObjectURL(url)
            }, 200);
        })
    }
}