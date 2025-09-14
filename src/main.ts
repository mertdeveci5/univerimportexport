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
        const startTime = Date.now();
        
        // Return a Promise that resolves when the callback is called
        return new Promise<void>((resolveMain, rejectMain) => {
            try {
                // Handle both XLS and XLSX files
                const processExcelFiles = async (files: IuploadfileList) => {
                    try {
                        let luckyFile = new LuckyFile(files, excelFile.name);
                        let luckysheetfile = luckyFile.Parse();
                        let exportJson = JSON.parse(luckysheetfile);
                        
                        if (callback != undefined) {
                            console.log('[PACKAGE] About to create UniverWorkBook...');
                            const univerData = new UniverWorkBook(exportJson);
                            console.log('[PACKAGE] UniverWorkBook created successfully');
                            callback(univerData.mode, luckysheetfile);
                            console.log('[PACKAGE] Callback invoked successfully');
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
                    HandleXls.convertXlsToXlsx(excelFile)
                        .then(files => {
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
            // Use enhanced export for better feature support
            const buffer = await exportUniverToExcel(snapshot);
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