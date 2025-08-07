import { LuckyFile } from "./ToLuckySheet/LuckyFile";
// import {SecurityDoor,Car} from './content';

import { HandleZip } from './HandleZip';
import { HandleXls } from './HandleXls';

import { IuploadfileList } from "./ICommon";

import { WorkBook } from "./UniverToExcel/Workbook";
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
                    console.error(err);
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
                    console.error(err);
                }
            });
    }


    static async transformExcelToUniver(
        excelFile: File,
        callback?: (files: IWorkbookData, fs?: string) => void,
        errorHandler?: (err: Error) => void
    ) {
        console.log('🚀 [PACKAGE] transformExcelToUniver START', {
            fileName: excelFile.name,
            fileSize: `${(excelFile.size / 1024).toFixed(2)} KB`,
            fileType: excelFile.type,
            timestamp: new Date().toISOString()
        });
        
        const startTime = Date.now();
        
        try {
            // Handle both XLS and XLSX files
            const processExcelFiles = async (files: IuploadfileList) => {
                console.log('📦 [PACKAGE] Processing Excel files...', {
                    fileCount: Object.keys(files).length,
                    elapsed: `${Date.now() - startTime}ms`
                });
                
                console.log('📦 [PACKAGE] Creating LuckyFile...');
                let luckyFile = new LuckyFile(files, excelFile.name);
                
                console.log('📦 [PACKAGE] Parsing LuckyFile...');
                let luckysheetfile = luckyFile.Parse();
                
                console.log('📦 [PACKAGE] Parsing JSON output...');
                let exportJson = JSON.parse(luckysheetfile);
                
                console.log('📦 [PACKAGE] Parsed data:', {
                    sheets: exportJson?.data?.length || 0,
                    elapsed: `${Date.now() - startTime}ms`
                });
                
                if (callback != undefined) {
                    console.log('📦 [PACKAGE] Creating UniverWorkBook...');
                    const univerData = new UniverWorkBook(exportJson);
                    
                    console.log('📦 [PACKAGE] Calling callback with data...');
                    callback(univerData.mode, luckysheetfile);
                    console.log('✅ [PACKAGE] transformExcelToUniver COMPLETE', {
                        totalTime: `${Date.now() - startTime}ms`
                    });
                }
            };

            // Check if it's an XLS file
            if (HandleXls.isXlsFile(excelFile)) {
                console.log('📁 [PACKAGE] XLS file detected, converting to XLSX...');
                const files = await HandleXls.convertXlsToXlsx(excelFile);
                console.log('📁 [PACKAGE] XLS conversion complete');
                await processExcelFiles(files);
            } else {
                // Handle XLSX file normally
                console.log('📁 [PACKAGE] XLSX file detected, unzipping...');
                
                // Wrap the callback-based unzipFile in a Promise
                await new Promise<void>((resolve, reject) => {
                    let handleZip: HandleZip = new HandleZip(excelFile);
                    handleZip.unzipFile(
                        async (files: IuploadfileList) => {
                            try {
                                await processExcelFiles(files);
                                resolve();
                            } catch (err) {
                                reject(err);
                            }
                        },
                        function (err: Error) {
                            console.error('❌ [PACKAGE] Unzip error:', {
                                error: err.message,
                                elapsed: `${Date.now() - startTime}ms`
                            });
                            if (errorHandler) {
                                errorHandler(err);
                            }
                            reject(err);
                        }
                    );
                });
            }
        } catch (err) {
            console.error('❌ [PACKAGE] Transform error:', {
                error: err instanceof Error ? err.message : String(err),
                stack: err instanceof Error ? err.stack : undefined,
                elapsed: `${Date.now() - startTime}ms`
            });
            if (errorHandler) {
                errorHandler(err as Error);
            } else {
                console.error(err);
            }
        }
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
            // console.log(1, new Date())
            const workbook = new WorkBook(snapshot);
            // console.log(snapshot, workbook)
            // console.log(2, new Date())
            const buffer = await workbook.xlsx.writeBuffer();
            // console.log(3, new Date())
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
            console.log(csv);

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