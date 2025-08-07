import { IuploadfileList } from "./ICommon";
import * as XLSX from "xlsx";

export class HandleXls {
    /**
     * Convert .xls file to .xlsx format and return as zip structure
     */
    static async convertXlsToXlsx(file: File): Promise<IuploadfileList> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = async (e) => {
                try {
                    const data = e.target?.result;
                    const workbook = XLSX.read(data, { type: 'binary' });
                    
                    // Convert to XLSX format
                    const xlsxBuffer = XLSX.write(workbook, { 
                        bookType: 'xlsx', 
                        type: 'array',
                        bookSST: true,
                        compression: true
                    });
                    
                    // Create a blob and convert to File
                    const xlsxBlob = new Blob([xlsxBuffer], { 
                        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
                    });
                    const xlsxFile = new File([xlsxBlob], file.name.replace('.xls', '.xlsx'), { 
                        type: xlsxBlob.type 
                    });
                    
                    // Now process as regular XLSX
                    const { HandleZip } = await import('./HandleZip');
                    const handleZip = new HandleZip(xlsxFile);
                    
                    handleZip.unzipFile(
                        (files: IuploadfileList) => resolve(files),
                        (err: Error) => reject(err)
                    );
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => reject(new Error('Failed to read XLS file'));
            reader.readAsBinaryString(file);
        });
    }
    
    /**
     * Check if a file is .xls format
     */
    static isXlsFile(file: File): boolean {
        return file.name.toLowerCase().endsWith('.xls') && 
               !file.name.toLowerCase().endsWith('.xlsx');
    }
}