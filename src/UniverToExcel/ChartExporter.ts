import { debug } from '../utils/debug';
import { jsonParse } from '../common/method';

/**
 * ChartExporter - Handles chart export for Excel files
 * 
 * This class processes chart data from Univer format and converts it
 * to ExcelJS-compatible chart objects for Excel export.
 */
export class ChartExporter {
    
    /**
     * Export charts from resources to ExcelJS worksheets
     */
    static exportCharts(workbook: any, worksheet: any, sheetId: string, resources: any[]): void {
        if (!resources || !Array.isArray(resources)) {
            return;
        }

        // Find chart plugin data
        const chartResource = resources.find(r => r.name === 'SHEET_CHART_PLUGIN');
        if (!chartResource) {
            debug.log('📊 [ChartExporter] No charts found in resources');
            return;
        }

        const chartData = jsonParse(chartResource.data);
        if (!chartData || typeof chartData !== 'object') {
            debug.log('📊 [ChartExporter] Invalid chart data format');
            return;
        }

        // Get charts for this specific sheet
        const sheetCharts = chartData[sheetId];
        if (!sheetCharts || !Array.isArray(sheetCharts)) {
            debug.log('📊 [ChartExporter] No charts found for sheet:', sheetId);
            return;
        }

        debug.log('📊 [ChartExporter] Found charts for sheet:', {
            sheetId,
            chartCount: sheetCharts.length,
            chartTypes: sheetCharts.map(c => c.chartType)
        });

        // Process each chart
        sheetCharts.forEach((chart, index) => {
            try {
                this.exportSingleChart(worksheet, chart, index);
            } catch (error) {
                debug.log('❌ [ChartExporter] Error exporting chart:', {
                    chartIndex: index,
                    chartId: chart.id,
                    error: error.message
                });
            }
        });
    }

    /**
     * Export a single chart to ExcelJS
     */
    private static exportSingleChart(worksheet: any, chartData: any, index: number): void {
        const { 
            id, 
            chartType, 
            rangeInfo, 
            context, 
            style 
        } = chartData;

        debug.log('📊 [ChartExporter] Exporting chart:', {
            id,
            chartType,
            hasRangeInfo: !!rangeInfo,
            hasContext: !!context
        });

        // Map Univer chart types to ExcelJS chart types
        const excelChartType = this.mapChartType(chartType);
        if (!excelChartType) {
            debug.log('⚠️ [ChartExporter] Unsupported chart type:', chartType);
            return;
        }

        // Extract data range
        const dataRange = this.extractDataRange(rangeInfo);
        if (!dataRange) {
            debug.log('⚠️ [ChartExporter] Invalid data range for chart:', id);
            return;
        }

        // Create ExcelJS chart configuration
        const chartConfig = this.createChartConfig(excelChartType, dataRange, context, style);
        
        try {
            // Add chart to worksheet
            const chart = worksheet.addChart(chartConfig);
            
            debug.log('✅ [ChartExporter] Successfully exported chart:', {
                id,
                type: excelChartType,
                range: dataRange
            });

        } catch (error) {
            debug.log('❌ [ChartExporter] Failed to add chart to worksheet:', {
                id,
                error: error.message,
                chartConfig
            });
        }
    }

    /**
     * Map Univer chart type to ExcelJS chart type
     */
    private static mapChartType(univerType: string): string | null {
        const typeMap: { [key: string]: string } = {
            'column': 'col',
            'bar': 'bar', 
            'line': 'line',
            'pie': 'pie',
            'doughnut': 'doughnut',
            'scatter': 'scatter',
            'area': 'area',
            'combo': 'combo'
        };

        return typeMap[univerType] || null;
    }

    /**
     * Extract data range from Univer range info
     */
    private static extractDataRange(rangeInfo: any): string | null {
        if (!rangeInfo || !rangeInfo.rangeInfo || !rangeInfo.rangeInfo.range) {
            return null;
        }

        const range = rangeInfo.rangeInfo.range;
        
        // Convert Univer range format to Excel range format
        // Univer: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 }
        // Excel: \"A1:C6\"
        
        if (typeof range === 'object' && 
            typeof range.startRow === 'number' &&
            typeof range.endRow === 'number' &&
            typeof range.startColumn === 'number' &&
            typeof range.endColumn === 'number') {
            
            const startCol = this.numberToColumnLetter(range.startColumn);
            const endCol = this.numberToColumnLetter(range.endColumn);
            const startRow = range.startRow + 1; // Convert to 1-based
            const endRow = range.endRow + 1;
            
            return `${startCol}${startRow}:${endCol}${endRow}`;
        }

        // If it's already a string range, use it directly
        if (typeof range === 'string') {
            return range;
        }

        return null;
    }

    /**
     * Create ExcelJS chart configuration
     */
    private static createChartConfig(chartType: string, dataRange: string, context: any, style: any): any {
        const config: any = {
            type: chartType,
            data: {
                // Basic data configuration
                series: [
                    {
                        name: 'Series 1',
                        ref: dataRange
                    }
                ]
            },
            options: {
                // Chart positioning (will be auto-positioned if not specified)
                top: style?.top || 100,
                left: style?.left || 100,
                width: style?.width || 400,
                height: style?.height || 300
            }
        };

        // Add chart-specific configurations
        if (context) {
            if (context.title) {
                config.title = {
                    name: context.title,
                    size: 14
                };
            }

            if (context.xAxisTitle) {
                config.categoryAxis = {
                    title: {
                        name: context.xAxisTitle
                    }
                };
            }

            if (context.yAxisTitle) {
                config.valueAxis = {
                    title: {
                        name: context.yAxisTitle
                    }
                };
            }

            // Handle legend
            if (context.showLegend !== false) {
                config.legend = {
                    position: context.legendPosition || 'bottom'
                };
            }
        }

        return config;
    }

    /**
     * Convert column number to Excel column letter (0-based)
     */
    private static numberToColumnLetter(colNumber: number): string {
        let result = '';
        let num = colNumber;
        
        while (num >= 0) {
            result = String.fromCharCode(65 + (num % 26)) + result;
            num = Math.floor(num / 26) - 1;
            if (num < 0) break;
        }
        
        return result;
    }

    /**
     * Check if ExcelJS supports chart export
     */
    static isChartExportSupported(): boolean {
        // ExcelJS has basic chart support, but it may be limited
        // This method can be used to check capabilities
        return true;
    }

    /**
     * Get supported chart types
     */
    static getSupportedChartTypes(): string[] {
        return ['column', 'bar', 'line', 'pie', 'doughnut', 'scatter', 'area'];
    }

    /**
     * Debug method to analyze chart data
     */
    static analyzeChartData(resources: any[]): any {
        const chartResource = resources?.find(r => r.name === 'SHEET_CHART_PLUGIN');
        if (!chartResource) {
            return { hasCharts: false };
        }

        const chartData = jsonParse(chartResource.data);
        if (!chartData) {
            return { hasCharts: false, invalidData: true };
        }

        const analysis = {
            hasCharts: true,
            sheetCount: Object.keys(chartData).length,
            sheets: {} as any
        };

        for (const [sheetId, charts] of Object.entries(chartData)) {
            if (Array.isArray(charts)) {
                analysis.sheets[sheetId] = {
                    chartCount: charts.length,
                    chartTypes: charts.map((c: any) => c.chartType),
                    chartIds: charts.map((c: any) => c.id)
                };
            }
        }

        return analysis;
    }
}