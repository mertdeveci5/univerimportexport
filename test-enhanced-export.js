#!/usr/bin/env node

/**
 * Test script for enhanced Univer to Excel export functionality
 * 
 * This script:
 * 1. Imports an Excel file to Univer format
 * 2. Exports it back to Excel using the enhanced export
 * 3. Compares features and validates round-trip integrity
 */

const fs = require('fs');
const path = require('path');
const LuckyExcel = require('./dist/luckyexcel.umd.js');

// Colors for console output
const colors = {
    reset: '\x1b[0m',
    bright: '\x1b[1m',
    red: '\x1b[31m',
    green: '\x1b[32m',
    yellow: '\x1b[33m',
    blue: '\x1b[34m',
    magenta: '\x1b[35m',
    cyan: '\x1b[36m'
};

function log(message, color = 'reset') {
    console.log(`${colors[color]}${message}${colors.reset}`);
}

function logSection(title) {
    console.log('\n' + '='.repeat(60));
    log(title, 'bright');
    console.log('='.repeat(60));
}

function logSuccess(message) {
    log('✅ ' + message, 'green');
}

function logError(message) {
    log('❌ ' + message, 'red');
}

function logInfo(message) {
    log('ℹ️  ' + message, 'cyan');
}

function logWarning(message) {
    log('⚠️  ' + message, 'yellow');
}

function analyzeUniverData(data) {
    const analysis = {
        workbook: {
            id: data.id,
            name: data.name,
            appVersion: data.appVersion,
            locale: data.locale,
            theme: data.theme
        },
        sheets: {},
        resources: {},
        styles: {
            count: 0,
            hasStyles: false
        }
    };

    // Analyze sheets
    if (data.sheetOrder && data.sheets) {
        data.sheetOrder.forEach(sheetId => {
            const sheet = data.sheets[sheetId];
            if (!sheet) return;

            const cellCount = sheet.cellData ? 
                Object.values(sheet.cellData).reduce((acc, row) => 
                    acc + (row ? Object.keys(row).length : 0), 0) : 0;

            analysis.sheets[sheet.name] = {
                id: sheetId,
                cellCount: cellCount,
                hasFormulas: checkForFormulas(sheet.cellData),
                hasMergedCells: !!(sheet.mergeData && sheet.mergeData.length > 0),
                hasArrayFormulas: !!(sheet.arrayFormulas && sheet.arrayFormulas.length > 0),
                hasConditionalFormatting: false,
                hasDataValidation: false,
                hasFilters: false,
                hasComments: false,
                hasHyperlinks: false,
                hasImages: false,
                rowCount: sheet.rowCount || 0,
                columnCount: sheet.columnCount || 0,
                hidden: sheet.hidden,
                tabColor: sheet.tabColor,
                freeze: sheet.freeze,
                defaultRowHeight: sheet.defaultRowHeight,
                defaultColumnWidth: sheet.defaultColumnWidth
            };
        });
    }

    // Analyze resources
    if (data.resources && Array.isArray(data.resources)) {
        data.resources.forEach(resource => {
            if (!analysis.resources[resource.name]) {
                analysis.resources[resource.name] = {
                    count: 0,
                    sheets: []
                };
            }
            
            analysis.resources[resource.name].count++;
            
            // Parse resource data to find which sheets it applies to
            try {
                const resourceData = JSON.parse(resource.data || '{}');
                
                // Update sheet-specific features based on resource type
                Object.keys(analysis.sheets).forEach(sheetName => {
                    const sheetId = Object.keys(data.sheets).find(id => 
                        data.sheets[id].name === sheetName);
                    
                    if (resourceData[sheetId]) {
                        analysis.resources[resource.name].sheets.push(sheetName);
                        
                        // Update sheet features based on resource type
                        switch(resource.name) {
                            case 'SHEET_CONDITIONAL_FORMATTING_PLUGIN':
                                analysis.sheets[sheetName].hasConditionalFormatting = true;
                                break;
                            case 'SHEET_DATA_VALIDATION_PLUGIN':
                                analysis.sheets[sheetName].hasDataValidation = true;
                                break;
                            case 'SHEET_FILTER_PLUGIN':
                                analysis.sheets[sheetName].hasFilters = true;
                                break;
                            case 'SHEET_COMMENT_PLUGIN':
                                analysis.sheets[sheetName].hasComments = true;
                                break;
                            case 'SHEET_HYPER_LINK_PLUGIN':
                                analysis.sheets[sheetName].hasHyperlinks = true;
                                break;
                            case 'SHEET_DRAWING_PLUGIN':
                                analysis.sheets[sheetName].hasImages = true;
                                break;
                        }
                    }
                });
            } catch (e) {
                // Resource data might not be JSON
            }
        });
    }

    // Analyze styles
    if (data.styles) {
        analysis.styles.count = Object.keys(data.styles).length;
        analysis.styles.hasStyles = analysis.styles.count > 0;
    }

    return analysis;
}

function checkForFormulas(cellData) {
    if (!cellData) return false;
    
    for (const row of Object.values(cellData)) {
        if (!row) continue;
        for (const cell of Object.values(row)) {
            if (cell && (cell.f || cell.si)) {
                return true;
            }
        }
    }
    return false;
}

function printAnalysis(analysis, title) {
    logSection(title);
    
    // Workbook info
    logInfo(`Workbook: ${analysis.workbook.name || 'Unnamed'}`);
    logInfo(`App Version: ${analysis.workbook.appVersion || 'Unknown'}`);
    logInfo(`Locale: ${analysis.workbook.locale || 'en-US'}`);
    
    // Sheets info
    console.log('\n📊 Sheets:');
    Object.entries(analysis.sheets).forEach(([name, sheet]) => {
        console.log(`  📋 ${name}:`);
        console.log(`     - Cells: ${sheet.cellCount}`);
        console.log(`     - Size: ${sheet.rowCount}x${sheet.columnCount}`);
        
        const features = [];
        if (sheet.hasFormulas) features.push('Formulas');
        if (sheet.hasMergedCells) features.push('Merged Cells');
        if (sheet.hasArrayFormulas) features.push('Array Formulas');
        if (sheet.hasConditionalFormatting) features.push('Conditional Formatting');
        if (sheet.hasDataValidation) features.push('Data Validation');
        if (sheet.hasFilters) features.push('Filters');
        if (sheet.hasComments) features.push('Comments');
        if (sheet.hasHyperlinks) features.push('Hyperlinks');
        if (sheet.hasImages) features.push('Images');
        if (sheet.freeze) features.push('Freeze Panes');
        if (sheet.hidden) features.push('Hidden');
        if (sheet.tabColor) features.push(`Tab Color: ${sheet.tabColor}`);
        
        if (features.length > 0) {
            console.log(`     - Features: ${features.join(', ')}`);
        }
    });
    
    // Resources info
    if (Object.keys(analysis.resources).length > 0) {
        console.log('\n🔧 Resources:');
        Object.entries(analysis.resources).forEach(([name, info]) => {
            const shortName = name.replace('SHEET_', '').replace('_PLUGIN', '');
            console.log(`  - ${shortName}: ${info.count} instance(s)`);
            if (info.sheets.length > 0) {
                console.log(`    Applied to: ${info.sheets.join(', ')}`);
            }
        });
    }
    
    // Styles info
    if (analysis.styles.hasStyles) {
        console.log('\n🎨 Styles:');
        console.log(`  - Style definitions: ${analysis.styles.count}`);
    }
}

async function testRoundTrip(inputFile, outputFile) {
    return new Promise((resolve, reject) => {
        logSection('Starting Enhanced Export Test');
        
        // Check if input file exists
        if (!fs.existsSync(inputFile)) {
            logError(`Input file not found: ${inputFile}`);
            reject(new Error('Input file not found'));
            return;
        }
        
        logInfo(`Input file: ${inputFile}`);
        logInfo(`Output file: ${outputFile}`);
        
        // Read the file
        const fileBuffer = fs.readFileSync(inputFile);
        
        // Create a File-like object for Node.js
        class NodeFile {
            constructor(buffer, name, options) {
                this.buffer = buffer;
                this.name = name;
                this.type = options.type;
                this.size = buffer.length;
            }
            
            arrayBuffer() {
                return Promise.resolve(this.buffer.buffer.slice(
                    this.buffer.byteOffset, 
                    this.buffer.byteOffset + this.buffer.byteLength
                ));
            }
            
            slice(start, end, type) {
                const sliced = this.buffer.slice(start, end);
                return new NodeFile(sliced, this.name, { type: type || this.type });
            }
        }
        
        const file = new NodeFile(fileBuffer, path.basename(inputFile), {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        
        logSection('Step 1: Import Excel to Univer');
        
        // Import Excel to Univer
        LuckyExcel.transformExcelToUniver(
            file,
            async (univerData) => {
                try {
                    logSuccess('Excel imported to Univer successfully');
                    
                    // Analyze imported data
                    const importAnalysis = analyzeUniverData(univerData);
                    printAnalysis(importAnalysis, 'Imported Data Analysis');
                    
                    // Save Univer data for inspection
                    const univerJsonFile = outputFile.replace('.xlsx', '_univer.json');
                    fs.writeFileSync(univerJsonFile, JSON.stringify(univerData, null, 2));
                    logInfo(`Univer data saved to: ${univerJsonFile}`);
                    
                    logSection('Step 2: Export Univer to Excel (Enhanced)');
                    
                    // Export back to Excel using enhanced export
                    await LuckyExcel.transformUniverToExcel({
                        snapshot: univerData,
                        fileName: path.basename(outputFile),
                        getBuffer: true,
                        success: (buffer) => {
                            try {
                                // Save the exported file
                                fs.writeFileSync(outputFile, buffer);
                                logSuccess(`Excel file exported successfully: ${outputFile}`);
                                
                                // File size comparison
                                const originalSize = fs.statSync(inputFile).size;
                                const exportedSize = fs.statSync(outputFile).size;
                                
                                console.log('\n📦 File Size Comparison:');
                                console.log(`  Original: ${(originalSize / 1024).toFixed(2)} KB`);
                                console.log(`  Exported: ${(exportedSize / 1024).toFixed(2)} KB`);
                                console.log(`  Difference: ${((exportedSize - originalSize) / 1024).toFixed(2)} KB`);
                                
                                // Now re-import the exported file to verify round-trip
                                logSection('Step 3: Verify Round-Trip (Re-import exported file)');
                                
                                const exportedBuffer = fs.readFileSync(outputFile);
                                const exportedFile = new NodeFile(exportedBuffer, path.basename(outputFile), {
                                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                });
                                
                                LuckyExcel.transformExcelToUniver(
                                    exportedFile,
                                    (reImportedData) => {
                                        logSuccess('Re-imported exported file successfully');
                                        
                                        // Analyze re-imported data
                                        const reImportAnalysis = analyzeUniverData(reImportedData);
                                        printAnalysis(reImportAnalysis, 'Re-imported Data Analysis');
                                        
                                        // Compare analyses
                                        logSection('Feature Comparison');
                                        compareAnalyses(importAnalysis, reImportAnalysis);
                                        
                                        // Save re-imported data for inspection
                                        const reImportJsonFile = outputFile.replace('.xlsx', '_reimport.json');
                                        fs.writeFileSync(reImportJsonFile, JSON.stringify(reImportedData, null, 2));
                                        logInfo(`Re-imported data saved to: ${reImportJsonFile}`);
                                        
                                        logSection('Test Complete');
                                        logSuccess('Enhanced export test completed successfully!');
                                        
                                        resolve({
                                            original: importAnalysis,
                                            reImported: reImportAnalysis
                                        });
                                    },
                                    (error) => {
                                        logError('Failed to re-import exported file: ' + error.message);
                                        reject(error);
                                    }
                                );
                            } catch (error) {
                                logError('Error processing exported file: ' + error.message);
                                reject(error);
                            }
                        },
                        error: (err) => {
                            logError('Export failed: ' + err.message);
                            console.error(err);
                            reject(err);
                        }
                    });
                } catch (error) {
                    logError('Error in processing: ' + error.message);
                    reject(error);
                }
            },
            (error) => {
                logError('Import failed: ' + error.message);
                reject(error);
            }
        );
    });
}

function compareAnalyses(original, reImported) {
    let hasIssues = false;
    
    // Compare sheet counts
    const originalSheetCount = Object.keys(original.sheets).length;
    const reImportedSheetCount = Object.keys(reImported.sheets).length;
    
    if (originalSheetCount === reImportedSheetCount) {
        logSuccess(`Sheet count preserved: ${originalSheetCount}`);
    } else {
        logError(`Sheet count mismatch: ${originalSheetCount} → ${reImportedSheetCount}`);
        hasIssues = true;
    }
    
    // Compare each sheet's features
    Object.entries(original.sheets).forEach(([sheetName, originalSheet]) => {
        const reImportedSheet = reImported.sheets[sheetName];
        
        if (!reImportedSheet) {
            logError(`Sheet missing after round-trip: ${sheetName}`);
            hasIssues = true;
            return;
        }
        
        console.log(`\n📋 Sheet: ${sheetName}`);
        
        // Compare features
        const features = [
            'cellCount', 'hasFormulas', 'hasMergedCells', 'hasArrayFormulas',
            'hasConditionalFormatting', 'hasDataValidation', 'hasFilters',
            'hasComments', 'hasHyperlinks', 'hasImages'
        ];
        
        features.forEach(feature => {
            if (originalSheet[feature] === reImportedSheet[feature]) {
                if (originalSheet[feature]) {
                    logSuccess(`  ${feature}: ✓ (preserved)`);
                }
            } else {
                logWarning(`  ${feature}: ${originalSheet[feature]} → ${reImportedSheet[feature]}`);
                if (originalSheet[feature] && !reImportedSheet[feature]) {
                    hasIssues = true;
                }
            }
        });
    });
    
    // Compare resources
    console.log('\n🔧 Resources Comparison:');
    const originalResources = Object.keys(original.resources);
    const reImportedResources = Object.keys(reImported.resources);
    
    originalResources.forEach(resource => {
        if (reImportedResources.includes(resource)) {
            logSuccess(`  ${resource}: ✓`);
        } else {
            logWarning(`  ${resource}: Missing after round-trip`);
        }
    });
    
    // Compare styles
    if (original.styles.hasStyles) {
        if (reImported.styles.hasStyles) {
            logSuccess(`\n🎨 Styles preserved: ${reImported.styles.count} style(s)`);
        } else {
            logWarning('\n🎨 Styles lost during round-trip');
        }
    }
    
    if (!hasIssues) {
        console.log('\n' + '='.repeat(60));
        logSuccess('🎉 All critical features preserved during round-trip!');
        console.log('='.repeat(60));
    } else {
        console.log('\n' + '='.repeat(60));
        logWarning('⚠️ Some features may have been lost during round-trip');
        console.log('='.repeat(60));
    }
}

// Main execution
async function main() {
    const testFile = '/Users/mertdeveci/Desktop/Code/alphafrontend/docs/excel/test.xlsx';
    const outputFile = '/Users/mertdeveci/Desktop/Code/univerimportexport/test-enhanced-output.xlsx';
    
    try {
        await testRoundTrip(testFile, outputFile);
        
        console.log('\n📁 Output files created:');
        console.log(`  - ${outputFile} (exported Excel file)`);
        console.log(`  - ${outputFile.replace('.xlsx', '_univer.json')} (Univer data)`);
        console.log(`  - ${outputFile.replace('.xlsx', '_reimport.json')} (re-imported data)`);
        
    } catch (error) {
        logError('Test failed: ' + error.message);
        console.error(error);
        process.exit(1);
    }
}

// Run the test
main();