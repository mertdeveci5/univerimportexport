/**
 * Constants for Excel Export functionality
 * 
 * This file contains all constants to avoid magic numbers and strings
 * throughout the export codebase.
 */

// Excel Column/Row Constants
export const EXCEL_COLUMN_CONSTANTS = {
    ALPHABET_SIZE: 26,
    ASCII_UPPERCASE_A: 65, // 'A' in ASCII
    MAX_COLUMNS: 16384, // Excel max columns (XFD)
    MAX_ROWS: 1048576  // Excel max rows
} as const;

// Excel Sheet Name Constants  
export const EXCEL_SHEET_CONSTANTS = {
    MAX_NAME_LENGTH: 31, // Excel's maximum sheet name length
    DEFAULT_SHEET_NAME: 'Sheet1',
    RESERVED_NAMES: [
        'History',
        'Print_Titles', 
        'Print_Area',
        'Sheet_Title',
        'Consolidate_Area',
        'Auto_Open',
        'Auto_Close',
        'Extract',
        'Database',
        'Criteria',
        'Print_Settings',
        'Recorder',
        'Data_Form',
        'Auto_Activate',
        'Auto_Deactivate',
        'Sheet_Title'
    ] as const,
    // Characters not allowed in Excel sheet names
    INVALID_CHAR_MAPPING: {
        '\\': '_',
        '/': '_',
        '?': '_',
        '*': '_',
        '[': '(',
        ']': ')',
        ':': '_'
    } as const
} as const;

// Formula Constants
export const FORMULA_CONSTANTS = {
    // Excel error codes
    ERROR_CODES: {
        REF: '#REF!',
        NAME: '#NAME?',
        VALUE: '#VALUE!',
        DIV0: '#DIV/0!',
        NULL: '#NULL!',
        NUM: '#NUM!',
        NA: '#N/A',
        GETTING_DATA: '#GETTING_DATA',
        SPILL: '#SPILL!'
    } as const,
    
    // Formula prefixes to clean
    PREFIXES_TO_REMOVE: [
        '_xlfn.',
        '_xlws.',
        '_xll.',
        '_xlpm.'
    ] as const,
    
    // Excel formula special characters
    SPECIAL_CHARS: {
        FORMULA_START: '=',
        CELL_ABSOLUTE: '$',
        RANGE_SEPARATOR: ':',
        ARGUMENT_SEPARATOR: ',',
        STRING_DELIMITER: '"',
        ARRAY_ROW_SEPARATOR: ';',
        ARRAY_COLUMN_SEPARATOR: ','
    } as const
} as const;

// Chart Export Constants
export const CHART_CONSTANTS = {
    PLUGIN_NAME: 'SHEET_CHART_PLUGIN',
    
    // Chart type mappings from Univer to ExcelJS
    TYPE_MAPPING: {
        'column': 'col',
        'bar': 'bar',
        'line': 'line',
        'pie': 'pie',
        'doughnut': 'doughnut',
        'scatter': 'scatter',
        'area': 'area',
        'combo': 'combo',
        'radar': 'radar',
        'stock': 'stock',
        'surface': 'surface',
        'bubble': 'bubble'
    } as const,
    
    // Default chart dimensions
    DEFAULT_DIMENSIONS: {
        TOP: 100,
        LEFT: 100,
        WIDTH: 400,
        HEIGHT: 300
    } as const,
    
    // Default chart properties
    DEFAULT_PROPERTIES: {
        SERIES_NAME: 'Series',
        TITLE_SIZE: 14,
        LEGEND_POSITION: 'bottom' as const
    } as const
} as const;

// Hyperlink Constants
export const HYPERLINK_CONSTANTS = {
    PLUGIN_NAME: 'SHEET_HYPER_LINK_PLUGIN',
    URL_FRAGMENTS: {
        GID_PREFIX: '#gid=',
        RANGE_PREFIX: 'range=',
        SHEET_SEPARATOR: '&',
        FRAGMENT_SEPARATOR: '#'
    } as const,
    LINK_TYPES: {
        EXTERNAL: 'External',
        INTERNAL: 'Internal',
        EMAIL: 'Email',
        FILE: 'File'
    } as const
} as const;

// Image Constants
export const IMAGE_CONSTANTS = {
    DEFAULT_EXTENSION: 'png',
    SUPPORTED_EXTENSIONS: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg'] as const,
    MAX_IMAGE_SIZE: 10 * 1024 * 1024 // 10MB in bytes
} as const;

// Number Format Constants
export const NUMBER_FORMAT_CONSTANTS = {
    GENERAL: 'General',
    TEXT: '@',
    NUMBER: '0',
    NUMBER_2_DECIMAL: '0.00',
    PERCENTAGE: '0%',
    PERCENTAGE_2_DECIMAL: '0.00%',
    DATE: 'yyyy/mm/dd',
    TIME: 'hh:mm:ss',
    CURRENCY: '$#,##0.00'
} as const;

// Boolean representation constants
export const BOOLEAN_VALUES = {
    TRUE: 1,
    FALSE: 0
} as const;

// Index conversion constants
export const INDEX_OFFSET = {
    ZERO_TO_ONE_BASED: 1  // Add this to convert 0-based to 1-based indexing
} as const;

// Resource Plugin Names
export const RESOURCE_PLUGINS = {
    CHART: 'SHEET_CHART_PLUGIN',
    HYPERLINK: 'SHEET_HYPER_LINK_PLUGIN',
    IMAGE: 'SHEET_IMAGE_PLUGIN',
    COMMENT: 'SHEET_COMMENT_PLUGIN',
    CONDITIONAL_FORMAT: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
    DATA_VALIDATION: 'SHEET_DATA_VALIDATION_PLUGIN',
    FILTER: 'SHEET_FILTER_PLUGIN',
    PIVOT_TABLE: 'SHEET_PIVOT_TABLE_PLUGIN'
} as const;

// Type exports for TypeScript type safety
export type ExcelErrorCode = typeof FORMULA_CONSTANTS.ERROR_CODES[keyof typeof FORMULA_CONSTANTS.ERROR_CODES];
export type ChartType = keyof typeof CHART_CONSTANTS.TYPE_MAPPING;
export type LinkType = typeof HYPERLINK_CONSTANTS.LINK_TYPES[keyof typeof HYPERLINK_CONSTANTS.LINK_TYPES];
export type ResourcePluginName = typeof RESOURCE_PLUGINS[keyof typeof RESOURCE_PLUGINS];