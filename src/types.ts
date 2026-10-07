export interface XlsxParserOptions {
    /** Resource budgets. All supplied values must be positive safe integers. */
    limits?: Partial<XlsxLimits>;
    /** When the option `dense: false` is passed, parsers will skip empty cells. */
    dense?: boolean;
    /** When the option `styles: false` is passed, parsers will skip cell styles. */
    styles?: boolean;
    /** When the option `drawings: false` is passed, parsers will skip drawings. */
    drawings?: boolean;
    /** When `true`, hidden rows will be skipped during parsing. */
    skipHiddenRows?: boolean;
}

export interface XlsxLimits {
    maxFileBytes: number;
    maxEntries: number;
    maxEntryBytes: number;
    maxTotalBytes: number;
    maxSheets: number;
    maxRows: number;
    maxColumns: number;
    /** Total grid positions across the workbook, including empty cells. */
    maxCells: number;
    /** Total merged positions, counting repeated/overlapping ranges. */
    maxMergedCells: number;
    maxDrawings: number;
    /** Maximum coordinate, offset or extent in pixels. */
    maxDrawingPixels: number;
    maxXmlNodes: number;
    maxXmlDepth: number;
    /** Maximum serialized HTML length (UTF-16 code units). */
    maxHtmlLength: number;
}

export interface XlsxRenderOptions {
    limits?: Partial<XlsxLimits>;
    /** Hidden sheets remain in parsed data, but are omitted from HTML by default. */
    includeHiddenSheets?: boolean;
}
