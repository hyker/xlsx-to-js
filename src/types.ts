export interface XlsxParserOptions {
    /** Cooperative cancellation. Use XlsxWorkerParser to interrupt synchronous XML work. */
    signal?: AbortSignal;
    onProgress?: (progress: XlsxParseProgress) => void;
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

export interface XlsxParseProgress {
    phase: 'archive' | 'metadata' | 'worksheet' | 'complete';
    completedSheets: number;
    totalSheets: number;
    sheetName?: string;
}

/** Zero-based page indexes. Each page is limited to 5,000 grid positions. */
export interface XlsxPageOptions extends XlsxRenderOptions {
    rowPage?: number;
    columnPage?: number;
    pageRows?: number;
    pageColumns?: number;
}

export interface XlsxSheetPage {
    html: string;
    rowPage: number;
    columnPage: number;
    totalRowPages: number;
    totalColumnPages: number;
    /** Visible rows/columns, including expansion needed for drawings. */
    totalRows: number;
    totalColumns: number;
    /** Original one-based sheet coordinates; zero for an empty page. */
    rowStart: number;
    rowEnd: number;
    columnStart: number;
    columnEnd: number;
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
    /** Intrinsic raster dimensions, independent of drawing layout size. */
    maxImageDimension: number;
    maxImagePixels: number;
    /** Sum of raster pixels for every image drawing, including repeated media. */
    maxTotalImagePixels: number;
    maxXmlNodes: number;
    maxXmlDepth: number;
    maxXmlAttributes: number;
    maxXmlAttributesPerElement: number;
    /** Total UTF-16 code units in parsed workbook strings, counting each cloned occurrence. */
    maxWorkbookTextLength: number;
    /** Maximum serialized HTML length (UTF-16 code units). */
    maxHtmlLength: number;
}

export interface XlsxRenderOptions {
    limits?: Partial<XlsxLimits>;
    /** Hidden sheets remain in parsed data, but are omitted from HTML by default. */
    includeHiddenSheets?: boolean;
}
