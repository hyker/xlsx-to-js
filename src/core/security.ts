import { XlsxLimits } from '../types';
import { Workbook } from './workbook/types';

/** Sized for a ~300 MB peak worker heap so previews survive modern phones.
 * Peak memory is dominated by the DOM of the largest XML part, roughly
 * 0.6-0.7 KB per node. A full 100,000-position grid of value or shared-string
 * cells needs about 310,000 nodes and 210,000 attributes, so `maxCells` is the
 * budget users normally reach; formula-heavy sheets reach `maxXmlNodes` near
 * 80,000 cells. Raise cells, nodes and attributes together.
 */
export const DEFAULT_LIMITS: Readonly<XlsxLimits> = Object.freeze({
    maxFileBytes: 10 * 1024 * 1024,
    maxEntries: 1024,
    maxEntryBytes: 16 * 1024 * 1024,
    maxTotalBytes: 64 * 1024 * 1024,
    maxSheets: 32,
    // Excel's row limit: tall sheets are bounded by maxCells instead.
    maxRows: 1048576,
    // Column layout work grows with this limit, so it stays below Excel's 16,384.
    maxColumns: 1024,
    maxCells: 100_000,
    maxMergedCells: 100_000,
    maxDrawings: 1000,
    maxDrawingPixels: 100000,
    maxImageDimension: 8192,
    maxImagePixels: 4 * 1024 * 1024,
    maxTotalImagePixels: 16 * 1024 * 1024,
    maxXmlNodes: 400_000,
    maxXmlDepth: 64,
    maxXmlAttributes: 400_000,
    maxXmlAttributesPerElement: 256,
    maxWorkbookTextLength: 16 * 1024 * 1024,
    maxHtmlLength: 16 * 1024 * 1024,
});

/** Parsing budgets match DEFAULT_LIMITS; each rendered page gets a smaller HTML budget. */
export const PREVIEW_LIMITS: Readonly<XlsxLimits> = Object.freeze({
    ...DEFAULT_LIMITS,
    maxHtmlLength: 2 * 1024 * 1024,
});

/** A workbook or render exceeded one of the configured `XlsxLimits`.
 * `limit` names the budget and `max` is the value that was in force, so
 * applications can explain the rejection or decide whether to raise it.
 */
export class XlsxLimitError extends Error {
    readonly name = 'XlsxLimitError';
    constructor(readonly limit: keyof XlsxLimits, readonly max: number, readonly detail: string) {
        super(`${detail} (limit ${limit} = ${max})`);
    }
}

export function limitError(limits: XlsxLimits, limit: keyof XlsxLimits, detail: string): XlsxLimitError {
    return new XlsxLimitError(limit, limits[limit], detail);
}

export function resolveLimits(overrides: Partial<XlsxLimits> = {}): XlsxLimits {
    const limits = { ...DEFAULT_LIMITS };
    for (const key of Object.keys(overrides) as (keyof XlsxLimits)[]) {
        if (!Object.prototype.hasOwnProperty.call(DEFAULT_LIMITS, key)) throw new Error('Unknown resource limit');
        const value = overrides[key];
        if (!Number.isSafeInteger(value) || value! <= 0) throw new Error(`Invalid resource limit: ${key}`);
        limits[key] = value!;
    }
    if (limits.maxRows > 1048576 || limits.maxColumns > 16384) throw new Error('Limits exceed Excel coordinates');
    return limits;
}

/** Match structured clone's object sharing, but count strings at every occurrence.
 * This also covers formulas, style text, metadata and repeated drawing media.
 */
export function assertWorkbookTextBudget(workbook: Workbook, limits: XlsxLimits): void {
    let remaining = limits.maxWorkbookTextLength;
    const seen = new WeakSet<object>();
    const pending: unknown[] = [workbook];
    while (pending.length) {
        const value = pending.pop();
        if (typeof value === 'string') {
            if (value.length > remaining) throw limitError(limits, 'maxWorkbookTextLength', 'Workbook text exceeds resource limits');
            remaining -= value.length;
        } else if (value !== null && typeof value === 'object' && !seen.has(value)) {
            seen.add(value);
            for (const child of Object.values(value)) pending.push(child);
        }
    }
}

export function boundedNumber(value: unknown, label: string, max: number, min = 0, integer = false): number {
    if (value === null || value === undefined || value === '' || (typeof value !== 'string' && typeof value !== 'number')) {
        throw new Error(`Invalid ${label}`);
    }
    // The length cap and unambiguous grammar keep validation linear. `\d+\.?\d*`
    // backtracks quadratically on long digit runs from untrusted attributes.
    if (typeof value === 'string' && (value.length > 64 ||
        !(integer ? /^\+?\d+$/ : /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/).test(value.trim()))) {
        throw new Error(`Invalid ${label}`);
    }
    const n = Number(value);
    if (!Number.isFinite(n) || n < min || n > max || (integer && !Number.isSafeInteger(n))) {
        throw new Error(`Invalid ${label}`);
    }
    return n;
}

export function parseReference(ref: string, limits: XlsxLimits = resolveLimits()): { row: number; col: number } {
    const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref);
    if (!match) throw new Error('Invalid cell reference');
    let col = 0;
    for (const letter of match[1]) col = col * 26 + letter.charCodeAt(0) - 64;
    const row = Number(match[2]);
    if (row > limits.maxRows) throw limitError(limits, 'maxRows', `Cell reference exceeds resource limits: ${ref} is beyond the row limit`);
    if (col > limits.maxColumns) throw limitError(limits, 'maxColumns', `Cell reference exceeds resource limits: ${ref} is beyond the column limit`);
    return { row, col };
}

export function parseRange(ref: string, limits: XlsxLimits = resolveLimits()) {
    const parts = ref.split(':');
    if (parts.length > 2) throw new Error('Invalid cell range');
    const start = parseReference(parts[0], limits);
    const end = parseReference(parts[1] ?? parts[0], limits);
    if (end.row < start.row || end.col < start.col) throw new Error('Reversed cell range');
    if (end.row * end.col > limits.maxCells) throw limitError(limits, 'maxCells', `Grid exceeds cell budget: A1:${parts[1] ?? parts[0]} spans ${end.row * end.col} grid positions`);
    return { start, end, area: (end.row - start.row + 1) * (end.col - start.col + 1) };
}

export function escapeHtml(value: unknown): string {
    return String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]!));
}

export function safeColor(value: unknown, fallback = ''): string {
    return typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value) ? value : fallback;
}

/** A quoted CSS string, subsequently escaped for its HTML attribute context. */
export function fontFamily(value: string): string {
    if (value.length > 256 || /[\u0000-\u001f\u007f]/.test(value)) return 'sans-serif';
    return `"${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

export function enumValue(value: unknown, allowed: readonly string[], fallback = ''): string {
    return typeof value === 'string' && allowed.includes(value) ? value : fallback;
}
