import { XlsxLimits } from '../types';

export const DEFAULT_LIMITS: Readonly<XlsxLimits> = Object.freeze({
    maxFileBytes: 10 * 1024 * 1024,
    maxEntries: 1024,
    maxEntryBytes: 8 * 1024 * 1024,
    maxTotalBytes: 32 * 1024 * 1024,
    maxSheets: 32,
    maxRows: 10000,
    maxColumns: 1024,
    maxCells: 250000,
    maxMergedCells: 250000,
    maxDrawings: 1000,
    maxDrawingPixels: 100000,
    maxXmlNodes: 100000,
    maxXmlDepth: 64,
    maxHtmlLength: 16 * 1024 * 1024,
});

/** Budgets for paginated previews. Full-grid and archive limits still apply. */
export const PREVIEW_LIMITS: Readonly<XlsxLimits> = Object.freeze({
    ...DEFAULT_LIMITS,
    maxEntryBytes: 16 * 1024 * 1024,
    maxTotalBytes: 64 * 1024 * 1024,
    maxXmlNodes: 1_000_000,
    maxHtmlLength: 2 * 1024 * 1024,
});

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

export function boundedNumber(value: unknown, label: string, max: number, min = 0, integer = false): number {
    if (value === null || value === undefined || value === '' || (typeof value !== 'string' && typeof value !== 'number')) {
        throw new Error(`Invalid ${label}`);
    }
    if (typeof value === 'string' && !(integer ? /^\+?\d+$/ : /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/).test(value.trim())) {
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
    if (row > limits.maxRows || col > limits.maxColumns) throw new Error('Cell reference exceeds resource limits');
    return { row, col };
}

export function parseRange(ref: string, limits: XlsxLimits = resolveLimits()) {
    const parts = ref.split(':');
    if (parts.length > 2) throw new Error('Invalid cell range');
    const start = parseReference(parts[0], limits);
    const end = parseReference(parts[1] ?? parts[0], limits);
    if (end.row < start.row || end.col < start.col) throw new Error('Reversed cell range');
    if (end.row * end.col > limits.maxCells) throw new Error('Grid exceeds cell budget');
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
