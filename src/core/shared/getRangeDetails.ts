import { DEFAULT_LIMITS, parseRange } from '../security';

/** Gets the origin and spans of a validated Excel range. */
export function getRangeDetails(range: string): { origin: string; colspan: number; rowspan: number } {
    // This helper only computes metadata, so it can cover Excel's full axis.
    const { start, end } = parseRange(range, { ...DEFAULT_LIMITS, maxRows: 1048576, maxColumns: 16384, maxCells: 1048576 * 16384 });
    return { origin: range.split(':')[0], colspan: end.col - start.col + 1, rowspan: end.row - start.row + 1 };
}
