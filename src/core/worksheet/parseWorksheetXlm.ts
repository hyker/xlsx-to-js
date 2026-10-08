import { StyleSheet } from '../style/types';
import { getChildByName, getChildrenByName, getElementByName, getElementsByName } from '../utils';
import { WorkSheet } from './types';
import { XlsxLimits } from '../../types';
import { boundedNumber, limitError, parseRange, parseReference, resolveLimits } from '../security';
import { parseXml, relationshipId } from '../xml';
import { compileNumberFormat, NumberFormatter } from '../numberFormat';

export const parseWorksheetXml = (
    str: string, styleSheet: StyleSheet, sharedStrings: string[], dense: boolean, skipHiddenRows: boolean,
    limits: XlsxLimits = resolveLimits(), date1904 = false, stylesEnabled = true,
    remaining = { cells: limits.maxCells, merges: limits.maxMergedCells }, locale?: string,
): WorkSheet => {
    const worksheet: WorkSheet = {
        id: 0, name: '', dimention: '', data: [], columnStyles: [], rowStyles: [], mergeCells: [], drawings: [],
        defaultColWidth: 8.43, baseColWidth: 8, defaultRowHeight: 15, zeroHeight: false, skipHiddenRows,
        defaultFontName: styleSheet.defaultFontName, defaultFontSize: styleSheet.defaultFontSize,
    };
    const doc = parseXml(str, 'worksheet', limits);
    const root = doc.documentElement;
    const drawing = getChildByName(root, 'drawing');
    if (drawing) worksheet.drawingRelationshipId = relationshipId(drawing);
    const data = getChildByName(root, 'sheetData');
    if (!data) throw new Error('Worksheet is missing sheetData');
    const dimension = getChildByName(root, 'dimension');
    const declared = dimension ? parseRange(dimension.getAttribute('ref') ?? '', limits) : undefined;
    const checkGrid = (row: number, col: number) => {
        if (row * col > remaining.cells) throw limitError(limits, 'maxCells', `Workbook grid exceeds resource limits: sheet grid reaches ${columnName(col)}${row}`);
    };
    if (declared) checkGrid(declared.end.row, declared.end.col);
    worksheet.dimention = dimension?.getAttribute('ref') ?? '';
    const format = getChildByName(root, 'sheetFormatPr');
    const isTrue = (v: string | null): boolean => v === '1' || v === 'true';
    if (format) {
        for (const key of ['baseColWidth', 'defaultColWidth', 'defaultRowHeight'] as const) {
            const value = format.getAttribute(key);
            if (value !== null) worksheet[key] = boundedNumber(value, key, key === 'defaultRowHeight' ? 409 : 255);
        }
        worksheet.zeroHeight = isTrue(format.getAttribute('zeroHeight'));
    }
    const columns = getChildrenByName(getChildByName(root, 'cols'), 'col');
    if (columns.length > limits.maxColumns) throw limitError(limits, 'maxColumns', 'Column styles exceed resource limits');
    for (const column of columns) {
        // Excel may style the entire column axis even when the used grid is small.
        const min = boundedNumber(column.getAttribute('min'), 'column range', 16384, 1, true);
        const max = boundedNumber(column.getAttribute('max'), 'column range', 16384, min, true);
        worksheet.columnStyles.push({ min, max, width: boundedNumber(column.getAttribute('width') ?? 8.43, 'column width', 255),
            hidden: isTrue(column.getAttribute('hidden')), collapsed: isTrue(column.getAttribute('collapsed')) });
    }
    let maxRow = declared?.end.row ?? 0, maxCol = declared?.end.col ?? 0, mergedArea = 0;
    const merged = new Set<number>();
    for (const merge of getChildrenByName(getChildByName(root, 'mergeCells'), 'mergeCell')) {
        const ref = merge.getAttribute('ref') ?? '';
        const range = parseRange(ref, limits);
        if ((mergedArea += range.area) > limits.maxMergedCells) throw limitError(limits, 'maxMergedCells', 'Merged cells exceed resource limits');
        if (mergedArea > remaining.merges) throw limitError(limits, 'maxMergedCells', 'Workbook merged cells exceed resource limits');
        for (let r = range.start.row; r <= range.end.row; r++) for (let c = range.start.col; c <= range.end.col; c++) {
            const key = (r - 1) * limits.maxColumns + c - 1;
            if (merged.has(key)) throw new Error('Overlapping merged cells');
            merged.add(key);
        }
        maxRow = Math.max(maxRow, range.end.row); maxCol = Math.max(maxCol, range.end.col);
        checkGrid(maxRow, maxCol);
        worksheet.mergeCells.push(ref);
    }
    const rows = getChildrenByName(data, 'row');
    if (rows.length > limits.maxRows) throw limitError(limits, 'maxRows', 'Rows exceed resource limits');
    const rowIds = new Set<number>(), refs = new Set<string>();
    const resolvedStyles = new Map<string, NonNullable<WorkSheet['data'][number][number]['style']>>();
    const formatters = new Map<number, NumberFormatter | undefined>();
    // Formatted text is allocated per cell, unlike shared strings, so charge it as it grows.
    let formattedLength = 0;
    let previousRow = 0;
    for (const row of rows) {
        const rowIndex = boundedNumber(row.getAttribute('r') ?? previousRow + 1, 'row index', limits.maxRows, 1, true);
        previousRow = rowIndex;
        if (rowIds.has(rowIndex)) throw new Error('Duplicate worksheet row');
        rowIds.add(rowIndex);
        const hidden = row.hasAttribute('hidden') ? isTrue(row.getAttribute('hidden')) : worksheet.zeroHeight && !isTrue(row.getAttribute('customHeight'));
        const collapsed = isTrue(row.getAttribute('collapsed'));
        worksheet.rowStyles.push({ r: rowIndex, height: boundedNumber(row.getAttribute('ht') ?? worksheet.defaultRowHeight, 'row height', 409), hidden, collapsed });
        let previousCol = 0;
        for (const cell of getChildrenByName(row, 'c')) {
            const ref = cell.getAttribute('r') ?? `${columnName(previousCol + 1)}${rowIndex}`;
            const pos = parseReference(ref, limits);
            previousCol = pos.col;
            if (pos.row !== rowIndex) throw new Error('Cell does not belong to its row');
            if (declared && (pos.row > declared.end.row || pos.col > declared.end.col)) throw new Error('Cell outside worksheet dimension');
            if (refs.has(ref)) throw new Error('Duplicate cell reference');
            refs.add(ref);
            if (refs.size > limits.maxCells) throw limitError(limits, 'maxCells', 'Cells exceed resource limits');
            maxRow = Math.max(maxRow, pos.row); maxCol = Math.max(maxCol, pos.col);
            checkGrid(maxRow, maxCol);
            const s = cell.getAttribute('s');
            const styleIndex = s === null ? undefined : boundedNumber(s, 'cell style index', styleSheet.cells.length - 1, 0, true);
            const cellStyle = styleIndex === undefined ? styleSheet.cells[0] : styleSheet.cells[styleIndex];
            const font = cellStyle && styleSheet.fonts[cellStyle.fontId];
            const fill = cellStyle && styleSheet.fills[cellStyle.fillId];
            const type = cell.getAttribute('t') ?? 'n';
            if (!['n', 's', 'str', 'inlineStr', 'b', 'e', 'd'].includes(type)) throw new Error('Invalid cell type');
            const formula = getElementByName(cell, 'f')?.textContent ?? '';
            let value = type === 'inlineStr'
                ? getElementsByName(getElementByName(cell, 'is'), 't').filter(t => (t.parentNode as Element | null)?.localName !== 'rPh').map(t => t.textContent ?? '').join('')
                : getElementByName(cell, 'v')?.textContent ?? '';
            if (type === 's') value = sharedStrings[boundedNumber(value, 'shared string index', sharedStrings.length - 1, 0, true)];
            else if (type === 'b' && !['0', '1', 'true', 'false'].includes(value)) throw new Error('Invalid boolean cell value');
            let raw: string | undefined;
            if (type === 'n' && value !== '') {
                const number = Number(value);
                if (!Number.isFinite(number)) throw new Error('Invalid numeric cell value');
                const id = cellStyle?.numFmtId ?? 0;
                if (id !== 0 && !(skipHiddenRows && hidden)) {
                    if (!formatters.has(id)) formatters.set(id, compileNumberFormat(id, styleSheet.numFmts?.get(id), date1904, locale));
                    const text = formatters.get(id)?.(number);
                    if (text !== undefined && text !== value) {
                        if ((formattedLength += text.length) > limits.maxWorkbookTextLength) throw limitError(limits, 'maxWorkbookTextLength', 'Formatted cell text exceeds resource limits');
                        raw = value; value = text;
                    }
                }
            }
            if (skipHiddenRows && hidden) continue;
            // Like Excel, numbers and dates align right; text, booleans and errors align left.
            const hAlign = cellStyle?.alignment?.horizontal || (type === 'n' && (raw ?? value) !== '' ? 'right' : 'left');
            const styleKey = `${styleIndex ?? 0}:${hAlign}`;
            let resolvedStyle = resolvedStyles.get(styleKey);
            if (stylesEnabled && cellStyle && font && fill && !resolvedStyle) {
                resolvedStyle = {
                    bgColor: fill.bgColor, fgColor: fill.patternType === 'solid' ? fill.fgColor : '',
                    fontName: font.name, fontSize: font.size, fontColor: font.color, bold: font.bold, italic: font.italic,
                    vAlign: cellStyle.alignment?.vertical ?? 'bottom', hAlign,
                    wrapText: cellStyle.alignment?.wrapText ?? false, border: styleSheet.borders[cellStyle.borderId],
                };
                resolvedStyles.set(styleKey, resolvedStyle);
            }
            // Allocate only after this cell and the growing grid pass validation.
            if (!worksheet.data[pos.row - 1]) worksheet.data[pos.row - 1] = [];
            worksheet.data[pos.row - 1][pos.col - 1] = {
                ref, value, formula,
                style: stylesEnabled ? resolvedStyle : undefined,
            };
            if (raw !== undefined) worksheet.data[pos.row - 1][pos.col - 1].raw = raw;
        }
    }
    if (maxRow && maxCol) {
        worksheet.dimention = `A1:${columnName(maxCol)}${maxRow}`;
        parseRange(worksheet.dimention, limits);
        worksheet.data.length = maxRow;
        for (let r = 0; r < maxRow; r++) {
            const row = worksheet.data[r] ?? (worksheet.data[r] = []);
            if (dense) for (let c = 0; c < maxCol; c++) {
                row[c] ??= { ref: '', value: '', formula: '' };
            }
        }
    }
    return worksheet;
};

function columnName(index: number): string {
    let result = '';
    while (index > 0) { result = String.fromCharCode(65 + (index - 1) % 26) + result; index = Math.floor((index - 1) / 26); }
    return result;
}
