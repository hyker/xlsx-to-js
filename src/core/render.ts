import { Workbook } from './workbook/types';
import { WorkSheet } from './worksheet/types';
import { Drawing } from './drawing/types';
import { XlsxLimits, XlsxRenderOptions } from '../types';
import { boundedNumber, enumValue, escapeHtml, fontFamily, parseRange, resolveLimits, safeColor } from './security';

const CSS = `<style>
.xlwb{font-family:Arial,sans-serif;color:#222}.xl-sheet{margin:12px 0}.xl-name{font-weight:600;margin:6px 0}
.xl{border-collapse:collapse;table-layout:fixed;border:1px solid #d0d7de}
.xl th,.xl td{border:1px solid #d0d7de;padding:0;white-space:pre;box-sizing:border-box}
.xl th{text-align:center;font-weight:600;font-size:12px;color:#57606a;background:#f6f8fa}
.xl td{background:#fff;font-size:13px;overflow:visible}.xl thead th{position:sticky;top:0;z-index:1}
.xl th.xl-row{position:sticky;left:0;z-index:1}.xl-wrap{position:relative;display:inline-block}
.xl-abs{position:absolute;left:0;top:0;z-index:2;pointer-events:none}.xl-abs img{position:absolute;object-fit:contain}
</style>`;

class Html {
    private chunks: string[] = [];
    private length = 0;
    constructor(private max: number) {}
    add(value: string): void {
        if ((this.length += value.length) > this.max) throw new Error('HTML exceeds output budget');
        this.chunks.push(value);
    }
    text(value: unknown): string {
        const str = String(value);
        if (str.length > this.max) throw new Error('Text exceeds output budget');
        return escapeHtml(str);
    }
    finish(): string { return this.chunks.join(''); }
}

const borders = new Map(Object.entries({
    hair: '1px solid', thin: '1px solid', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed',
    dashDotDot: '1px dashed', medium: '2px solid', mediumDashed: '2px dashed', mediumDashDot: '2px dashed',
    mediumDashDotDot: '2px dashed', thick: '3px solid', double: '3px double', slantDashDot: '2px dashed',
}));

function cellCss(style: WorkSheet['data'][number][number]['style']): string {
    if (!style) return '';
    let css = `white-space:${style.wrapText ? 'pre-wrap' : 'pre'};`;
    const horizontal = enumValue(style.hAlign, ['left', 'center', 'right', 'justify']);
    const vertical = enumValue(style.vAlign, ['top', 'center', 'bottom'], 'bottom');
    if (horizontal) css += `text-align:${horizontal};`;
    css += `vertical-align:${vertical === 'center' ? 'middle' : vertical};`;
    if (style.fontName) css += `font-family:${fontFamily(style.fontName)};`;
    if (style.fontSize) css += `font-size:${boundedNumber(style.fontSize, 'font size', 409) * 96 / 72}px;`;
    if (style.bold) css += 'font-weight:bold;';
    if (style.italic) css += 'font-style:italic;';
    const color = safeColor(style.fontColor), fill = safeColor(style.fgColor) || safeColor(style.bgColor);
    if (color) css += `color:${color};`;
    if (fill) css += `background-color:${fill};`;
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
        const border = style.border?.[side];
        const rule = border && borders.get(border.style);
        if (rule) css += `border-${side}:${rule} ${safeColor(border.color, '#000000')};`;
    }
    return css;
}

function columnName(index: number): string {
    let result = '';
    while (index > 0) { result = String.fromCharCode(65 + (index - 1) % 26) + result; index = Math.floor((index - 1) / 26); }
    return result;
}

function imageMime(base64: string, limits: XlsxLimits): string {
    if (base64.length > Math.ceil(limits.maxEntryBytes / 3) * 4 || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(base64)) {
        throw new Error('Invalid image data');
    }
    const head = atob(base64.slice(0, 16));
    if (head.startsWith('\x89PNG\r\n\x1a\n')) return 'image/png';
    if (head.startsWith('\xff\xd8\xff')) return 'image/jpeg';
    if (head.startsWith('GIF87a') || head.startsWith('GIF89a')) return 'image/gif';
    throw new Error('Unsupported image format');
}

function validateDrawing(drawing: Drawing, limits: XlsxLimits): void {
    for (const position of [drawing.position?.from, drawing.position?.to]) {
        if (!position) throw new Error('Invalid drawing position');
        boundedNumber(position.col, 'drawing column', limits.maxColumns - 1, 0, true);
        boundedNumber(position.row, 'drawing row', limits.maxRows - 1, 0, true);
        boundedNumber(position.colOff, 'drawing offset', limits.maxDrawingPixels * 9525, 0, true);
        boundedNumber(position.rowOff, 'drawing offset', limits.maxDrawingPixels * 9525, 0, true);
    }
    if (drawing.position.to.col < drawing.position.from.col || drawing.position.to.row < drawing.position.from.row) throw new Error('Reversed drawing anchor');
    for (const geometry of [drawing.sizeEMU, drawing.absEMU]) {
        if (geometry) for (const value of Object.values(geometry)) boundedNumber(value, 'drawing geometry', limits.maxDrawingPixels * 9525, 0, true);
    }
}

interface Budget { cells: number; merges: number; drawings: number }

export function renderWorkbook(workbook: Workbook, index: number | undefined, options: XlsxRenderOptions): string {
    const limits = resolveLimits(options.limits);
    if (workbook.workSheets.length > limits.maxSheets) throw new Error('Sheet count exceeds resource limits');
    const html = new Html(limits.maxHtmlLength);
    const budget: Budget = { cells: 0, merges: 0, drawings: 0 };
    html.add(`<div class="xlwb">${CSS}`);
    const sheets = index === undefined ? workbook.workSheets : [workbook.workSheets[index]];
    for (const sheet of sheets) {
        if (!options.includeHiddenSheets && sheet.state && sheet.state !== 'visible') continue;
        renderSheet(sheet, html, limits, budget);
    }
    html.add('</div>');
    return html.finish();
}

function renderSheet(sheet: WorkSheet, html: Html, limits: XlsxLimits, budget: Budget): void {
    const range = sheet.dimention ? parseRange(sheet.dimention, limits) : undefined;
    let rows = range?.end.row ?? 0, cols = range?.end.col ?? 0;
    const drawings = sheet.drawings.filter(d => ['image', 'shape', 'textbox'].includes(d.type));
    if ((budget.drawings += sheet.drawings.length) > limits.maxDrawings) throw new Error('Drawing count exceeds resource limits');
    for (const drawing of drawings) {
        validateDrawing(drawing, limits);
        rows = Math.max(rows, drawing.position.from.row + 1, drawing.position.to.row + 1);
        cols = Math.max(cols, drawing.position.from.col + 1, drawing.position.to.col + 1);
    }
    if (rows * cols + budget.cells > limits.maxCells) throw new Error('Rendered grid exceeds cell budget');
    html.add(`<div class="xl-sheet"><div class="xl-name">${html.text(sheet.name)}</div>`);
    if (!rows || !cols) { html.add('<div>(empty sheet)</div></div>'); return; }
    if (sheet.rowStyles.length > limits.maxRows || sheet.columnStyles.length > limits.maxColumns) throw new Error('Layout styles exceed resource limits');
    const rowStyles = new Map(sheet.rowStyles.map(style => [boundedNumber(style.r, 'row index', limits.maxRows, 1, true), style]));
    let mdw = 7;
    if (typeof document !== 'undefined') {
        try {
            const context = document.createElement('canvas').getContext('2d');
            if (context) {
                context.font = `${boundedNumber(sheet.defaultFontSize ?? 11, 'font size', 409) * 96 / 72}px ${fontFamily(sheet.defaultFontName ?? 'Calibri')}`;
                const width = context.measureText('0').width;
                if (Number.isFinite(width) && width > 0 && width < 1024) mdw = width;
            }
        } catch { /* Canvas/font metrics are optional; use the deterministic fallback. */ }
    }
    const colPx = (units: number): number => {
        const width = boundedNumber(units, 'column width', 255);
        return width < 1 ? Math.floor(width * (mdw + 5)) : Math.floor(width * mdw + 5);
    };
    const defaultColPx = colPx(sheet.defaultColWidth);
    const defaultRowPx = Math.round(boundedNumber(sheet.defaultRowHeight, 'row height', 409) * 96 / 72);
    // Prefix sums make anchor layout and grid expansion linear, rather than quadratic.
    const widths = [36], heights = [28], x = [0], y = [0];
    const addColumn = () => {
        const col = widths.length;
        let width = defaultColPx;
        for (const style of sheet.columnStyles) {
            boundedNumber(style.min, 'column range', 16384, 1, true);
            boundedNumber(style.max, 'column range', 16384, style.min, true);
            if (col >= style.min && col <= style.max) width = style.hidden || style.collapsed ? 0 : colPx(style.width);
        }
        widths.push(width); x.push(x[col - 1] + (width ? width + 1 : 0));
    };
    const addRow = () => {
        const row = heights.length, style = rowStyles.get(row);
        const hidden = style ? style.hidden || style.collapsed : sheet.zeroHeight;
        const height = hidden ? 0 : style ? Math.round(boundedNumber(style.height, 'row height', 409) * 96 / 72) : defaultRowPx;
        heights.push(height); y.push(y[row - 1] + (height ? height + 1 : 0));
    };
    while (widths.length <= cols) addColumn();
    while (heights.length <= rows) addRow();
    const rect = (d: Drawing) => {
        const from = d.position.from, to = d.position.to;
        const absolute = d.anchorType === 'absolute' && d.absEMU;
        const left = absolute ? absolute.x / 9525 : x[from.col] + from.colOff / 9525;
        const top = absolute ? absolute.y / 9525 : y[from.row] + from.rowOff / 9525;
        const width = d.sizeEMU ? d.sizeEMU.cx / 9525 : x[to.col] + to.colOff / 9525 - left;
        const height = d.sizeEMU ? d.sizeEMU.cy / 9525 : y[to.row] + to.rowOff / 9525 - top;
        for (const value of [left, top, width, height, left + width, top + height]) boundedNumber(value, 'drawing rectangle', limits.maxDrawingPixels);
        return { left: left + 36, top: top + 28, width: Math.max(1, width), height: Math.max(1, height) };
    };
    for (const d of drawings) {
        const box = rect(d);
        const right = box.left - 36 + box.width, bottom = box.top - 28 + box.height;
        while (x[cols] < right) {
            if (cols >= limits.maxColumns || (cols + 1) * rows + budget.cells > limits.maxCells) throw new Error('Drawing expansion exceeds grid budget');
            cols++; addColumn();
        }
        while (y[rows] < bottom) {
            if (rows >= limits.maxRows || (rows + 1) * cols + budget.cells > limits.maxCells) throw new Error('Drawing expansion exceeds grid budget');
            rows++; addRow();
        }
    }
    budget.cells += rows * cols;
    const anchors = new Map<number, { rows: number; cols: number }>(), covered = new Set<number>();
    for (const ref of sheet.mergeCells) {
        const merge = parseRange(ref, limits);
        if ((budget.merges += merge.area) > limits.maxMergedCells) throw new Error('Merged cells exceed resource limits');
        if (merge.end.row > rows || merge.end.col > cols) throw new Error('Merge outside rendered dimension');
        const key = (merge.start.row - 1) * cols + merge.start.col - 1;
        anchors.set(key, { rows: merge.end.row - merge.start.row + 1, cols: merge.end.col - merge.start.col + 1 });
        for (let r = merge.start.row; r <= merge.end.row; r++) for (let c = merge.start.col; c <= merge.end.col; c++) {
            const current = (r - 1) * cols + c - 1;
            if (covered.has(current)) throw new Error('Overlapping merged cells');
            covered.add(current);
        }
    }
    html.add('<div class="xl-wrap"><table class="xl"><colgroup><col style="width:36px">');
    for (let c = 1; c <= cols; c++) html.add(`<col style="width:${widths[c]}px;${widths[c] ? '' : 'display:none;'}">`);
    html.add('</colgroup><thead><tr><th class="xl-corner"></th>');
    for (let c = 1; c <= cols; c++) html.add(`<th class="xl-col" data-col="${c}"${widths[c] ? '' : ' style="display:none"'}>${columnName(c)}</th>`);
    html.add('</tr></thead><tbody>');
    for (let r = 1; r <= rows; r++) {
        if (sheet.skipHiddenRows && heights[r] === 0) continue;
        html.add(`<tr style="height:${heights[r]}px;${heights[r] ? '' : 'display:none;'}"><th class="xl-row" data-row="${r}">${r}</th>`);
        for (let c = 1; c <= cols; c++) {
            const key = (r - 1) * cols + c - 1, merge = anchors.get(key);
            if (covered.has(key) && !merge) continue;
            const cell = sheet.data[r - 1]?.[c - 1];
            let attrs = '';
            if (merge) {
                const span = sheet.skipHiddenRows ? heights.slice(r, r + merge.rows).filter(h => h > 0).length : merge.rows;
                attrs = ` rowspan="${span}" colspan="${merge.cols}"`;
            }
            const css = `${cellCss(cell?.style)}${widths[c] ? '' : 'display:none;'}`;
            html.add(`<td class="xl-cell" data-ref="${columnName(c)}${r}"${attrs}${css ? ` style="${escapeHtml(css)}"` : ''}>${html.text(cell?.value ?? '')}</td>`);
        }
        html.add('</tr>');
    }
    html.add('</tbody></table>');
    if (drawings.length) html.add('<div class="xl-abs">');
    for (const drawing of drawings) {
        const { left, top, width, height } = rect(drawing);
        const position = `position:absolute;left:${left}px;top:${top}px;width:${width}px;height:${height}px;`;
        if (drawing.type === 'image') {
            const mime = imageMime(drawing.base64, limits);
            html.add(`<img alt="${html.text(drawing.description ?? '')}" src="data:${mime};base64,${drawing.base64}" style="${position}">`);
        } else {
            const p = drawing.properties ?? {};
            const fill = safeColor(p.fillColor, 'transparent'), line = safeColor(p.lineColor, 'transparent');
            const lineWidth = p.lineWidth ? boundedNumber(p.lineWidth, 'drawing line width', 12700000) / 12700 : 0;
            let css = `${position}background-color:${fill};border:${lineWidth}px solid ${line};`;
            let text = '';
            if (drawing.type === 'shape') css += `border-radius:${p.shapeType === 'roundRect' ? 8 : 0}px;`;
            else {
                const size = boundedNumber(p.fontPt ?? 11, 'drawing font size', 409) * 96 / 72;
                css += `white-space:pre-wrap;padding:2px;color:${safeColor(p.textColor, '#222222')};font-size:${size}px;font-weight:${p.bold ? 'bold' : 'normal'};font-style:${p.italic ? 'italic' : 'normal'};text-align:${enumValue(p.textAlign, ['left', 'center', 'right', 'justify'], 'left')};`;
                text = html.text(p.text ?? '');
            }
            html.add(`<div class="xl-${drawing.type}" style="${escapeHtml(css)}">${text}</div>`);
        }
    }
    if (drawings.length) html.add('</div>');
    html.add('</div></div>');
}
