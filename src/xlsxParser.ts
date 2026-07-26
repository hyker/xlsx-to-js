import JSZip from "jszip";
import { XlsxParserOptions } from "./types";
import { parseThemeXml, Theme } from "./core/theme";
import { parseStylesXml, StyleSheet } from "./core/style";
import { parseWorkbookXml, Workbook } from "./core/workbook";
import { parseSharedStringsXml } from "./core/sharedString";
import { parseWorksheetXml } from "./core/worksheet";
import { DrawingFile, MediaFile } from "./core/drawing/types";
import { getColumnIndex } from "./core/utils";
import { parseDrawingXml } from "./core/drawing";

export class XlsxParser {
    
    /**
     * Convert a workbook produced by readFile() into HTML markup that emulates
     * an Excel-like grid with column (A, B, C) and row (1, 2, 3) headers.
     * Returns a self-contained HTML snippet (with minimal inline styles).
     */
    toHTML(workbook: Workbook): string {
        const colToName = (n: number): string => {
            let s = '';
            while (n > 0) {
                const m = (n - 1) % 26;
                s = String.fromCharCode(65 + m) + s;
                n = Math.floor((n - 1) / 26);
            }
            return s;
        };

        const getEndFromDimension = (dimension: string): { rows: number, cols: number } => {
            if (!dimension) return { rows: 0, cols: 0 };
            const endRef = dimension.includes(':') ? dimension.split(':')[1] : dimension;
            const colLetters = (endRef.match(/[A-Z]+/) || ['A'])[0];
            const rowNum = parseInt((endRef.match(/\d+/) || ['0'])[0], 10) || 0;
            return { rows: rowNum, cols: getColumnIndex(colLetters) };
        };

        const escapeHtml = (value: string): string => {
            return value
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
        };

        const styles = `
<style>
/* Container */
.xlwb { font-family: Arial, sans-serif; color: #222; }
.xl-sheet { margin: 12px 0; }
.xl-name { font-weight: 600; margin: 6px 0; }

/* Grid */
.xl { border-collapse: collapse; table-layout: fixed; border: 1px solid #d0d7de; }
.xl thead th { position: sticky; top: 0; background: #f6f8fa; z-index: 1; }
.xl th.xl-row { position: sticky; left: 0; background: #f6f8fa; z-index: 1; }
.xl th, .xl td { border: 1px solid #d0d7de; padding: 0; white-space: pre; box-sizing: border-box; }
.xl td, .xl th { overflow: visible; }
.xl th { text-align: center; font-weight: 600; font-size: 12px; color: #57606a; }
.xl td { background: #fff; font-size: 13px; }
.xl .xl-corner { background: #f6f8fa; }
.xl .xl-col { }
.xl th.xl-row { }
</style>`;

        let html = `<div class="xlwb">${styles}`;

        for (const sheet of workbook.workSheets) {
            const { rows, cols } = getEndFromDimension(sheet.dimention);
            if (rows === 0 || cols === 0) {
                html += `<div class="xl-sheet"><div class="xl-name">${escapeHtml(sheet.name)}</div><div>(hoja vacía)</div></div>`;
                continue;
            }

            // Build merge maps
            const mergeTopLeft = new Map<string, { rowspan: number; colspan: number }>();
            const coveredCells = new Set<string>();
            for (const range of sheet.mergeCells) {
                const [start, end] = range.split(':');
                const sCol = (start.match(/[A-Z]+/) || ['A'])[0];
                const sRow = parseInt((start.match(/\d+/) || ['1'])[0], 10);
                const eCol = (end.match(/[A-Z]+/) || ['A'])[0];
                const eRow = parseInt((end.match(/\d+/) || ['1'])[0], 10);

                const cStart = getColumnIndex(sCol);
                const cEnd = getColumnIndex(eCol);
                const colspan = Math.max(1, cEnd - cStart + 1);
                const rowspan = Math.max(1, eRow - sRow + 1);

                mergeTopLeft.set(`${sCol}${sRow}`, { rowspan, colspan });
                for (let r = sRow; r <= eRow; r++) {
                    for (let c = cStart; c <= cEnd; c++) {
                        const ref = `${colToName(c)}${r}`;
                        if (ref !== `${sCol}${sRow}`) coveredCells.add(ref);
                    }
                }
            }

            // Render sheet
            html += `<div class="xl-sheet">`;
            html += `<div class="xl-name">${escapeHtml(sheet.name)}</div>`;
            html += `<table class="xl">`;

            // Header row with column letters
            html += `<thead><tr>`;
            html += `<th class="xl-corner"></th>`;
            for (let c = 1; c <= cols; c++) {
                html += `<th class="xl-col">${colToName(c)}</th>`;
            }
            html += `</tr></thead>`;

            // Body with row numbers and cells
            html += `<tbody>`;
            const cssBorder = (style: string, color: string): string => {
                if (!style) return '';
                const map: Record<string, string> = {
                    hair: '1px solid',
                    thin: '1px solid',
                    dotted: '1px dotted',
                    dashed: '1px dashed',
                    dashDot: '1px dashed',
                    dashDotDot: '1px dashed',
                    medium: '2px solid',
                    mediumDashed: '2px dashed',
                    mediumDashDot: '2px dashed',
                    mediumDashDotDot: '2px dashed',
                    thick: '3px solid',
                    double: '3px double',
                    slantDashDot: '2px dashed',
                };
                const widthStyle = map[style] || '1px solid';
                return `${widthStyle} ${color || '#000'}`;
            };

            for (let r = 1; r <= rows; r++) {
                html += `<tr>`;
                html += `<th class="xl-row">${r}</th>`;
                for (let c = 1; c <= cols; c++) {
                    const ref = `${colToName(c)}${r}`;
                    if (coveredCells.has(ref)) continue; // skip cells covered by a merge

                    const cell = sheet.data[r - 1]?.[c - 1] as any | undefined;
                    const merge = mergeTopLeft.get(ref);

                    const attrs: string[] = [];
                    if (merge) {
                        if (merge.rowspan > 1) attrs.push(`rowspan="${merge.rowspan}"`);
                        if (merge.colspan > 1) attrs.push(`colspan="${merge.colspan}"`);
                    }

                    let inlineStyle = '';
                    const st = cell?.style;
                    if (st) {
                        if (st.hAlign) inlineStyle += `text-align:${st.hAlign};`;
                        if (st.vAlign) inlineStyle += `vertical-align:${st.vAlign};`;
                        inlineStyle += st.wrapText ? 'white-space:pre-wrap;' : 'white-space:pre;';
                        if (st.fontName) inlineStyle += `font-family:${st.fontName};`;
                        if (st.fontSize) inlineStyle += `font-size:${st.fontSize}px;`;
                        if (st.bold) inlineStyle += 'font-weight:bold;';
                        if (st.italic) inlineStyle += 'font-style:italic;';
                        if (st.fontColor) inlineStyle += `color:${st.fontColor};`;
                        // Prefer foreground color for solid fills, fallback to bgColor
                        if (st.fgColor) inlineStyle += `background-color:${st.fgColor};`;
                        else if (st.bgColor) inlineStyle += `background-color:${st.bgColor};`;
                        // Borders
                        if (st.border) {
                            if (st.border.top?.style) inlineStyle += `border-top:${cssBorder(st.border.top.style, st.border.top.color)};`;
                            if (st.border.right?.style) inlineStyle += `border-right:${cssBorder(st.border.right.style, st.border.right.color)};`;
                            if (st.border.bottom?.style) inlineStyle += `border-bottom:${cssBorder(st.border.bottom.style, st.border.bottom.color)};`;
                            if (st.border.left?.style) inlineStyle += `border-left:${cssBorder(st.border.left.style, st.border.left.color)};`;
                        }
                    }

                    const display = cell?.value ?? '';
                    html += `<td class="xl-cell"${attrs.length ? ' ' + attrs.join(' ') : ''}${inlineStyle ? ` style="${inlineStyle}"` : ''}>${escapeHtml(String(display))}</td>`;
                }
                html += `</tr>`;
            }
            html += `</tbody>`;
            html += `</table>`;
            html += `</div>`; // .xl-sheet
        }

        html += `</div>`; // .xlwb
        return html;
    }

    /**
     * Convert and render only a single sheet from the workbook as HTML.
     * The first sheet has index 0.
     */
    toHTMLSheet(workbook: Workbook, sheetIndex: number): string {
        const colToName = (n: number): string => {
            let s = '';
            while (n > 0) {
                const m = (n - 1) % 26;
                s = String.fromCharCode(65 + m) + s;
                n = Math.floor((n - 1) / 26);
            }
            return s;
        };

        const getEndFromDimension = (dimension: string): { rows: number, cols: number } => {
            if (!dimension) return { rows: 0, cols: 0 };
            const endRef = dimension.includes(':') ? dimension.split(':')[1] : dimension;
            const colLetters = (endRef.match(/[A-Z]+/) || ['A'])[0];
            const rowNum = parseInt((endRef.match(/\d+/) || ['0'])[0], 10) || 0;
            return { rows: rowNum, cols: getColumnIndex(colLetters) };
        };

        const escapeHtml = (value: string): string => {
            return value
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/\"/g, '&quot;')
                .replace(/'/g, '&#39;');
        };

        const styles = `
<style>
.xlwb { font-family: Arial, sans-serif; color: #222; }
.xl-sheet { margin: 12px 0; }
.xl-name { font-weight: 600; margin: 6px 0; }
.xl { border-collapse: collapse; table-layout: fixed; border: 1px solid #d0d7de; }
.xl thead th { position: sticky; top: 0; background: #f6f8fa; z-index: 1; }
.xl th.xl-row { position: sticky; left: 0; background: #f6f8fa; z-index: 1; }
.xl th, .xl td { border: 1px solid #d0d7de; padding: 4px 6px; white-space: pre; }
.xl th { text-align: center; font-weight: 600; font-size: 12px; color: #57606a; }
.xl td { background: #fff; font-size: 13px; }
.xl .xl-corner { background: #f6f8fa; width: 36px; min-width: 36px; }
.xl .xl-col { width: 96px; min-width: 64px; }
.xl th.xl-row { width: 36px; min-width: 36px; }
/* Overlay for drawings */
.xl-wrap { position: relative; display: inline-block; }
.xl-abs { position: absolute; left: 0; top: 0; z-index: 2; pointer-events: none; }
.xl-abs img { position: absolute; image-rendering: auto; object-fit: contain; }
</style>`;

        const sheet = workbook.workSheets[sheetIndex];
        if (!sheet) return '';

        const { rows, cols } = getEndFromDimension(sheet.dimention);
        let nRows = rows;
        let nCols = cols;
        let html = `<div class=\"xlwb\">${styles}`;

        if (rows === 0 || cols === 0) {
            html += `<div class=\"xl-sheet\"><div class=\"xl-name\">${escapeHtml(sheet.name)}</div><div>(hoja vacía)</div></div>`;
            html += `</div>`;
            return html;
        }

        // Build merge maps
        const mergeTopLeft = new Map<string, { rowspan: number; colspan: number }>();
        const coveredCells = new Set<string>();
        for (const range of sheet.mergeCells) {
            const [start, end] = range.split(':');
            const sCol = (start.match(/[A-Z]+/) || ['A'])[0];
            const sRow = parseInt((start.match(/\d+/) || ['1'])[0], 10);
            const eCol = (end.match(/[A-Z]+/) || ['A'])[0];
            const eRow = parseInt((end.match(/\d+/) || ['1'])[0], 10);

            const cStart = getColumnIndex(sCol);
            const cEnd = getColumnIndex(eCol);
            const colspan = Math.max(1, cEnd - cStart + 1);
            const rowspan = Math.max(1, eRow - sRow + 1);

            mergeTopLeft.set(`${sCol}${sRow}`, { rowspan, colspan });
            for (let r = sRow; r <= eRow; r++) {
                for (let c = cStart; c <= cEnd; c++) {
                    const ref = `${colToName(c)}${r}`;
                    if (ref !== `${sCol}${sRow}`) coveredCells.add(ref);
                }
            }
        }

        html += `<div class=\"xl-sheet\">`;
        html += `<div class=\"xl-name\">${escapeHtml(sheet.name)}</div>`;
        html += `<div class=\"xl-wrap\">`;
        // Prepare pixel maps based on column/row styles for better alignment
        const CELL_W = 96; // px per column (matches .xl-col width)
        const CELL_H = 20; // px per row (approx.)
        const findRowHeightUnits = (idx1: number): number => {
            const st = sheet.rowStyles.find(rs => rs.r === idx1);
            if (st && (st.hidden || st.collapsed)) return 0;
            return (st && st.height > 0 ? st.height : sheet.defaultRowHeight) || sheet.defaultRowHeight || 15;
        };
        const measureMDW = (): number => {
            try {
                const fontName = sheet.defaultFontName || 'Calibri';
                const fontPx = sheet.defaultFontSize || 11;
                const cacheKey = `${fontName}-${fontPx}`;
                const g: any = (globalThis as any);
                g.__mdwCache = g.__mdwCache || {};
                if (g.__mdwCache[cacheKey]) return g.__mdwCache[cacheKey];

                // Prefer DOM measurement for installed fonts
                const span = document.createElement('span');
                span.style.position = 'absolute';
                span.style.visibility = 'hidden';
                span.style.whiteSpace = 'pre';
                span.style.fontFamily = `${fontName}, sans-serif`;
                span.style.fontSize = `${fontPx}px`;
                const sample = '0'.repeat(200);
                span.textContent = sample;
                document.body.appendChild(span);
                const wDom = span.getBoundingClientRect().width / sample.length;
                span.remove();
                if (!isNaN(wDom) && wDom > 0) {
                    g.__mdwCache[cacheKey] = wDom;
                    return wDom;
                }

                // Fallback to canvas measurement
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                if (ctx) {
                    ctx.font = `${fontPx}px ${fontName}, sans-serif`;
                    const w0 = ctx.measureText('0'.repeat(200)).width / 200;
                    if (!isNaN(w0) && w0 > 0) {
                        g.__mdwCache[cacheKey] = w0;
                        return w0;
                    }
                }
            } catch {}
            return 7;
        };
        const MDW_RUNTIME = measureMDW();
        const colWidthUnitsToPx = (w: number): number => {
            if (w <= 0) return 0;
            const PADDING = 5;
            if (w < 1) {
                return Math.max(0, Math.floor(w * (MDW_RUNTIME + PADDING)));
            }
            return Math.max(0, Math.floor(w * MDW_RUNTIME + PADDING));
        };
        const rowHeightPtToPx = (pt: number): number => Math.max(0, Math.round(pt * (96 / 72)));
        // Start with default width for all columns, then override in document order (last wins)
        const defaultColPx = colWidthUnitsToPx(sheet.defaultColWidth || 8.43);
        const colWidthsPx: number[] = new Array(nCols + 1).fill(defaultColPx);
        sheet.columnStyles.forEach(cs => {
            const px = (cs.hidden || cs.collapsed) ? 0 : colWidthUnitsToPx(cs.width > 0 ? cs.width : (sheet.defaultColWidth || 8.43));
            const start = Math.max(1, cs.min);
            const end = Math.min(nCols, cs.max);
            for (let c = start; c <= end; c++) {
                colWidthsPx[c] = px;
            }
        });
        const rowHeightsPx: number[] = new Array(nRows + 1).fill(0);
        for (let i = 1; i <= nRows; i++) rowHeightsPx[i] = rowHeightPtToPx(findRowHeightUnits(i));

        // Expand rows/cols if drawings extend beyond current dimension
        const BORDER_W = 1;
        const EMU_PER_PX = 9525;
        const sumColsPx = (count: number): number => {
            let s = 0;
            for (let i = 1; i <= count; i++) s += (colWidthsPx[i] || defaultColPx);
            return s + count * BORDER_W;
        };
        const sumRowsPx = (count: number): number => {
            let s = 0;
            for (let i = 1; i <= count; i++) s += (rowHeightsPx[i] || rowHeightPtToPx(findRowHeightUnits(i)));
            return s + count * BORDER_W;
        };
        const defaultRowPx = rowHeightPtToPx(sheet.defaultRowHeight || 15);
        if (sheet.drawings && sheet.drawings.length > 0) {
            for (const d of sheet.drawings as any[]) {
                const from = d.position.from;
                const to = d.position.to;
                let leftPx = 0;
                for (let c = 1; c <= from.col; c++) leftPx += (colWidthsPx[c] || defaultColPx);
                leftPx += from.col * BORDER_W + Math.floor((from.colOff || 0) / EMU_PER_PX);
                let topPx = 0;
                for (let r = 1; r <= from.row; r++) topPx += (rowHeightsPx[r] || defaultRowPx);
                topPx += from.row * BORDER_W + Math.floor((from.rowOff || 0) / EMU_PER_PX);
                let widthPx = 0;
                let heightPx = 0;
                if (d.sizeEMU) {
                    widthPx = Math.max(1, Math.floor(d.sizeEMU.cx / EMU_PER_PX));
                    heightPx = Math.max(1, Math.floor(d.sizeEMU.cy / EMU_PER_PX));
                } else {
                    if (to.col > from.col) {
                        for (let c = from.col + 1; c <= to.col; c++) widthPx += (colWidthsPx[c] || defaultColPx);
                        widthPx += (to.col - from.col) * BORDER_W;
                    } else if (to.col === from.col) {
                        widthPx = Math.floor(((to.colOff || 0) - (from.colOff || 0)) / EMU_PER_PX);
                    }
                    if (to.row > from.row) {
                        for (let r = from.row + 1; r <= to.row; r++) heightPx += (rowHeightsPx[r] || defaultRowPx);
                        heightPx += (to.row - from.row) * BORDER_W;
                    } else if (to.row === from.row) {
                        heightPx = Math.floor(((to.rowOff || 0) - (from.rowOff || 0)) / EMU_PER_PX);
                    }
                }
                const rightNeeded = leftPx + widthPx;
                while (sumColsPx(nCols) < rightNeeded) {
                    nCols += 1;
                    colWidthsPx[nCols] = defaultColPx;
                }
                const bottomNeeded = topPx + heightPx;
                while (sumRowsPx(nRows) < bottomNeeded) {
                    nRows += 1;
                    rowHeightsPx[nRows] = defaultRowPx;
                }
            }
        }
        html += `<table class=\"xl\">`;
        // Column widths via colgroup (include row header column first)
        html += `<colgroup>`;
        html += `<col style=\"width:36px;min-width:36px;max-width:36px\">`;
        for (let c = 1; c <= nCols; c++) {
            const w = colWidthsPx[c];
            html += `<col style=\"width:${w}px;min-width:${w}px;max-width:${w}px\">`;
        }
        html += `</colgroup>`;

            // Header row with column letters
            html += `<thead><tr>`;
            html += `<th class=\"xl-corner\"></th>`;
            for (let c = 1; c <= nCols; c++) {
            html += `<th class=\"xl-col\" data-col=\"${c}\">${colToName(c)}</th>`;
            }
            html += `</tr></thead>`;

        // Body with row numbers and cells
        html += `<tbody>`;
        const cssBorder = (style: string, color: string): string => {
            if (!style) return '';
            const map: Record<string, string> = {
                hair: '1px solid',
                thin: '1px solid',
                dotted: '1px dotted',
                dashed: '1px dashed',
                dashDot: '1px dashed',
                dashDotDot: '1px dashed',
                medium: '2px solid',
                mediumDashed: '2px dashed',
                mediumDashDot: '2px dashed',
                mediumDashDotDot: '2px dashed',
                thick: '3px solid',
                double: '3px double',
                slantDashDot: '2px dashed',
            };
            const widthStyle = map[style] || '1px solid';
            return `${widthStyle} ${color || '#000'}`;
        };

        for (let r = 1; r <= nRows; r++) {
            // try to reflect row height visually for better overlay alignment
            const trH = rowHeightsPx ? (rowHeightsPx[r] || CELL_H) : CELL_H;
            html += `<tr style=\"height:${trH}px\">`;
            html += `<th class=\"xl-row\" data-row=\"${r}\">${r}</th>`;
            for (let c = 1; c <= nCols; c++) {
                const ref = `${colToName(c)}${r}`;
                if (coveredCells.has(ref)) continue; // skip cells covered by a merge

                const cell = sheet.data[r - 1]?.[c - 1] as any | undefined;
                const merge = mergeTopLeft.get(ref);

                const attrs: string[] = [];
                if (merge) {
                    if (merge.rowspan > 1) attrs.push(`rowspan=\"${merge.rowspan}\"`);
                    if (merge.colspan > 1) attrs.push(`colspan=\"${merge.colspan}\"`);
                }

                let inlineStyle = '';
                const st = cell?.style;
                if (st) {
                    if (st.hAlign) inlineStyle += `text-align:${st.hAlign};`;
                    if (st.vAlign) inlineStyle += `vertical-align:${st.vAlign};`;
                    inlineStyle += st.wrapText ? 'white-space:pre-wrap;' : 'white-space:pre;';
                    if (st.fontName) inlineStyle += `font-family:${st.fontName};`;
                    if (st.fontSize) inlineStyle += `font-size:${st.fontSize}px;`;
                    if (st.bold) inlineStyle += 'font-weight:bold;';
                    if (st.italic) inlineStyle += 'font-style:italic;';
                    if (st.fontColor) inlineStyle += `color:${st.fontColor};`;
                    // Prefer foreground color for solid fills, fallback to bgColor
                    if (st.fgColor) inlineStyle += `background-color:${st.fgColor};`;
                    else if (st.bgColor) inlineStyle += `background-color:${st.bgColor};`;
                    // Borders
                    if (st.border) {
                        if (st.border.top?.style) inlineStyle += `border-top:${cssBorder(st.border.top.style, st.border.top.color)};`;
                        if (st.border.right?.style) inlineStyle += `border-right:${cssBorder(st.border.right.style, st.border.right.color)};`;
                        if (st.border.bottom?.style) inlineStyle += `border-bottom:${cssBorder(st.border.bottom.style, st.border.bottom.color)};`;
                        if (st.border.left?.style) inlineStyle += `border-left:${cssBorder(st.border.left.style, st.border.left.color)};`;
                    }
                }

                const display = cell?.value ?? '';
                html += `<td class=\"xl-cell\" data-ref=\"${ref}\"${attrs.length ? ' ' + attrs.join(' ') : ''}${inlineStyle ? ` style=\"${inlineStyle}\"` : ''}>${escapeHtml(String(display))}</td>`;
            }
            html += `</tr>`;
        }
        html += `</tbody>`;
        html += `</table>`;

        // Drawings overlay (images, shapes, textboxes)
        const HEAD_W = 36; // row header width (matches .xl .xl-corner)
        const HEAD_H = 28; // header row height (approx.)
        const drawings = sheet.drawings.filter(d => (d.type === 'image' || d.type === 'shape' || d.type === 'textbox'));
        if (drawings.length > 0) {
            html += `<div class=\"xl-abs\">`;
            const computeRect = (d: any) => {
                const from = d.position.from;
                const to = d.position.to;
                let left = HEAD_W;
                let top = HEAD_H;
                if (d.anchorType === 'absolute' && d.absEMU) {
                    left += Math.floor(d.absEMU.x / EMU_PER_PX);
                    top += Math.floor(d.absEMU.y / EMU_PER_PX);
                } else {
                    for (let c = 1; c <= from.col; c++) left += colWidthsPx[c];
                    left += from.col * BORDER_W;
                    left += Math.floor((from.colOff || 0) / EMU_PER_PX);
                    for (let r = 1; r <= from.row; r++) top += rowHeightsPx[r];
                    top += from.row * BORDER_W;
                    top += Math.floor((from.rowOff || 0) / EMU_PER_PX);
                }
                let width = 0;
                let height = 0;
                if (d.sizeEMU) {
                    width = Math.max(1, Math.floor(d.sizeEMU.cx / EMU_PER_PX));
                    height = Math.max(1, Math.floor(d.sizeEMU.cy / EMU_PER_PX));
                } else {
                    if (to.col > from.col) {
                        for (let c = from.col + 1; c <= to.col; c++) width += colWidthsPx[c];
                        width += (to.col - from.col) * BORDER_W;
                    } else if (to.col === from.col) {
                        width = Math.floor(((to.colOff || 0) - (from.colOff || 0)) / EMU_PER_PX);
                    }
                    if (to.row > from.row) {
                        for (let r = from.row + 1; r <= to.row; r++) height += rowHeightsPx[r];
                        height += (to.row - from.row) * BORDER_W;
                    } else if (to.row === from.row) {
                        height = Math.floor(((to.rowOff || 0) - (from.rowOff || 0)) / EMU_PER_PX);
                    }
                }
                return { left, top, width: Math.max(1, width), height: Math.max(1, height) };
            };
            for (const d of drawings) {
                const { left, top, width, height } = computeRect(d as any);
                if (d.type === 'image') {
                    const b64 = d.base64;
                    const mime = ((): string => {
                        if (b64.startsWith('/9j/')) return 'image/jpeg';
                        if (b64.startsWith('iVBOR')) return 'image/png';
                        if (b64.startsWith('R0lGOD')) return 'image/gif';
                        return 'image/*';
                    })();
                    html += `<img alt=\"\" src=\"data:${mime};base64,${b64}\" style=\"left:${left}px;top:${top}px;width:${width}px;height:${height}px;\" />`;
                } else if (d.type === 'shape') {
                    const fill = d.properties?.fillColor || 'transparent';
                    const radius = d.properties?.shapeType === 'roundRect' ? '8px' : '0';
                    const lineColor = d.properties?.lineColor || 'rgba(0,0,0,0.3)';
                    const lineWidth = d.properties?.lineWidth ? `${Math.max(1, parseInt(d.properties.lineWidth,10)/12700)}px` : '1px';
                    html += `<div class=\"xl-shape\" style=\"position:absolute;left:${left}px;top:${top}px;width:${width}px;height:${height}px;background:${fill};border-radius:${radius};border:${lineWidth} solid ${lineColor};\"></div>`;
                } else if (d.type === 'textbox') {
                    const text = d.properties?.text || '';
                    const fill = d.properties?.fillColor || 'transparent';
                    const lineColor = d.properties?.lineColor || 'transparent';
                    const lineWidth = d.properties?.lineWidth ? `${Math.max(1, parseInt(d.properties.lineWidth,10)/12700)}px` : '0px';
                    const color = d.properties?.textColor || '#222';
                    const fontPt = d.properties?.fontPt || 11;
                    const fontPx = Math.round(fontPt * (96/72));
                    const fw = d.properties?.bold ? 'bold' : 'normal';
                    const fs = d.properties?.italic ? 'italic' : 'normal';
                    const ta = d.properties?.textAlign || 'left';
                    html += `<div class=\"xl-textbox\" style=\"position:absolute;left:${left}px;top:${top}px;width:${width}px;height:${height}px;background:${fill};border:${lineWidth} solid ${lineColor};white-space:pre-wrap;padding:2px;pointer-events:auto;color:${color};font-size:${fontPx}px;font-weight:${fw};font-style:${fs};text-align:${ta};\">${escapeHtml(text)}\n</div>`;
                }
            }
            html += `</div>`; // .xl-abs
        }

        html += `</div>`; // .xl-wrap
        html += `</div>`; // .xl-sheet

        html += `</div>`; // .xlwb
        return html;
    }
    
    /**
     * Extract data from spreadsheet bytes stored in a `ArrayBuffer`.
     * @param file Workbook file
     * @param options Parser options
     * @returns 
     */
    async readFile (file: ArrayBuffer, options: XlsxParserOptions = { dense: false, styles: false, drawings: false, skipHiddenRows: false }) {
        const nZip = new JSZip();
        const result = await nZip.loadAsync(file);
        const files: string[] = [];
        result.forEach((p) => files.push(p));
        const sheetFiles = files.filter(x => x.includes('xl/worksheets/sheet'));
        const drawingFiles = options.drawings ? files.filter(x => x.includes('xl/drawings/drawing')) : [];
        const drawingRelFiles = options.drawings ? files.filter(x => x.includes('xl/drawings/_rels/drawing')) : [];
        const mediaFiles = options.drawings ? files.filter(x => x.includes('xl/media/')) : [];
        
        let sharedStrings: string[] = [];
        let themes: Theme[] = [];
        let drawings: DrawingFile[] = [];
        const media: MediaFile[] = [];
        let style: StyleSheet = {
            fonts: [],
            fills: [],
            borders: [],
            cells: [],
        };
        let workbook: Workbook = {
            workSheets: [],
        };
       
        for(const p of files) {
            // Read workbook.xml
            if (p === 'xl/workbook.xml') {
                const xml = await nZip.file(p)?.async('string');
                if (xml) {
                    const nWorkbook = parseWorkbookXml(xml);
                    workbook = nWorkbook;
                }
            }
            // Read sharedStrings.xml
            if (p === 'xl/sharedStrings.xml') {
                const xml = await nZip.file(p)?.async('string');
                if (xml) {
                    const nSharedStrings = parseSharedStringsXml(xml);
                    sharedStrings = nSharedStrings;
                }
            }
            // Read theme1.xml
            if (p === 'xl/theme/theme1.xml' && (options.styles ?? false)) {
                const xml = await nZip.file(p)?.async('string');
                if (xml) {
                    const nThemes = parseThemeXml(xml);
                    themes = nThemes;
                }
            }
            // Read styles.xml
            if (p === 'xl/styles.xml' && (options.styles ?? false)) {
                const xml = await nZip.file(p)?.async('string');
                if (xml) {
                    const nStyle = parseStylesXml(xml, themes);
                    style = nStyle;
                }
            }    
        }

        for(let i=0; i < mediaFiles.length; i++) {
            // Read image[n].png
            const base64 = (await nZip.file(mediaFiles[i])?.async('base64')) ?? '';
            media.push({
                name: mediaFiles[i].split('/').pop() || '',
                base64: base64,
            });
        }

        for(let i=0; i < drawingFiles.length; i++) {
            // Read drawing[n].xml
            const xml = await nZip.file(drawingFiles[i])?.async('string');
            // Read _rels/drawing[n].xml.rels
            const relName = drawingFiles[i].split('/').pop() || '';
            const relFilePath = drawingRelFiles.find(x => x.includes(`${relName}.rels`)) ?? '';
            const rel = await nZip.file(relFilePath)?.async('string');

            if (xml) {
                const nDrawing = parseDrawingXml(xml, rel ?? '', media, themes);
                drawings.push({ src: relName, drawings: nDrawing});
            }
        }

        for(let i=0; i < sheetFiles.length; i++) {
            // Read sheet[n].xml
            const xml = await nZip.file(sheetFiles[i])?.async('string');
            // Read _rels/sheet[n].xml.rels
            const sheetName = sheetFiles[i].split('/').pop() || '';
            const relFilePath = `xl/worksheets/_rels/${sheetName}.rels`;
            const rel = await nZip.file(relFilePath)?.async('string');
            if (xml) {
                const nSheet = parseWorksheetXml(xml, rel, style, sharedStrings, drawings, options.dense ?? false, options.skipHiddenRows ?? false);
                workbook.workSheets[i].dimention = nSheet.dimention;
                workbook.workSheets[i].columnStyles = nSheet.columnStyles;
                workbook.workSheets[i].rowStyles = nSheet.rowStyles;
                workbook.workSheets[i].data = nSheet.data;
                workbook.workSheets[i].mergeCells = nSheet.mergeCells;
                workbook.workSheets[i].drawings = nSheet.drawings;
                workbook.workSheets[i].defaultColWidth = nSheet.defaultColWidth;
                workbook.workSheets[i].baseColWidth = nSheet.baseColWidth;
                workbook.workSheets[i].defaultRowHeight = nSheet.defaultRowHeight;
                workbook.workSheets[i].zeroHeight = nSheet.zeroHeight;
                workbook.workSheets[i].defaultFontName = nSheet.defaultFontName;
                workbook.workSheets[i].defaultFontSize = nSheet.defaultFontSize;
            }
        }
        
        return workbook;
    }
}
