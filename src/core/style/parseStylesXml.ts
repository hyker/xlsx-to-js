import { parseXml } from '../xml';
import { XlsxLimits } from '../../types';
import { boundedNumber, resolveLimits } from '../security';
import { Theme } from "../theme/types";
import { getElementByName, getElementsByName } from "../utils";
import { getColor } from "./getColor";
import { BorderStyleSheet, StyleSheet } from "./types";


export const parseStylesXml = (str: string, themes: Theme[], limits: XlsxLimits = resolveLimits()): StyleSheet => {
    const styleSheet: StyleSheet = {
        fonts: [],
        fills: [],
        borders: [],
        cells: [],
    };
    const xmlDoc = parseXml(str, 'styleSheet', limits);
    const styleSheetElement = getElementByName(xmlDoc, 'styleSheet');
    const fontsElement = getElementByName(styleSheetElement, 'fonts');
    const fillsElement = getElementByName(styleSheetElement, 'fills');
    const bordersElement = getElementByName(styleSheetElement, 'borders');
    const cellStylesElement = getElementByName(styleSheetElement, 'cellStyles');
    const cellStyleXfsElement = getElementByName(styleSheetElement, 'cellStyleXfs');
    const cellXfsElement = getElementByName(styleSheetElement, 'cellXfs');
    const fontsArray = getElementsByName(fontsElement, 'font');
    const fillsArray = getElementsByName(fillsElement, 'fill');
    const bordersArray = getElementsByName(bordersElement, 'border');
    const xfsArray = getElementsByName(cellXfsElement, 'xf');
    const cellStyleXfsArray = getElementsByName(cellStyleXfsElement, 'xf');
    const cellStylesArray = getElementsByName(cellStylesElement, 'cellStyle');

    if (fontsArray) {
        fontsArray.forEach(x => {
            const sz = getElementByName(x, 'sz');
            const color = getElementByName(x, 'color');
            const name = getElementByName(x, 'name');
            const family = getElementByName(x, 'family');
            const bold = getElementByName(x, 'b');
            const italic = getElementByName(x, 'i');

            styleSheet.fonts.push({
                size: boundedNumber(sz?.getAttribute('val') ?? 0, 'font size', 409),
                color: getColor(color, themes),
                name: name?.getAttribute('val') ?? '',
                family: boundedNumber(family?.getAttribute('val') ?? 0, 'font family', 5, 0, true),
                bold: !!bold && !['0', 'false'].includes(bold.getAttribute('val') ?? '1'),
                italic: !!italic && !['0', 'false'].includes(italic.getAttribute('val') ?? '1'),
            });
        })
    }

    if (fillsArray) {
        fillsArray.forEach(x => {
            const patternFill = getElementByName(x, 'patternFill');
            const fgColor = getElementByName(patternFill, 'fgColor');
            const bgColor = getElementByName(patternFill, 'bgColor');
            styleSheet.fills.push({
                patternType: patternFill?.getAttribute('patternType') ?? '',
                bgColor: getColor(bgColor, themes),
                fgColor: getColor(fgColor, themes),
            });
        })
    }

    if (bordersArray) {
        bordersArray.forEach(x => {
            const border: BorderStyleSheet = {
                left: { style: '', color: '' },
                right: { style: '', color: '' },
                top: { style: '', color: '' },
                bottom: { style: '', color: '' },
                diagonal: { style: '', color: '' },
            };
            const leftBorder = getElementByName(x, 'left');
            const rightBorder = getElementByName(x, 'right');
            const topBorder = getElementByName(x, 'top');
            const bottomBorder = getElementByName(x, 'bottom');
            const diagonalBorder = getElementByName(x, 'diagonal');

            if (leftBorder) {
                border.left.style = leftBorder.getAttribute('style') ?? '';
                border.left.color = getColor(getElementByName(leftBorder, 'color'), themes);
            }
            if (rightBorder) {
                border.right.style = rightBorder.getAttribute('style') ?? '';
                border.right.color = getColor(getElementByName(rightBorder, 'color'), themes);
            }
            if (topBorder) {
                border.top.style = topBorder.getAttribute('style') ?? '';
                border.top.color = getColor(getElementByName(topBorder, 'color'), themes);
            }
            if (bottomBorder) {
                border.bottom.style = bottomBorder.getAttribute('style') ?? '';
                border.bottom.color = getColor(getElementByName(bottomBorder, 'color'), themes);
            }
            if (diagonalBorder) {
                border.diagonal.style = diagonalBorder.getAttribute('style') ?? '';
                border.diagonal.color = getColor(getElementByName(diagonalBorder, 'color'), themes);
            }

            styleSheet.borders.push(border);
        });
    }

    if (xfsArray) {
        xfsArray.forEach(x => {
            const alignment = getElementByName(x, 'alignment');
            const horizontal = alignment?.getAttribute('horizontal') ?? '';
            const vertical = alignment?.getAttribute('vertical') ?? 'bottom';
            const horizontalNorm = horizontal === 'centerContinuous' ? 'center' : (horizontal === 'distributed' ? 'justify' : horizontal);
            styleSheet.cells.push({
                numFmtId: boundedNumber(x.getAttribute('numFmtId') ?? 0, 'numFmtId', 65535, 0, true),
                fontId: boundedNumber(x.getAttribute('fontId') ?? 0, 'fontId', styleSheet.fonts.length - 1, 0, true),
                fillId: boundedNumber(x.getAttribute('fillId') ?? 0, 'fillId', styleSheet.fills.length - 1, 0, true),
                borderId: boundedNumber(x.getAttribute('borderId') ?? 0, 'borderId', styleSheet.borders.length - 1, 0, true),
                xfId: boundedNumber(x.getAttribute('xfId') ?? 0, 'xfId', 65535, 0, true),
                applyFont: +(x.getAttribute('applyFont') ?? 0),
                applyBorder: +(x.getAttribute('applyBorder') ?? 0),
                applyAlignment: +(x.getAttribute('applyAlignment') ?? 0),
                applyFill: +(x.getAttribute('applyFill') ?? 0),
                alignment: alignment
                    ? {
                        horizontal: horizontalNorm,
                        vertical: vertical,
                        wrapText: ['1', 'true'].includes(alignment?.getAttribute('wrapText') ?? '0'),
                    }
                    : undefined, 
            });
        })
    }

    // Resolve Normal style default font
    try {
        const normalStyle = cellStylesArray.find(cs => (cs.getAttribute('name') || '').toLowerCase() === 'normal');
        const xfId = normalStyle ? +(normalStyle.getAttribute('xfId') || '0') : 0;
        const xf = cellStyleXfsArray[xfId];
        const fontId = xf ? +(xf.getAttribute('fontId') || '0') : 0;
        const f = styleSheet.fonts[fontId];
        if (f) {
            styleSheet.defaultFontName = f.name;
            styleSheet.defaultFontSize = f.size;
        }
    } catch {}

    return styleSheet;
};
