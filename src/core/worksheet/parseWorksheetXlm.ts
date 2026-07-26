import { DrawingFile } from "../drawing/types";
import { StyleSheet } from "../style/types";
import { excelSerialToJSDate, getElementByName, getElementsByName, getPositionInArray, getRangeArray, getSheetDimension } from "../utils";
import { ColStyle, WorkSheet } from "./types";

export const parseWorksheetXml = (str: string, relStr: string | undefined, styleSheet: StyleSheet, sharedStrings: string[], drawingFiles: DrawingFile[], dense: boolean, skipHiddenRows: boolean): WorkSheet => {
    const worksheet: WorkSheet = {
        id: 0,
        name: '',
        dimention: '',
        data: [],
        columnStyles: [],
        rowStyles: [],
        mergeCells: [],
        drawings: [],
        defaultColWidth: 8.43,
        baseColWidth: 8,
        defaultRowHeight: 15,
        zeroHeight: false,
    };
    const xmlDoc = new DOMParser().parseFromString(str, 'text/xml');
    const xmlRel = relStr ? new DOMParser().parseFromString(relStr, 'text/xml') : undefined;
    const worksheetElement = getElementByName(xmlDoc, 'worksheet');
    const dimensionElement = getElementByName(worksheetElement, 'dimension');
    const sheetDataElement = getElementByName(worksheetElement, 'sheetData');
    const mergeCellsElement = getElementByName(worksheetElement, 'mergeCells');
    const colsElement = getElementByName(worksheetElement, 'cols');
    const sheetFormatPrElement = getElementByName(worksheetElement, 'sheetFormatPr');
    const drawingElement = getElementsByName(xmlRel, 'Relationship');
    const colsArray = getElementsByName(colsElement, 'col');
    const rowsArray = getElementsByName(sheetDataElement, 'row');
    const mergeCellArray = getElementsByName(mergeCellsElement, 'mergeCell');
    

    if (sheetDataElement && !dimensionElement) {
        worksheet.dimention = getSheetDimension(sheetDataElement);
    }

    if (dimensionElement) {
        worksheet.dimention = dimensionElement.getAttribute('ref') ?? '';
    }

    if (worksheet.dimention !== '') {
        worksheet.data = getRangeArray(
            worksheet.dimention,
            dense ? { ref: '', value: '', formula: '' } : undefined
        );
    }

    if (sheetFormatPrElement) {
        const baseColWidthAttr = sheetFormatPrElement.getAttribute('baseColWidth');
        const defaultColWidthAttr = sheetFormatPrElement.getAttribute('defaultColWidth');
        const defaultRowHeightAttr = sheetFormatPrElement.getAttribute('defaultRowHeight');

        if (baseColWidthAttr !== null) worksheet.baseColWidth = +baseColWidthAttr;
        if (defaultColWidthAttr !== null) worksheet.defaultColWidth = +defaultColWidthAttr;
        if (defaultRowHeightAttr !== null) worksheet.defaultRowHeight = +defaultRowHeightAttr;

        worksheet.zeroHeight = sheetFormatPrElement.getAttribute('zeroHeight') === "1";
    }

    const isTrue = (value: string | null): boolean => value === '1' || value === 'true';

    if (colsArray) {
        const columnStyles: ColStyle[] = [];
        colsArray.forEach(x => {
            columnStyles.push({
                min: +(x.getAttribute('min') ?? 1),
                max: +(x.getAttribute('max') ?? 1),
                width: +(x.getAttribute('width') ?? 1),
                hidden: isTrue(x.getAttribute('hidden')),
                collapsed: isTrue(x.getAttribute('collapsed')),
            });
        });
        worksheet.columnStyles = columnStyles;
    }

    if (rowsArray) {
        rowsArray.forEach(x => {
            const index = +(x.getAttribute('r') ?? 0) - 1;
            const hiddenProp = isTrue(x.getAttribute('hidden'));
            const collapsedProp = isTrue(x.getAttribute('collapsed'));
    
    // Propagate default font info if available from stylesheet
    try {
        // styleSheet is in closure via parameter
        // @ts-ignore
        if ((styleSheet as any).defaultFontName) worksheet.defaultFontName = (styleSheet as any).defaultFontName;
        // @ts-ignore
        if ((styleSheet as any).defaultFontSize) worksheet.defaultFontSize = (styleSheet as any).defaultFontSize;
    } catch {}

            if (index >= 0 && (!skipHiddenRows || (!hiddenProp && !collapsedProp))) {
                const cols = getElementsByName(x, 'c');
                cols.forEach(y => {
                    const r = y.getAttribute('r') ?? 'A'; // Row
                    const s = +(y.getAttribute('s') ?? -1); // Style
                    const t = y.getAttribute('t') ?? ''; // Type
                    const formula = getElementByName(y, 'f')?.textContent ?? '';
                    const value = t === 'inlineStr'
                        ? getElementsByName(getElementByName(y, 'is'), 't').map(text => text.textContent ?? '').join('')
                        : getElementByName(y, 'v')?.textContent ?? '';

                    const pos = getPositionInArray(r);
                    const cellStyle = s >= 0 ? styleSheet.cells[s] : undefined;
                    const font = cellStyle ? styleSheet.fonts[cellStyle.fontId] : undefined;
                    const fill = cellStyle ? styleSheet.fills[cellStyle.fillId] : undefined;

                    worksheet.data[pos.row][pos.col] = {
                        ref: r,
                        value: t !== 's' 
                            ? (
                                (value !== '' && !isNaN(+value) && cellStyle?.numFmtId === 14)
                                    ? excelSerialToJSDate(+value).toLocaleDateString() // Date
                                    : value // Number
                            ) 
                            : sharedStrings[+value],
                        formula: formula,
                        style: (cellStyle && font && fill)
                            ? {
                                bgColor: fill.bgColor,
                                fgColor: fill.fgColor,
                                fontName: font.name,
                                fontSize: font.size,
                                fontColor: font.color,
                                bold: font.bold,
                                italic: font.italic,
                                vAlign: cellStyle.alignment?.vertical ?? 'bottom',
                                hAlign: cellStyle.alignment?.horizontal ?? '',
                                wrapText: cellStyle.alignment?.wrapText ?? false,
                                border: styleSheet.borders[cellStyle.borderId],
                            }
                            : undefined,
                    }

                    if (worksheet.data[pos.row][pos.col]?.style?.hAlign === '') {
                        worksheet.data[pos.row][pos.col]!.style!.hAlign = isNaN(+worksheet.data[pos.row][pos.col].value) ? 'left' : 'right';
                    }
                });

                // Row style
                if (!skipHiddenRows || (cols.length > 0)) {
                    worksheet.rowStyles.push({
                        r: +(x.getAttribute('r') ?? 0),
                        height: +(x.getAttribute('ht') ?? worksheet.defaultRowHeight),
                        hidden: hiddenProp,
                        collapsed: collapsedProp,
                    });
                }
            }
        });

        if (mergeCellArray) {
            mergeCellArray.forEach(x => {
                worksheet.mergeCells.push(x.getAttribute('ref') ?? '');
            });
        }
    }

    if (skipHiddenRows) {
        if (worksheet.defaultRowHeight === 0 && worksheet.zeroHeight) {
            worksheet.data = worksheet.data.filter(x => x.length > 0 && x[0] !== undefined && x.some(y => y.ref !== ''));
        }
    }

    const drawingRel = drawingElement.find(rel =>
        rel.getAttribute('Type')?.includes('drawing')
    );

    if (drawingRel) {
        const drawingTarget = drawingRel.getAttribute('Target') || '';
        const drawingFileName = drawingTarget.split('/').pop() || '';

        const drawingFile = drawingFiles.find(df => df.src === drawingFileName);
        worksheet.drawings = drawingFile?.drawings ?? [];
    }
    
    return worksheet;
}
