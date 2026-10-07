import { Drawing } from "../drawing/types";

export interface RowStyle {
    /** Row position. */
    r: number;
    /** Row height. */
    height: number;
    /** Indicates whether the row should be hidden */
    hidden: boolean;
    /** Indicates whether the row should be collapsed */
    collapsed: boolean;
}

export interface ColStyle {
    /** Column from which the style will be applied */
    min: number;
    /** Column to which the style will be applied */
    max: number;
    /** Column width */
    width: number;
    /** Indicates whether the column should be hidden */
    hidden: boolean;
    /** Indicates whether the column should be collapsed */
    collapsed: boolean;
}

interface CellBorderStyle {
    /** The line style for this border. */
    style: string;
    /** Color hex. */
    color: string;
}

interface CellBorder {
    /** Specifies the color and line style for the left border of a cell. */
    left: CellBorderStyle;
    /** Specifies the color and line style for the right border of a cell. */
    right: CellBorderStyle;
    /** Specifies the color and line style for the top border of a cell. */
    top: CellBorderStyle;
    /** Specifies the color and line style for the bottom border of a cell. */
    bottom: CellBorderStyle;
    /** Specifies the color and line style for the diagonal border of a cell. */
    diagonal: CellBorderStyle;
}

interface CellStyle {
    /** Background color of the cell fill pattern. */
    bgColor: string;
    /** Foreground color of the cell fill pattern. */
    fgColor: string;
    /** Face name of this font. */
    fontName: string;
    /** Font size. */
    fontSize: number;
    /** Font color. */
    fontColor: string;
    /** Vertical alignment. */
    vAlign: string;
    /** Horizontal alignment. */
    hAlign: string;
    /** Cell border format (left, right, top, bottom, diagonal). */
    border: CellBorder;
    /** A boolean value indicating if the text in a cell should be line-wrapped within the cell. */
    wrapText: boolean;
    /** Bold. */
    bold: boolean;
    /** Italic. */
    italic: boolean;
}

interface CellData {
    /** Cell reference. */
    ref: string;
    /** Cell value. */
    value: string;
    /** Cell formula. */
    formula: string;
    /** Defines styles applied to the cell. */
    style?: CellStyle;
}

/**
 * A two-dimensional grid of cells that are organized into rows and columns.
 */
export interface WorkSheet {
    /** Drawing relationship discovered during the worksheet's single XML pass. */
    drawingRelationshipId?: string;
    state?: 'visible' | 'hidden' | 'veryHidden';
    /** Internal package relationship used to resolve this sheet. */
    relationshipId?: string;
    /** Preserve coordinates while omitting hidden rows from HTML. */
    skipHiddenRows?: boolean;
    /** Sheet id. */
    id: number;
    /** Sheet name. */
    name: string;
    /** Specifies the used range of the worksheet. */
    dimention: string;
    /** Cell table. */
    data: (CellData)[][];
    /** Column styles. */
    columnStyles: ColStyle[];
    /** Row styles. */
    rowStyles: RowStyle[];
    /** Merge cells. */
    mergeCells: string[];
    /** Default column width */
    defaultColWidth: number;
    /** Base column width (characters of '0' in Normal font) */
    baseColWidth: number;
    /** Default row height */
    defaultRowHeight: number;
    /** 'true' if rows are hidden by default */
    zeroHeight: boolean;
    /** Drawings */
    drawings: Drawing[];
    /** Default (Normal) font family if resolvable */
    defaultFontName?: string;
    /** Default (Normal) font size in points if resolvable */
    defaultFontSize?: number;
}
