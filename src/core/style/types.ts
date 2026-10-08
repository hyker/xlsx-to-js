export interface FontStyleSheet {
    /** Font size. */
    size: number;
    /** Color hex. */
    color: string;
    /** Face name of this font. */
    name: string;
    /** Font family number. */
    family: number;
    /** Bold. */
    bold: boolean;
    /** Italic. */
    italic: boolean;
}

export interface FillStyleSheet {
    /** Specifies the fill pattern type (including solid and none) Default is none, when missing. */
    patternType: string;
    /** Foreground color of the cell fill pattern. */
    fgColor: string;
    /** Background color of the cell fill pattern. */
    bgColor: string;
}

export interface BorderOption {
    /** The line style for this border. */
    style: string;
    /** Color hex. */
    color: string;
}

export interface BorderStyleSheet {
    /** Specifies the color and line style for the left border of a cell. */
    left: BorderOption;
    /** Specifies the color and line style for the right border of a cell. */
    right: BorderOption;
    /** Specifies the color and line style for the top border of a cell. */
    top: BorderOption;
    /** Specifies the color and line style for the bottom border of a cell. */
    bottom: BorderOption;
    /** Specifies the color and line style for the diagonal border of a cell. */
    diagonal: BorderOption;
}

export interface AlignmentOption {
    /** Represents the horizontal alignment setting. */
    horizontal: string;
    /** Represents the vertical alignment setting. */
    vertical: string;
    /** A boolean value indicating if the text in a cell should be line-wrapped within the cell. */
    wrapText: boolean;
}

export interface CellStyleSheet {
    /** Id of the number format (numFmt) record used by this cell format. */
    numFmtId: number;
    /** Zero-based index of the font record used by this cell format. */
    fontId: number;
    /** Zero-based index of the fill record used by this cell format. */
    fillId: number;
    /** Zero-based index of the border record used by this cell format. */
    borderId: number;
    /** For xf records contained in cellXfs this is the zero-based index of an xf record contained in cellStyleXfs corresponding to the cell style applied to the cell. */
    xfId: number;
    /** A boolean value indicating whether the font formatting specified for this xf should be applied. */
    applyFont: number;
    /** A boolean value indicating whether the border formatting specified for this xf should be applied. */
    applyBorder: number;
    /** A boolean value indicating whether the alignment formatting specified for this xf should be applied. */
    applyAlignment: number;
    /** A boolean value indicating whether the fill formatting specified for this xf should be applied. */
    applyFill: number;
    /** Formatting information pertaining to text alignment in cells. */
    alignment?: AlignmentOption;
}

export interface StyleSheet {
    /** Contains all font definitions for this workbook. */
    fonts: FontStyleSheet[];
    /** Defines the cell fills portion of the Styles part. */
    fills: FillStyleSheet[];
    /** Expresses a single set of cell border formats (left, right, top, bottom, diagonal). */
    borders: BorderStyleSheet[];
    /** Cell styles. */
    cells: CellStyleSheet[];
    /** Format codes declared in numFmts, by numFmtId. */
    numFmts?: Map<number, string>;
    /** Default (Normal) font family if resolvable. */
    defaultFontName?: string;
    /** Default (Normal) font size in points if resolvable. */
    defaultFontSize?: number;
}
