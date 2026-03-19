# Xlsx-to-js

A TypeScript-based library for parsing Excel (XLSX) with browser support.

[Live demo](https://javier-mora.github.io/xlsx-to-js/)

## Getting Started
### Installation

With [npm](https://www.npmjs.com/package/xlsx-to-js):

```bash
npm install --save xlsx-to-js
```
Import library:
```javascript
import { XlsxParser } from "xlsx-to-js";
```
## Usage

### Parsing Workbooks

Extract data from spreadsheet bytes
```javascript
const xlsxParser = new XlsxParser();
const workbook = await xlsxParser.readFile(file, { dense: true, styles: true, drawings: true, skipHiddenRows: true });
```
The `readFile` method extract data from spreadsheet bytes stored in a `ArrayBuffer`.

The second argument to `options` accepts the properties:
|Option         |Default    |Description|
|---------------|-----------|-----------|
|dense          |false      | When the option `dense: false` is passed, parsers will skip empty cells. |
|styles         |false      | When the opction `styles: false` is passed, parsers will skip cell styles. |
|drawings       |false      | When the option `drawings: false` is passed, parsers will skip parsing drawings and graphical objects. |
|skipHiddenRows |false      | When the option `skipHiddenRows: true` is passed, hidden rows will be ignored during parsing. |


### Render to HTML

Render the parsed workbook as Excel-like HTML.

- Render all sheets at once:
```ts
const xlsxParser = new XlsxParser();
const workbook = await xlsxParser.readFile(file, {
  dense: true,
  styles: true,
  drawings: true,
  skipHiddenRows: true,
});

const fullHtml = xlsxParser.toHTML(workbook);
document.getElementById('container')!.innerHTML = fullHtml;
```

- Render a single sheet (recommended for performance in UIs with tabs):
```ts
const xlsxParser = new XlsxParser();
const workbook = await xlsxParser.readFile(file, {
  dense: true,
  styles: true,
  drawings: true,
  skipHiddenRows: true,
});

// Render first sheet (index 0)
const sheetHtml = xlsxParser.toHTMLSheet(workbook, 0);
document.getElementById('sheetView')!.innerHTML = sheetHtml;
```

Notes:
- Pass `styles: true` to include cell fonts, colors, alignment, borders, and fills.
- Pass `drawings: true` to include images, shapes, and textboxes positioned like in Excel.
- The generated HTML includes column letters and row numbers. It also respects merged cells and most layout details.

**Supported Features**
- **Cell Content:** strings, numbers, dates (basic serial-date -> locale string), formulas (stored, not evaluated).
- **Merged Cells:** respects merge ranges and renders proper `rowspan/colspan`.
- **Styles:** font family/size, bold/italic, text color, background fill (theme, rgb), horizontal/vertical alignment, wrap, borders (most styles) when `styles: true`.
- **Dimensions:** column widths and row heights converted to pixels with Excel-like logic; honors per-column/row overrides and hidden/collapsed.
- **Headers:** row numbers and column letters like Excel.
- **Drawings:** images (png/jpg/gif), shapes (fill/border), textboxes (text color, align, font), positioned via anchors when `drawings: true`.
- **Multiple Sheets:** render all or one at a time (`toHTMLSheet`).

**Limitations**
- **No Calc Engine:** formulas are not evaluated; cell `.formula` is exposed, `.value` is parsed text/number/date.
- **Styles Fidelity:** border variants, distributed/justify vertical alignment, and some number formats may not fully match Excel.
- **Themes/Tint:** theme shade/tint handled pragmáticamente; minor color differences possible versus desktop Excel.
- **Fonts/MDW:** column width conversion depends on runtime font metrics; small pixel drifts may occur across platforms.
- **Drawings Coverage:** connectors, grouped shapes, rotations, and complex effects are not fully rendered.
- **Hidden Rows/Cols:** when `skipHiddenRows: true`, hidden rows are skipped; hidden columns get width 0 but still occupy position.
- **Print/Views:** print areas, panes freeze, page breaks and advanced view options are not applied to HTML.

**Supported Environments**
- **Browsers:** modern Chromium/Firefox/Safari (ES2019+, `DOMParser`, `Canvas` for font metrics). Tested on latest Chrome/Edge/Firefox/Safari.
- **Node.js:** intended for browser use. In Node you must polyfill `DOMParser` (e.g., jsdom) to use XML parsing and HTML rendering.
- **Module Format:** published as ESM; works with bundlers like Vite/Webpack/Rollup.

## References
- ISO/IEC 29500:2012 "Office Open XML File Formats — Fundamentals And Markup Language Reference"
