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
|limits         |`DEFAULT_LIMITS`| Resource budgets described below. |


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
- Both rendering APIs use the same layout, styles, and drawings.
- Hidden and very-hidden sheets are retained in `workbook.workSheets` with their `state`, but omitted from HTML by default. To display them deliberately, pass `{ includeHiddenSheets: true }` as the second argument to `toHTML`, or the third argument to `toHTMLSheet`. Sheet indexes always refer to the original workbook order.

### Untrusted workbooks and resource limits

The parser rejects malformed XML, DTD/entity declarations, invalid coordinates and indexes, duplicate cells, overlapping merges, invalid required relationships, archive traversal, duplicate archive paths, and ZIP checksum mismatches. Fonts are serialized as quoted CSS strings; alignment and colors are restricted to supported values. Neither formulas nor external relationships are executed or fetched. Embedded images are restricted to PNG, JPEG, and GIF when rendered.

Budgets apply before grid allocation and during streamed decompression, including when ZIP size metadata is forged. Archive entry and declared expansion limits include unused parts. Grid, merge, and drawing budgets cover the whole workbook; render budgets also cover expansion caused by drawings. `dense: false` avoids allocating empty cells but does not bypass grid budgets.

```ts
import { XlsxParser, DEFAULT_LIMITS } from 'xlsx-to-js';

const limits = {
  ...DEFAULT_LIMITS,
  maxFileBytes: 5 * 1024 * 1024,
  maxCells: 100_000,
};
const workbook = await new XlsxParser().readFile(file, {
  styles: true,
  drawings: true,
  limits,
});
const html = new XlsxParser().toHTML(workbook, { limits });
```

Parsing and rendering accept independent `limits` overrides. Pass your chosen budgets to both. Overrides must be positive safe integers; row and column limits cannot exceed Excel's coordinate limits. Raising budgets increases CPU and memory exposure.

| Limit | Default |
|---|---:|
| `maxFileBytes` | 10 MiB |
| `maxEntries` | 1,024 |
| `maxEntryBytes` | 8 MiB |
| `maxTotalBytes` | 32 MiB |
| `maxSheets` | 32 |
| `maxRows` | 10,000 |
| `maxColumns` | 1,024 |
| `maxCells` | 250,000 grid positions, including blanks |
| `maxMergedCells` | 250,000 positions |
| `maxDrawings` | 1,000 |
| `maxDrawingPixels` | 100,000 per coordinate, offset, extent, or rectangle bound |
| `maxXmlNodes` | 100,000 per XML part, including text nodes |
| `maxXmlDepth` | 64 |
| `maxHtmlLength` | 16,777,216 UTF-16 code units |

This strict reader supports classic, single-volume, unencrypted ZIP archives with UTF-8/ASCII paths. ZIP64 and Unicode path-override extra fields are rejected. Unsupported required sheet types (such as chart sheets), missing worksheet relationships, and inconsistent dimensions fail explicitly rather than returning partial data.

For a security-critical application, also isolate parsing in a terminable process or an execution context with a compatible XML parser, enforce application time/memory limits, and render previews in an iframe with scripts disabled and a restrictive network CSP. The library itself does not create that isolation. Browser image decoding and DOM/XML parsing still use resources outside these library budgets. Hidden sheets and rows are display metadata, never an access-control boundary. When displaying parser errors or filenames, use `textContent`.

### Development and regression tests

Run `npm ci` and `npm test`. The test command type-checks, builds, and runs adversarial ZIP/XML/HTML tests using Node's test runner and jsdom. Test dependencies require Node 22.22.2+, Node 24.15.0+, or Node 26+ (as supported by jsdom); these are development requirements, not browser runtime requirements.

For the demo, run `npm ci --prefix examples/vite`, `npx tsc --noEmit -p examples/vite/tsconfig.json`, and `npm --prefix examples/vite run build-storybook`.

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
