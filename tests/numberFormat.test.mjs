import test from 'node:test';
import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';
import { XlsxParser } from '../dist/index.js';
import { NS, esc, fixture, worksheetXml } from './fixtures.mjs';
if (process.env.XLSX_XML_PARSER === 'xmldom') await import('../dist/worker.js');
else globalThis.DOMParser = new JSDOM('').window.DOMParser;
const parser = new XlsxParser();

const styles = (numFmtId, code) => `<styleSheet xmlns="${NS}">${code === undefined ? '' : `<numFmts count="1"><numFmt numFmtId="${numFmtId}" formatCode="${esc(code)}"/></numFmts>`}<fonts><font><name val="Arial"/><sz val="11"/></font></fonts><fills><fill><patternFill patternType="none"/></fill></fills><borders><border/></borders><cellXfs><xf numFmtId="${numFmtId}" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`;
async function cell(value, { code, id = code === undefined ? 0 : 164, locale = 'en-US', type = '' } = {}) {
    const v = type === 'inlineStr' ? `<is><t>${value}</t></is>` : `<v>${value}</v>`;
    const w = await parser.readFile(await fixture({
        sheet: worksheetXml(`<row r="1"><c r="A1" s="0"${type ? ` t="${type}"` : ''}>${v}</c></row>`),
        style: styles(id, code),
    }), { styles: true, locale });
    return w.workSheets[0].data[0][0];
}

test('custom number formats', async () => {
    const cases = [
        ['0.00', 3.14159, '3.14'], ['0.00', 1.005, '1.01'], ['#,##0', 1234567, '1,234,567'],
        ['#,##0.00', -1234.5, '-1,234.50'], ['0%', 0.256, '26%'], ['0.0%', 0.1234, '12.3%'],
        ['0.00E+00', 12345, '1.23E+04'], ['0.00E+00', 0.000123, '1.23E-04'], ['##0.0E+0', 12345, '12.3E+3'],
        ['#,##0;(#,##0)', -1234, '(1,234)'], ['#,##0.00_);[Red](#,##0.00)', -5, '(5.00)'], ['#,##0.00_);[Red](#,##0.00)', 5, '5.00 '],
        ['0;-0;"zero"', 0, 'zero'], ['"$"#,##0.00', 1234.5, '$1,234.50'], ['[$€-407] #,##0.00', 1234.5, '€ 1,234.50'],
        ['#,##0 "kr"', 1234, '1,234 kr'], ['#,##0,"k"', 1234567, '1,235k'], ['00000', 123, '00123'],
        ['000-00-0000', 123456789, '123-45-6789'], ['#.##', 0.5, '.5'], ['0.00', -0.001, '0.00'],
        ['_-* #,##0 "kr"_-;-* #,##0 "kr"_-;_-* "-" "kr"_-', -1234, '-1,234 kr '],
        ['"Total: "General', 0.1 + 0.2, 'Total: 0.3'], [';;;', 5, ''],
    ];
    for (const [code, value, expected] of cases) assert.equal((await cell(value, { code })).value, expected, code);
});

test('custom date and time formats', async () => {
    const cases = [
        ['yyyy-mm-dd', 45000, '2023-03-15'], ['dd.mm.yyyy', 45000, '15.03.2023'], ['d mmm yyyy', 45000, '15 Mar 2023'],
        ['mmmm d, yyyy', 45000, 'March 15, 2023'], ['dddd', 45000, 'Wednesday'], ['mmmmm-yy', 45000, 'M-23'],
        ['h:mm AM/PM', 45000.75, '6:00 PM'], ['h:mm a/p', 45000.25, '6:00 a'], ['hh:mm:ss', 0.5 + 5 / 86400, '12:00:05'],
        ['[h]:mm:ss', 1.5, '36:00:00'], ['[mm]:ss', 1 / 24, '60:00'], ['mm:ss.0', 61.4 / 86400, '01:01.4'],
        ['yyyy-mm-dd hh:mm', 45000.9999999, '2023-03-16 00:00'], ['[$-409]m/d/yy h:mm', 45000.5, '3/15/23 12:00'],
    ];
    for (const [code, value, expected] of cases) assert.equal((await cell(value, { code })).value, expected, code);
});

test('built-in formats without numFmts declarations', async () => {
    const cases = [[2, 3.14159, '3.14'], [3, 1234567, '1,234,567'], [10, 0.12345, '12.35%'], [11, 12345, '1.23E+04'],
        [15, 45000, '15-Mar-23'], [17, 45000, 'Mar-23'], [18, 0.75, '6:00 PM'], [20, 0.25, '6:00'], [21, 0.25, '6:00:00'],
        [22, 45000.5, '3/15/2023 12:00'], [38, -1234, '(1,234)'], [46, 1.5, '36:00:00'], [14, 45000, '3/15/2023']];
    for (const [id, value, expected] of cases) assert.equal((await cell(value, { id })).value, expected, `numFmtId ${id}`);
});

test('formatting keeps the stored value, aligns by cell type and falls back when unsupported', async () => {
    const formatted = await cell(3.14159, { code: '0.00' });
    assert.equal(formatted.value, '3.14'); assert.equal(formatted.raw, '3.14159'); assert.equal(formatted.style.hAlign, 'right');
    const plain = await cell(42, { id: 0 });
    assert.equal(plain.value, '42'); assert.equal('raw' in plain, false); assert.equal(plain.style.hAlign, 'right');
    // Numeric-looking text is text in Excel and stays left aligned.
    assert.equal((await cell('00123', { type: 'inlineStr' })).style.hAlign, 'left');
    for (const code of ['[>100]0;0', '# ?/?', '@', 'x'.repeat(256)]) {
        const fallback = await cell(0.5, { code });
        assert.equal(fallback.value, '0.5', code); assert.equal('raw' in fallback, false);
    }
    // Locale-dependent currency built-ins and out-of-range dates show the stored value.
    assert.equal((await cell(5, { id: 7 })).value, '5');
    assert.equal((await cell(-1, { code: 'yyyy-mm-dd' })).value, '-1');
    assert.equal((await cell(3e6, { code: 'yyyy-mm-dd' })).value, '3000000');
});

test('locale controls separators and names but not code order', async () => {
    assert.equal((await cell(1234.5, { code: '#,##0.00', locale: 'sv-SE' })).value, '1 234,50');
    assert.equal((await cell(45000, { code: 'd mmmm yyyy', locale: 'sv-SE' })).value, '15 mars 2023');
    assert.equal((await cell(45000, { code: 'dd.mm.yyyy', locale: 'sv-SE' })).value, '15.03.2023');
    assert.equal((await cell(45000, { id: 14, locale: 'sv-SE' })).value, '2023-03-15');
    await assert.rejects(cell(1, { code: '0', locale: 'not a locale!' }), RangeError);
});

test('formatted text is charged against the workbook text budget while parsing', async () => {
    const code = `0"${'x'.repeat(200)}"`;
    const rows = Array.from({ length: 50 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}" s="0"><v>1</v></c></row>`).join('');
    const file = await fixture({ sheet: worksheetXml(rows, 'A1:A50'), style: styles(164, code) });
    await assert.rejects(parser.readFile(file, { limits: { maxWorkbookTextLength: 5000 } }),
        e => e.name === 'XlsxLimitError' && e.limit === 'maxWorkbookTextLength');
});
