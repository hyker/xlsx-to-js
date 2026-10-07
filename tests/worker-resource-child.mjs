// Killable checks use the exact XML implementation bundled in browser workers.
import assert from 'node:assert/strict';
import { fixture, worksheetXml, stylesXml } from './fixtures.mjs';
import { XlsxParser, PREVIEW_LIMITS } from '../dist/index.js';
await import('../dist/worker.js');
const parser = new XlsxParser(), mode = process.argv[2];
if (mode === 'nodes') {
    const bytes = await fixture({ sheet: worksheetXml('<a/>'.repeat(1_000_001)), style: null });
    await assert.rejects(parser.readFile(bytes, { limits: PREVIEW_LIMITS }), /complexity/);
} else if (mode === 'attributes') {
    const bytes = await fixture({ sheet: worksheetXml(`<a ${Array.from({ length: 30000 }, (_, i) => `p${i}="x"`).join(' ')}/>`), style: null });
    await assert.rejects(parser.readFile(bytes), /attribute complexity/);
} else if (mode === 'capacity') {
    const letters = n => { let result = ''; for (; n; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result; return result; };
    const data = Array.from({ length: 1000 }, (_, r) => `<row r="${r + 1}">${Array.from({ length: 100 }, (_, c) => `<c r="${letters(c + 1)}${r + 1}" s="0"><v>45000</v></c>`).join('')}</row>`).join('');
    const bytes = await fixture({ sheet: worksheetXml(data, 'A1:CV1000'), style: stylesXml({ numFmtId: 14 }) });
    const start = performance.now();
    const workbook = await parser.readFile(bytes, { styles: true, limits: PREVIEW_LIMITS });
    // Exercise the same result clone performed when posting a worker response.
    const cloned = structuredClone(workbook);
    assert.equal(cloned.workSheets[0].data[999][99].value.includes('2023'), true);
    assert.equal(cloned.workSheets[0].data[0][0].style, cloned.workSheets[0].data[999][99].style);
    const page = parser.toHTMLSheetPage(cloned, 0, { limits: PREVIEW_LIMITS });
    assert.ok(page.html.length < 500000);
    console.log(JSON.stringify({ parseClonePreviewMs: performance.now() - start, peakRssBytes: process.resourceUsage().maxRSS * 1024 }));
} else throw new Error('Unknown resource test mode');
