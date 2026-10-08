import test from 'node:test';
import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { XlsxParser, XlsxWorkerParser, XlsxLimitError, DEFAULT_LIMITS, PREVIEW_LIMITS } from '../dist/index.js';
import { fixture, worksheetXml, stylesXml, drawingXml } from './fixtures.mjs';
const dom = new JSDOM('<main></main>');
if (process.env.XLSX_XML_PARSER === 'xmldom') await import('../dist/worker.js');
else globalThis.DOMParser = dom.window.DOMParser;
const parser = new XlsxParser();
const read = async (sheet, options = {}, extra = {}) => parser.readFile(await fixture({ sheet, ...extra }), options);
const inspect = html => { const main = dom.window.document.querySelector('main'); main.innerHTML = html; return main; };

test('sparse 250,000-position preview is bounded and preserves last-page coordinates', async () => {
    const limits = { maxCells: 250_000 };
    const w = await read(worksheetXml('<row r="10000"><c r="Y10000"><v>42</v></c></row>', 'A1:Y10000'), { limits });
    const p = parser.toHTMLSheetPage(w, 0, { rowPage: 99, limits });
    assert.equal(p.totalRows, 10000); assert.equal(p.totalColumns, 25);
    assert.equal(p.totalRowPages, 100); assert.equal(p.rowStart, 9901);
    const root = inspect(p.html);
    assert.equal(root.querySelectorAll('td').length, 2500);
    assert.equal(root.querySelector('[data-ref="Y10000"]').textContent, '42');
    assert.ok(p.html.length < 200000);
});

test('two-dimensional pages clip merged regions and retain their original anchor value', async () => {
    const w = await read(worksheetXml('<row r="1"><c r="A1"><v>42</v></c></row>', 'A1:D4', '<mergeCells><mergeCell ref="A1:C3"/></mergeCells>'));
    const p = parser.toHTMLSheetPage(w, 0, { pageRows: 2, pageColumns: 2, rowPage: 1, columnPage: 1 });
    const root = inspect(p.html), td = root.querySelector('[data-ref="C3"]');
    assert.equal(td.textContent, '42'); assert.equal(td.rowSpan, 1); assert.equal(td.colSpan, 1);
    assert.equal(td.dataset.mergeRef, 'A1');
    assert.equal(root.querySelectorAll('td').length, 4);
});

test('hidden merge anchors retain visible cell positions without revealing skipped values', async () => {
    const sheet = worksheetXml('<row r="1" hidden="1"><c r="A1"><v>42</v></c></row><row r="2"><c r="C2"><v>9</v></c></row>', 'A1:C2', '<mergeCells><mergeCell ref="A1:B2"/></mergeCells>');
    const w = await read(sheet, { skipHiddenRows: true });
    for (const html of [parser.toHTML(w), parser.toHTMLSheetPage(w, 0).html]) {
        const root = inspect(html), merged = root.querySelector('[data-ref="A2"]');
        assert.equal(merged.colSpan, 2); assert.equal(merged.rowSpan, 1); assert.equal(merged.textContent, '');
        assert.equal(root.querySelector('[data-ref="C2"]').textContent, '9');
        assert.equal(root.querySelector('[data-row="1"]'), null);
    }
});

test('merges anchored in a hidden column render their visible portion', async () => {
    const w = await read(worksheetXml('<row r="1"><c r="A1"><v>42</v></c><c r="C1"><v>9</v></c></row>', 'A1:C1', '<cols><col min="1" max="1" hidden="1"/></cols><mergeCells><mergeCell ref="A1:B1"/></mergeCells>'));
    for (const html of [parser.toHTML(w), parser.toHTMLSheetPage(w, 0).html]) {
        const root = inspect(html), td = root.querySelector('[data-ref="B1"]');
        assert.equal(td.textContent, '42'); assert.equal(td.style.display, ''); assert.equal(td.colSpan, 1);
        assert.equal(root.querySelector('[data-ref="C1"]').textContent, '9');
    }
});

test('collapsed outline summaries remain visible', async () => {
    const w = await read(worksheetXml('<row r="1" collapsed="1"><c r="A1"><v>42</v></c></row>', 'A1', '<cols><col min="1" max="1" collapsed="1" width="10"/></cols>'), { skipHiddenRows: true });
    const root = inspect(parser.toHTML(w));
    assert.equal(root.querySelector('tbody tr').style.display, '');
    assert.equal(root.querySelector('[data-col="1"]').style.display, '');
    assert.equal(root.querySelector('td').textContent, '42');
});

test('omitted dimensions and references use the same coordinate inference', async () => {
    const w = await read(worksheetXml('<row><c><v>42</v></c><c><v>43</v></c></row><row><c><v>44</v></c></row>', null));
    assert.equal(w.workSheets[0].dimention, 'A1:B2');
    assert.equal(w.workSheets[0].data[1][0].ref, 'A2');
    assert.equal(w.workSheets[0].data[0][1].value, '43');
});

test('page options are bounded and never bypass full-sheet resource budgets', async () => {
    const w = await read(worksheetXml('', 'A1:B2'));
    for (const options of [{ pageRows: 0 }, { pageColumns: -1 }, { rowPage: 0.5 }, { rowPage: 99 }, { pageRows: 101, pageColumns: 50 }, { limits: { maxCells: 1 } }]) {
        assert.throws(() => parser.toHTMLSheetPage(w, 0, options));
    }
    assert.throws(() => parser.toHTMLSheetPage(w, -1), /sheet index/);
    w.workSheets[0].state = 'hidden';
    assert.equal(parser.toHTMLSheetPage(w, 0).totalRows, 0);
    assert.equal(parser.toHTMLSheetPage(w, 0, { includeHiddenSheets: true }).totalRows, 2);
});

test('shared style classes stay isolated between fragments and cannot inject HTML', async () => {
    const w = await read(worksheetXml('<row r="1"><c r="A1" s="0"><v>1</v></c><c r="B1" s="0"><v>2</v></c></row>', 'A1:B1'), { styles: true }, { style: stylesXml({ font: '</style><img src=x onerror=alert(1)>' }) });
    assert.equal(w.workSheets[0].data[0][0].style, w.workSheets[0].data[0][1].style);
    const html = parser.toHTMLSheetPage(w, 0).html;
    const root = inspect(html);
    assert.equal(root.querySelectorAll('img').length, 0);
    assert.equal(root.querySelectorAll('[onerror]').length, 0);
    assert.equal(root.querySelectorAll('style').length, 2);
    assert.equal(root.querySelectorAll('td')[0].className, root.querySelectorAll('td')[1].className);
    const red = structuredClone(w), blue = structuredClone(w);
    red.workSheets[0].data[0][0].style.fontColor = '#ff0000';
    blue.workSheets[0].data[0][0].style.fontColor = '#0000ff';
    const fragments = inspect(parser.toHTML(red) + parser.toHTML(blue));
    const cells = fragments.querySelectorAll('td');
    assert.equal(dom.window.getComputedStyle(cells[0]).color, 'rgb(255, 0, 0)');
    assert.equal(dom.window.getComputedStyle(cells[2]).color, 'rgb(0, 0, 255)');
});

test('drawing-enabled read parses each worksheet only once', async () => {
    const Native = globalThis.DOMParser; let count = 0;
    globalThis.DOMParser = class { parseFromString(xml, type) { if (xml.includes('<worksheet')) count++; return new Native().parseFromString(xml, type); } };
    try { await read(worksheetXml(undefined, 'A1', '<drawing r:id="drawingRel"/>'), { drawings: true }, { drawing: drawingXml() }); assert.equal(count, 1); }
    finally { globalThis.DOMParser = Native; }
});

test('date formatting preserves UTC values and works with styles disabled', async () => {
    const sheet = worksheetXml('<row r="1"><c r="A1" s="0"><v>45000</v></c></row>');
    const w = await read(sheet, { styles: false }, { style: stylesXml({ numFmtId: 14 }) });
    assert.equal(w.workSheets[0].data[0][0].value, new Date(Date.UTC(2023, 2, 15)).toLocaleDateString(undefined, { timeZone: 'UTC' }));
});

test('parser progress and cooperative cancellation report a complete lifecycle', async () => {
    const file = await fixture(); const progress = [];
    await parser.readFile(file, { onProgress: p => progress.push(p) });
    assert.deepEqual(progress.map(p => p.phase), ['archive', 'metadata', 'worksheet', 'complete']);
    assert.equal(progress.at(-1).completedSheets, 1);
    const controller = new AbortController(); controller.abort();
    await assert.rejects(parser.readFile(file, { signal: controller.signal }), { name: 'AbortError' });
});

test('worker client terminates on abort, timeout, failures and callback exceptions', async () => {
    for (const mode of ['abort', 'timeout', 'error', 'callback']) {
        const worker = { terminated: false, postMessage() {}, terminate() { this.terminated = true; } };
        const controller = new AbortController();
        const result = new XlsxWorkerParser(() => worker).readFile(new ArrayBuffer(1), { signal: controller.signal, timeoutMs: mode === 'timeout' ? 5 : 1000, onProgress: () => { throw new Error('callback failed'); } });
        if (mode === 'abort') controller.abort();
        if (mode === 'error') worker.onerror({ message: 'worker failed' });
        if (mode === 'callback') worker.onmessage({ data: { type: 'progress', progress: {} } });
        await assert.rejects(result); assert.equal(worker.terminated, true);
    }
});

test('worker preflight rejects oversized bytes before creating a worker', async () => {
    let created = false;
    const client = new XlsxWorkerParser(() => { created = true; });
    await assert.rejects(client.readFile(new ArrayBuffer(2), { limits: { maxFileBytes: 1 } }), /file budget/);
    assert.equal(created, false);
    assert.ok(Object.isFrozen(PREVIEW_LIMITS));
});

test('worker client caps concurrent parse and export jobs and releases slots after cancellation', async () => {
    const workers = [];
    const client = new XlsxWorkerParser(() => {
        const worker = { terminated: false, postMessage(message) { this.message = message; }, terminate() { this.terminated = true; } };
        workers.push(worker); return worker;
    });
    const controller = new AbortController();
    const parsing = client.readFile(new ArrayBuffer(1), { signal: controller.signal });
    await assert.rejects(client.readFile(new ArrayBuffer(1)), /concurrency/);
    await assert.rejects(client.toHTML({ workSheets: [] }), /concurrency/);
    assert.equal(workers.length, 1);
    controller.abort(); await assert.rejects(parsing, { name: 'AbortError' });
    const exporting = client.toHTML({ workSheets: [] });
    assert.equal(workers[1].message.type, 'render');
    workers[1].onmessage({ data: { type: 'result', html: '<div>exported</div>' } });
    assert.equal(await exporting, '<div>exported</div>');
    assert.ok(workers.every(w => w.terminated));
    for (const maxConcurrentJobs of [0, -1, NaN, 0.5]) assert.throws(() => new XlsxWorkerParser(() => {}, { maxConcurrentJobs }), /concurrency/);
});

test('worker exports terminate on timeout, abort and clone errors, and reject oversized text before cloning', async () => {
    const workbook = await read(worksheetXml('<row r="1"><c r="A1" t="inlineStr"><is><t>large text</t></is></c></row>'));
    for (const mode of ['abort', 'timeout', 'clone']) {
        const worker = { terminated: false, postMessage() { if (mode === 'clone') throw new Error('clone failed'); }, terminate() { this.terminated = true; } };
        const controller = new AbortController();
        const result = new XlsxWorkerParser(() => worker).toHTML(workbook, { signal: controller.signal, timeoutMs: mode === 'timeout' ? 5 : 1000 });
        if (mode === 'abort') controller.abort();
        await assert.rejects(result); assert.equal(worker.terminated, true);
    }
    let created = false;
    await assert.rejects(new XlsxWorkerParser(() => { created = true; }).toHTML(workbook, { limits: { maxWorkbookTextLength: 1 } }), /Workbook text/);
    assert.equal(created, false);
});

test('worker XML floods reject within a 128 MiB heap and a five-second deadline', () => {
    for (const mode of ['nodes', 'attributes']) {
        const child = spawnSync(process.execPath, ['--max-old-space-size=128', fileURLToPath(new URL('./worker-resource-child.mjs', import.meta.url)), mode], { timeout: 5000, encoding: 'utf8' });
        assert.equal(child.error, undefined, child.error?.message);
        assert.equal(child.status, 0, child.stderr);
    }
});

test('100,000 styled cells, result cloning and preview stay within a measured capacity ceiling', () => {
    const child = spawnSync(process.execPath, ['--max-old-space-size=384', fileURLToPath(new URL('./worker-resource-child.mjs', import.meta.url)), 'capacity'], { timeout: 20000, encoding: 'utf8' });
    assert.equal(child.error, undefined, child.error?.message);
    assert.equal(child.status, 0, child.stderr);
    const metrics = JSON.parse(child.stdout.trim());
    // Includes ZIP fixture generation, the worker's XML implementation and clone.
    // This smoke ceiling is not a guarantee of browser or low-memory device use.
    assert.ok(metrics.peakRssBytes < 768 * 1024 * 1024, JSON.stringify(metrics));
    console.log(`worker capacity metrics: ${JSON.stringify(metrics)}`);
});

test('worker timeout values cannot wrap the platform timer and abort during creation is respected', async () => {
    let created = false;
    const client = new XlsxWorkerParser(() => { created = true; });
    for (const timeoutMs of [0, -1, Infinity, NaN, 0.5, 2147483648]) {
        await assert.rejects(client.readFile(new ArrayBuffer(1), { timeoutMs }), /Invalid parsing timeout/);
        await assert.rejects(client.toHTML({ workSheets: [] }, { timeoutMs }), /Invalid parsing timeout/);
    }
    assert.equal(created, false);
    const controller = new AbortController();
    const worker = { terminated: false, postMessage() { assert.fail('Aborted job must not be posted'); }, terminate() { this.terminated = true; } };
    const result = new XlsxWorkerParser(() => { controller.abort(); return worker; }).readFile(new ArrayBuffer(1), { signal: controller.signal });
    await assert.rejects(result, { name: 'AbortError' });
    assert.equal(worker.terminated, true);
});

const columnLetters = n => { let result = ''; for (; n; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result; return result; };
const gridXml = (rows, cols, cell = (r, c) => `<v>${r * c}</v>`) => worksheetXml(Array.from({ length: rows }, (_, r) =>
    `<row r="${r + 1}">${Array.from({ length: cols }, (_, c) => `<c r="${columnLetters(c + 1)}${r + 1}" s="0">${cell(r + 1, c + 1)}</c>`).join('')}</row>`).join(''),
    `A1:${columnLetters(cols)}${rows}`);

test('numeric attribute validation stays linear for long digit runs', async () => {
    for (const ht of ['1'.repeat(200_000) + 'x', '1'.repeat(65)]) {
        const start = performance.now();
        await assert.rejects(read(worksheetXml(`<row r="1" ht="${ht}"><c r="A1"><v>1</v></c></row>`)), /Invalid row height/);
        assert.ok(performance.now() - start < 1000, `validation took ${performance.now() - start} ms`);
    }
    for (const ht of ['15', '15.75', '.5', '1e2', '20.']) {
        const w = await read(worksheetXml(`<row r="1" ht="${ht}"><c r="A1"><v>1</v></c></row>`));
        assert.equal(w.workSheets[0].rowStyles[0].height, Number(ht));
    }
});

test('default limits accept ordinary sheets up to the cell budget', async () => {
    // Tall and wide shapes that were rejected by the former node and row caps.
    for (const [rows, cols] of [[5000, 8], [20000, 5], [10000, 10]]) {
        const w = await read(gridXml(rows, cols), { limits: DEFAULT_LIMITS });
        assert.equal(w.workSheets[0].data.length, rows);
    }
    await assert.rejects(read(gridXml(10001, 10), { limits: DEFAULT_LIMITS }),
        e => e instanceof XlsxLimitError && e.limit === 'maxCells' && e.max === DEFAULT_LIMITS.maxCells);
    assert.deepEqual({ ...PREVIEW_LIMITS, maxHtmlLength: DEFAULT_LIMITS.maxHtmlLength }, { ...DEFAULT_LIMITS });
});

test('limit errors name the exceeded budget and its configured value', async () => {
    const cases = [
        [{ maxXmlNodes: 3 }, 'maxXmlNodes', /workbook part has too many XML nodes/],
        [{ maxXmlDepth: 2 }, 'maxXmlDepth', /nested too deeply/],
        [{ maxXmlAttributesPerElement: 1 }, 'maxXmlAttributesPerElement', /element/],
        [{ maxRows: 1 }, 'maxRows', /B2 is beyond the row limit/],
        [{ maxColumns: 1 }, 'maxColumns', /B2 is beyond the column limit/],
        [{ maxCells: 3 }, 'maxCells', /grid/],
        [{ maxFileBytes: 10 }, 'maxFileBytes', /file budget/],
    ];
    for (const [limits, limit, detail] of cases) {
        await assert.rejects(read(worksheetXml('<row r="2"><c r="B2" s="0"><v>1</v></c></row>', 'A1:B2'), { limits }), error => {
            assert.ok(error instanceof XlsxLimitError, `${limit}: ${error}`);
            assert.equal(error.name, 'XlsxLimitError');
            assert.equal(error.limit, limit); assert.equal(error.max, limits[limit]);
            assert.match(error.message, detail); assert.match(error.message, new RegExp(`limit ${limit} = ${limits[limit]}`));
            return true;
        });
    }
    const w = await read(gridXml(2, 2));
    assert.throws(() => parser.toHTML(w, { limits: { maxHtmlLength: 100 } }), { name: 'XlsxLimitError', limit: 'maxHtmlLength' });
});

test('limit errors survive the worker boundary as XlsxLimitError', async () => {
    const child = spawnSync(process.execPath, [fileURLToPath(new URL('./worker-resource-child.mjs', import.meta.url)), 'limit-error'], { timeout: 5000, encoding: 'utf8' });
    assert.equal(child.status, 0, child.stderr);
    const posted = JSON.parse(child.stdout.trim());
    assert.equal(posted.type, 'error'); assert.equal(posted.limit.limit, 'maxXmlNodes'); assert.equal(posted.limit.max, 3);
    const worker = { postMessage() {}, terminate() {} };
    const result = new XlsxWorkerParser(() => worker).readFile(new ArrayBuffer(1));
    worker.onmessage({ data: posted });
    await assert.rejects(result, error => error instanceof XlsxLimitError && error.limit === 'maxXmlNodes' && error.max === 3 && error.message === posted.error);
});

test('a full default-budget grid parses within a 352 MiB heap', () => {
    // Measured minimum is ~300 MiB for the parse itself; the margin covers fixture generation.
    const child = spawnSync(process.execPath, ['--max-old-space-size=352', fileURLToPath(new URL('./worker-resource-child.mjs', import.meta.url)), 'phone-budget'], { timeout: 30000, encoding: 'utf8' });
    assert.equal(child.error, undefined, child.error?.message);
    assert.equal(child.status, 0, child.stderr);
});
