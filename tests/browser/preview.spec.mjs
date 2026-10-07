import { test, expect } from '@playwright/test';
import { NS, fixture, worksheetXml, stylesXml, drawingXml, gifWithComment, relationships } from '../fixtures.mjs';

let large;
const letters = n => { let result = ''; for (; n; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result; return result; };
test.beforeAll(async () => {
    const data = Array.from({ length: 1000 }, (_, r) => `<row r="${r + 1}">${Array.from({ length: 100 }, (_, c) => `<c r="${letters(c + 1)}${r + 1}" s="0"><v>45000</v></c>`).join('')}</row>`).join('');
    large = Buffer.from(await fixture({ sheet: worksheetXml(data, 'A1:CV1000'), style: stylesXml({ numFmtId: 14 }) }));
});
test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('[data-action="sample"]')).toBeVisible();
});
const upload = (page, name, buffer) => page.locator('[data-action="upload"]').setInputFiles({ name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer });
const loaded = page => expect(page.locator('[data-role="status"] strong')).toHaveText('Workbook loaded');

// Let real exports finish, but hold their results until the test releases them.
// This keeps cancellation deterministic even when CI clicks take longer than
// a timer, and lets us exercise a stale callback after worker termination.
const holdExportResults = page => page.evaluate(() => {
    window.exportJobs = 0;
    window.heldExports = [];
    const Native = window.Worker;
    window.Worker = class extends Native {
        constructor(...args) {
            super(...args);
            this.rendering = false;
            this.wasTerminated = false;
            this.addEventListener('message', event => {
                if (this.rendering && event.data.type === 'result') {
                    event.stopImmediatePropagation();
                    window.heldExports.push({ worker: this, event });
                }
            });
        }
        postMessage(message, ...args) {
            if (message.type === 'render') { this.rendering = true; window.exportJobs++; }
            super.postMessage(message, ...args);
        }
        terminate() { this.wasTerminated = true; super.terminate(); }
    };
    window.releaseExportResult = () => {
        const { worker, event } = window.heldExports.shift();
        worker.onmessage?.(event);
        return worker.wasTerminated;
    };
});

test('sample workbook renders styles and merges, then switches sheets', async ({ page }) => {
    await page.locator('[data-option="styles"]').check();
    await page.locator('[data-option="drawings"]').check();
    await page.locator('[data-action="sample"]').click();
    await loaded(page);
    await expect(page.locator('.xl-name')).toHaveText('Overview');
    await page.getByRole('button', { name: 'Dataset', exact: true }).click();
    await expect(page.locator('.xl-name')).toHaveText('Dataset');
    await expect(page.locator('td[rowspan],td[colspan]').first()).toBeVisible();
});

test('100,000 styled date cells keep a bounded DOM and navigate both axes', async ({ page, browserName }, testInfo) => {
    if (browserName === 'chromium') {
        const cdp = await page.context().newCDPSession(page);
        await cdp.send('Emulation.setCPUThrottlingRate', { rate: 4 });
    }
    await page.locator('[data-option="styles"]').check();
    await page.evaluate(() => { window.previewTicks = 0; window.previewTimer = setInterval(() => window.previewTicks++, 20); });
    await upload(page, 'large-styled-dates.xlsx', large);
    await loaded(page);
    await expect(page.locator('.xl-cell')).toHaveCount(5000);
    await expect(page.locator('[data-ref="A1"]')).toHaveText(/2023/);
    expect(await page.evaluate(() => window.previewTicks)).toBeGreaterThan(2);
    await expect.poll(() => page.locator('.sb-demo').getAttribute('data-first-display-ms')).not.toBeNull();
    const metrics = await page.evaluate(async () => {
        const start = performance.now();
        await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        const root = document.querySelector('.sb-demo');
        const memory = performance.memory;
        return { parseMs: Number(root.dataset.parseMs), loadMs: Number(root.dataset.loadMs), firstDisplayMs: Number(root.dataset.firstDisplayMs), ...JSON.parse(root.dataset.renderMetrics), displayWaitMs: performance.now() - start,
            domNodes: root.querySelectorAll('*').length, heapBytes: memory?.usedJSHeapSize ?? null };
    });
    console.log(`${browserName} preview metrics: ${JSON.stringify(metrics)}`);
    await testInfo.attach('preview-metrics.json', { body: JSON.stringify(metrics, null, 2), contentType: 'application/json' });
    expect(metrics.domNodes).toBeLessThan(5500);
    expect(metrics.htmlLength).toBeLessThan(500000);
    // Broad smoke threshold; device-specific timings are recorded as artifacts.
    expect(metrics.loadMs).toBeLessThan(25000);
    await page.locator('[data-action="row-next"]').click();
    await expect(page.locator('[data-ref="A101"]')).toHaveText(/2023/);
    await page.locator('[data-action="col-next"]').click();
    await expect(page.locator('[data-ref="AY101"]')).toHaveText(/2023/);
    await expect(page.locator('.xl-cell')).toHaveCount(5000);
    const scrolling = await page.evaluate(async () => {
        const canvas = document.querySelector('[data-role="canvas"]');
        const start = performance.now();
        canvas.scrollTop = canvas.scrollHeight;
        await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        return { ms: performance.now() - start, scrollTop: canvas.scrollTop };
    });
    expect(scrolling.scrollTop).toBeGreaterThan(0);
    await testInfo.attach('scroll-metrics.json', { body: JSON.stringify(scrolling), contentType: 'application/json' });
});

test('cancellation stops pending work and allows a subsequent load', async ({ page }) => {
    await page.evaluate(() => {
        const Native = window.Worker;
        window.Worker = class extends Native {
            postMessage(...args) { setTimeout(() => super.postMessage(...args), 500); }
        };
    });
    await upload(page, 'cancel.xlsx', large);
    await expect(page.locator('[data-action="cancel"]')).toBeEnabled();
    await page.locator('[data-action="cancel"]').click();
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Loading cancelled');
    await page.locator('[data-action="sample"]').click();
    await loaded(page);
    await expect(page.locator('.xl-name')).toHaveText('Overview');
});

test('newer workbook wins over an outstanding upload', async ({ page }) => {
    await upload(page, 'older.xlsx', large);
    await page.locator('[data-action="sample"]').click();
    await loaded(page);
    await expect(page.locator('.xl-name')).toHaveText('Overview');
    // Wait for any old worker result to prove it cannot replace the selected source.
    await page.waitForTimeout(500);
    await expect(page.locator('[data-role="status"] span')).toContainText('embedded sample');
    await expect(page.locator('.xl-name')).toHaveText('Overview');
});

test('oversized uploads are rejected before reading their bytes', async ({ page }) => {
    await page.evaluate(() => {
        window.uploadReadCalls = 0;
        File.prototype.arrayBuffer = async () => { window.uploadReadCalls++; throw new Error('Should not read oversized file'); };
    });
    await upload(page, 'oversized.xlsx', Buffer.alloc(10 * 1024 * 1024 + 1));
    await expect(page.locator('[data-role="status"] strong')).toHaveText('File too large');
    expect(await page.evaluate(() => window.uploadReadCalls)).toBe(0);
});

test('worker rejects malformed XML and the UI recovers', async ({ page }) => {
    const bad = Buffer.from(await fixture({ sheet: '<worksheet><sheetData></sheetData></worksheet\njunk>' }));
    await upload(page, 'malformed.xlsx', bad);
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Loading failed');
    await expect(page.locator('[data-role="status"] span')).toContainText('XML');
    await page.locator('[data-action="sample"]').click();
    await loaded(page);
});

test('worker rejects amplified workbook text before returning a result and the UI recovers', async ({ page }) => {
    const data = Array.from({ length: 1000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>0</v></c></row>`).join('');
    const bytes = Buffer.from(await fixture({
        sheet: worksheetXml(data, 'A1:A1000'),
        shared: `<sst xmlns="${NS}"><si><t>${'x'.repeat(32767)}</t></si></sst>`,
    }));
    expect(bytes.length).toBeLessThan(10000);
    await upload(page, 'amplified.xlsx', bytes);
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Loading failed');
    await expect(page.locator('[data-role="status"] span')).toContainText('Workbook text');
    await expect(page.locator('.xl-cell')).toHaveCount(0);
    await page.locator('[data-action="sample"]').click();
    await loaded(page);
});

test('large embedded GIF validates, renders and decodes within explicit output budgets', async ({ page }) => {
    const bytes = Buffer.from(await fixture({
        sheet: worksheetXml(undefined, 'A1', '<drawing r:id="drawingRel"/>'),
        drawing: drawingXml({ image: true }),
        drawingRelationships: relationships([{ id: 'imageRel', kind: 'image', target: '../media/image.gif' }]),
        extra: [['xl/media/image.gif', gifWithComment()]],
    }));
    const result = await page.evaluate(async ({ encoded, moduleUrl, workerUrl }) => {
        const { XlsxParser, XlsxWorkerParser } = await import(moduleUrl);
        const buffer = Uint8Array.from(atob(encoded), c => c.charCodeAt(0)).buffer;
        const workbook = await new XlsxWorkerParser(() => new Worker(workerUrl, { type: 'module' }))
            .readFile(buffer, { drawings: true });
        const parser = new XlsxParser();
        const preview = parser.toHTMLSheetPage(workbook, 0);
        const full = parser.toHTML(workbook);
        document.querySelector('[data-role="canvas"]').innerHTML = preview.html;
        let budgetError;
        try { parser.toHTMLSheetPage(workbook, 0, { limits: { maxHtmlLength: 2 * 1024 * 1024 } }); }
        catch (error) { budgetError = error.message; }
        return { previewLength: preview.html.length, fullLength: full.length, budgetError };
    }, {
        encoded: bytes.toString('base64'),
        moduleUrl: '/@fs' + process.cwd() + '/dist/index.js',
        workerUrl: '/@fs' + process.cwd() + '/dist/worker.js',
    });
    expect(result.previewLength).toBeGreaterThan(5 * 1024 * 1024);
    expect(result.fullLength).toBeLessThan(16 * 1024 * 1024);
    expect(result.budgetError).toContain('HTML exceeds output budget');
    const image = page.locator('.xl-abs img');
    await expect.poll(() => image.evaluate(img => img.complete && img.naturalWidth === 1 && img.naturalHeight === 1)).toBe(true);
});

test('hidden-column merges align visible cells with their column headers', async ({ page }) => {
    const bytes = Buffer.from(await fixture({ sheet: worksheetXml('<row r="1"><c r="A1"><v>42</v></c><c r="C1"><v>9</v></c></row>', 'A1:C1', '<cols><col min="1" max="1" hidden="1"/></cols><mergeCells><mergeCell ref="A1:B1"/></mergeCells>') }));
    await upload(page, 'merged.xlsx', bytes);
    await loaded(page);
    const merged = await page.locator('[data-ref="B1"]').boundingBox();
    const colB = await page.locator('[data-col="2"]').boundingBox();
    const colC = await page.locator('[data-col="3"]').boundingBox();
    const cellC = await page.locator('[data-ref="C1"]').boundingBox();
    expect(Math.abs(merged.x - colB.x)).toBeLessThan(1);
    expect(Math.abs(cellC.x - colC.x)).toBeLessThan(1);
    await expect(page.locator('[data-ref="B1"]')).toHaveText('42');
});


test('embedded drawings are clipped and repositioned across column pages', async ({ page }) => {
    const bytes = Buffer.from(await fixture({
        sheet: worksheetXml('<row r="1"><c r="A1"><v>42</v></c></row>', 'A1:BC2', '<drawing r:id="drawingRel"/>'),
        drawing: drawingXml({ image: true, col: '49', extent: '1905000' }),
        drawingRelationships: relationships([{ id: 'imageRel', kind: 'image', target: '../media/image.gif' }]),
        extra: [['xl/media/image.gif', Buffer.from('R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw==', 'base64')]],
    }));
    await page.locator('[data-option="drawings"]').check();
    await upload(page, 'drawings.xlsx', bytes);
    await loaded(page);
    await page.locator('[data-action="col-next"]').click();
    const img = page.locator('.xl-abs img');
    await expect(img).toBeVisible();
    await expect.poll(() => img.evaluate(image => image.complete && image.naturalWidth > 0)).toBe(true);
    expect(await img.evaluate(image => parseFloat(image.style.left))).toBeLessThan(36);
    await expect(page.locator('.xl-wrap')).toHaveCSS('overflow', 'hidden');
});

test('full exports retain header alignment for merges spanning hidden columns', async ({ page }) => {
    const html = await page.evaluate(async (url) => {
        const { XlsxParser } = await import(url);
        const sheet = { id: 1, name: 'Full merge', dimention: 'A1:D2', data: [[{ value: 'merged' }, , , { value: 'last' }]],
            defaultColWidth: 8.43, defaultRowHeight: 15, zeroHeight: false,
            columnStyles: [{ min: 2, max: 2, width: 10, hidden: true, collapsed: false }], rowStyles: [],
            mergeCells: ['A1:C1'], drawings: [] };
        return new XlsxParser().toHTML({ workSheets: [sheet] });
    }, '/@fs' + process.cwd() + '/dist/index.js');
    await page.locator('[data-role="canvas"]').evaluate((canvas, html) => canvas.innerHTML = html, html);
    const header = await page.locator('[data-col="4"]').boundingBox();
    const last = await page.locator('[data-ref="D1"]').boundingBox();
    expect(Math.abs(header.x - last.x)).toBeLessThan(1);
});


test('full exports keep visible row spans aligned across hidden rows', async ({ page }) => {
    const html = await page.evaluate(async (url) => {
        const { XlsxParser } = await import(url);
        const sheet = { id: 1, name: 'Full row merge', dimention: 'A1:B4', data: [[{ value: 'merged' }], [], [], [{ value: 'last' }, { value: 'second' }]],
            defaultColWidth: 8.43, defaultRowHeight: 15, zeroHeight: false,
            columnStyles: [], rowStyles: [{ r: 2, height: 15, hidden: true, collapsed: false }],
            mergeCells: ['A1:A3'], drawings: [] };
        return new XlsxParser().toHTML({ workSheets: [sheet] });
    }, '/@fs' + process.cwd() + '/dist/index.js');
    await page.locator('[data-role="canvas"]').evaluate((canvas, html) => canvas.innerHTML = html, html);
    const header = await page.locator('[data-col="1"]').boundingBox();
    const last = await page.locator('[data-ref="A4"]').boundingBox();
    expect(Math.abs(header.x - last.x)).toBeLessThan(1);
});

test('oversized intrinsic image pixels fail before display and the UI recovers', async ({ page }) => {
    const { pngImage } = await import('../fixtures.mjs');
    await page.locator('[data-option="drawings"]').check();
    const bytes = Buffer.from(await fixture({ sheet: worksheetXml(undefined, 'A1', '<drawing r:id="drawingRel"/>'),
        drawing: drawingXml({ image: true }), drawingRelationships: relationships([{ id: 'imageRel', kind: 'image', target: '../media/image.png' }]),
        extra: [['xl/media/image.png', pngImage(4096, 4096)]],
    }));
    await upload(page, 'pixel-bomb.xlsx', bytes);
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Loading failed');
    await expect(page.locator('[data-role="status"] span')).toContainText('Image pixels');
    await expect(page.locator('.xl-abs img')).toHaveCount(0);
    await page.locator('[data-action="sample"]').click(); await loaded(page);
});

test('static PNG and JPEG headers allow images that decode in real browsers', async ({ page }) => {
    const { pngImage } = await import('../fixtures.mjs');
    const jpeg = Buffer.from(await page.evaluate(() => {
        const canvas = document.createElement('canvas'); canvas.width = 32; canvas.height = 16;
        const context = canvas.getContext('2d'); context.fillStyle = '#336699'; context.fillRect(0, 0, 32, 16);
        return canvas.toDataURL('image/jpeg').split(',')[1];
    }), 'base64');
    await page.locator('[data-option="drawings"]').check();
    for (const [name, image, dimensions] of [['PNG', pngImage(32, 16), [32, 16]], ['JPEG', jpeg, [32, 16]]]) {
        const bytes = Buffer.from(await fixture({ sheet: worksheetXml(undefined, 'A1', '<drawing r:id="drawingRel"/>'),
            drawing: drawingXml({ image: true }), drawingRelationships: relationships([{ id: 'imageRel', kind: 'image', target: '../media/image.bin' }]),
            extra: [['xl/media/image.bin', image]],
        }));
        await upload(page, `${name}.xlsx`, bytes); await loaded(page);
        const size = await page.locator('.xl-abs img').evaluate(async img => { await img.decode(); return [img.naturalWidth, img.naturalHeight]; });
        expect(size).toEqual(dimensions);
    }
});

test('full exports run in a worker and cancellation leaves preview usable', async ({ page }) => {
    await page.locator('xlsx-parser-demo').evaluate(element => { element.mode = 'all'; });
    await page.locator('[data-action="sample"]').click(); await loaded(page);
    await holdExportResults(page);
    await page.locator('[data-action="export"]').click();
    await expect.poll(() => page.evaluate(() => window.heldExports.length)).toBe(1);
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Exporting workbook');
    await page.locator('[data-action="cancel"]').click();
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Export cancelled');
    expect(await page.evaluate(() => window.heldExports[0].worker.wasTerminated)).toBe(true);
    await page.evaluate(() => window.releaseExportResult());
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Export cancelled');
    await expect(page.locator('.xl-cell').first()).toBeVisible();
    const downloadEvent = page.waitForEvent('download');
    await page.locator('[data-action="export"]').click();
    await expect.poll(() => page.evaluate(() => window.heldExports.length)).toBe(1);
    expect(await page.evaluate(() => window.heldExports[0].worker.wasTerminated)).toBe(false);
    await page.evaluate(() => window.releaseExportResult());
    const download = await downloadEvent;
    expect(download.suggestedFilename()).toBe('workbook.html');
    const stream = await download.createReadStream(), chunks = [];
    for await (const chunk of stream) chunks.push(chunk);
    const html = Buffer.concat(chunks).toString();
    expect(html).toContain('Overview'); expect(html).toContain('Dataset');
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Export ready');
    expect(await page.evaluate(() => window.exportJobs)).toBe(2);
});

test('a new load cancels an outstanding export without a stale download', async ({ page }) => {
    await page.locator('xlsx-parser-demo').evaluate(element => { element.mode = 'all'; });
    await page.locator('[data-action="sample"]').click(); await loaded(page);
    const downloads = []; page.on('download', download => downloads.push(download));
    await holdExportResults(page);
    await page.locator('[data-action="export"]').click();
    await expect.poll(() => page.evaluate(() => window.heldExports.length)).toBe(1);
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Exporting workbook');
    await page.locator('[data-action="sample"]').click(); await loaded(page);
    expect(await page.evaluate(() => window.heldExports[0].worker.wasTerminated)).toBe(true);
    await page.evaluate(() => window.releaseExportResult());
    await expect(page.locator('[data-action="export"]')).toBeEnabled();
    await expect(page.locator('[data-role="status"] strong')).toHaveText('Workbook loaded');
    expect(downloads).toHaveLength(0);
});
