import { test, expect } from '@playwright/test';
import { fixture, worksheetXml, stylesXml, drawingXml, relationships } from '../fixtures.mjs';

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
