import { test, expect } from '@playwright/test';

test('device memory check page runs the production worker and records results', async ({ page }) => {
    await page.goto('/memory.html?cases=numbers-25k,worst-case');
    for (const id of ['numbers-25k', 'worst-case']) {
        await expect(page.locator(`tr[data-case="${id}"]`)).toHaveAttribute('data-status', 'ok', { timeout: 40_000 });
    }
    const saved = await page.evaluate(() => JSON.parse(localStorage.getItem('xlsx-to-js-memory-check')));
    expect(saved.running).toBeNull();
    expect(saved.results.map(r => r.id)).toEqual(['numbers-25k', 'worst-case']);
    // A marker left by a killed tab is reported as a crash after reload.
    await page.evaluate(() => localStorage.setItem('xlsx-to-js-memory-check', JSON.stringify({ results: [], running: 'numbers-100k' })));
    await page.goto('/memory.html');
    await expect(page.locator('tr[data-case="numbers-100k"]')).toHaveAttribute('data-status', 'crashed');
});
