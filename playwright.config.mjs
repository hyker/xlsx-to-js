import { defineConfig, devices } from '@playwright/test';
export default defineConfig({
    testDir: './tests/browser',
    timeout: 45_000,
    expect: { timeout: 15_000 },
    workers: 1,
    reporter: [['list'], ['html', { open: 'never' }]],
    use: { baseURL: 'http://127.0.0.1:4173', trace: 'retain-on-failure' },
    projects: [
        { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
        { name: 'firefox', use: { ...devices['Desktop Firefox'] } },
        { name: 'webkit', use: { ...devices['Desktop Safari'] } },
    ],
    webServer: {
        command: 'npm --prefix examples/vite run dev -- --host 127.0.0.1 --port 4173 --strictPort',
        url: 'http://127.0.0.1:4173', reuseExistingServer: !process.env.CI,
    },
});
