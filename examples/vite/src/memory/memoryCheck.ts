import { DEFAULT_LIMITS, XlsxParser, XlsxWorkerParser, type XlsxLimits } from '../../../../dist/index.js';
import type { GenerateRequest } from './generator.worker';

interface Case extends GenerateRequest {
    id: string;
    group: 'default' | 'headroom';
    description: string;
    limits?: Partial<XlsxLimits>;
}
interface Result { id: string; status: string; generateMs?: number; parseMs?: number; renderMs?: number; zipKB?: number }
interface Saved { results: Result[]; running: string | null }

/** Each case scales the defaults; V8 heap figures come from Node with the same worker parser. */
const raised = (factor: number): Partial<XlsxLimits> => ({
    maxCells: DEFAULT_LIMITS.maxCells * factor, maxMergedCells: DEFAULT_LIMITS.maxMergedCells * factor,
    maxXmlNodes: DEFAULT_LIMITS.maxXmlNodes * factor, maxXmlAttributes: DEFAULT_LIMITS.maxXmlAttributes * factor,
});
const CASES: Case[] = [
    { id: 'numbers-25k', group: 'default', shape: 'numbers', cells: 25_000, description: '25k number cells (~95 MB V8 heap)' },
    { id: 'numbers-50k', group: 'default', shape: 'numbers', cells: 50_000, description: '50k number cells (~165 MB)' },
    { id: 'numbers-100k', group: 'default', shape: 'numbers', cells: 100_000, description: '100k number cells, the default cell limit (~300 MB)' },
    { id: 'strings-100k', group: 'default', shape: 'strings', cells: 100_000, description: '100k shared-string cells (~300 MB)' },
    { id: 'formulas-75k', group: 'default', shape: 'formulas', cells: 75_000, description: '75k formula cells, near the default node limit (~310 MB)' },
    { id: 'worst-case', group: 'default', shape: 'worst-case', cells: 0, description: 'Hostile file at the node/attribute limits (~475 MB)' },
    { id: 'numbers-200k', group: 'headroom', shape: 'numbers', cells: 200_000, limits: raised(2), description: '200k cells with 2x limits (~570 MB)' },
    { id: 'numbers-300k', group: 'headroom', shape: 'numbers', cells: 300_000, limits: raised(3), description: '300k cells with 3x limits (~850 MB)' },
];

const KEY = 'xlsx-to-js-memory-check';
const load = (): Saved => { try { return JSON.parse(localStorage.getItem(KEY) ?? '') as Saved; } catch { return { results: [], running: null }; } };
const save = (state: Saved) => { try { localStorage.setItem(KEY, JSON.stringify(state)); } catch { /* Results still render. */ } };

const generate = (request: GenerateRequest) => new Promise<ArrayBuffer>((resolve, reject) => {
    const worker = new Worker(new URL('./generator.worker.ts', import.meta.url), { type: 'module' });
    worker.onmessage = event => { worker.terminate(); resolve(event.data as ArrayBuffer); };
    worker.onerror = event => { worker.terminate(); reject(new Error(event.message || 'Generator failed')); };
    worker.postMessage(request);
});

const element = <K extends keyof HTMLElementTagNameMap>(tag: K, text = '') => { const e = document.createElement(tag); e.textContent = text; return e; };

export function mountMemoryCheck(root: HTMLElement): void {
    const state = load();
    // A marker left by a run that never finished means the tab was killed, most often for memory.
    if (state.running) { state.results.push({ id: state.running, status: 'crashed: tab reloaded during this case' }); state.running = null; save(state); }

    const device = element('pre', [
        `userAgent: ${navigator.userAgent}`,
        `deviceMemory: ${(navigator as Navigator & { deviceMemory?: number }).deviceMemory ?? 'n/a'} GB (Chromium only, rounded)`,
        `hardwareConcurrency: ${navigator.hardwareConcurrency}`,
    ].join('\n'));
    const table = element('table'), status = element('p', 'Idle.'), output = element('pre');
    const buttons = element('div');
    const button = (label: string, onClick: () => void) => { const b = element('button', label); b.onclick = onClick; buttons.append(b); return b; };

    const render = () => {
        table.replaceChildren();
        const head = table.createTHead().insertRow();
        for (const h of ['Case', 'Result', 'Generate ms', 'Parse ms (worker)', 'First page ms', 'ZIP KB']) head.append(element('th', h));
        const body = table.createTBody();
        for (const c of CASES) {
            const result = [...state.results].reverse().find(r => r.id === c.id);
            const row = body.insertRow();
            row.dataset.case = c.id;
            row.dataset.status = result?.status.split(':')[0] ?? 'pending';
            const name = element('td'); name.append(element('strong', c.id), element('br'), element('small', c.description));
            row.append(name);
            for (const value of [result?.status ?? 'not run', result?.generateMs, result?.parseMs, result?.renderMs, result?.zipKB]) row.append(element('td', value === undefined ? '' : String(value)));
        }
        output.textContent = JSON.stringify({ device: device.textContent, results: state.results }, null, 2);
    };

    let running = false;
    const run = async (cases: Case[]) => {
        if (running) return;
        running = true;
        for (const c of cases) {
            status.textContent = `Running ${c.id}...`;
            state.running = c.id; save(state);
            const result: Result = { id: c.id, status: 'ok' };
            try {
                let start = performance.now();
                const bytes = await generate(c);
                result.generateMs = Math.round(performance.now() - start);
                result.zipKB = Math.round(bytes.byteLength / 1024);
                const limits = { ...DEFAULT_LIMITS, ...c.limits };
                const parser = new XlsxWorkerParser(() => new Worker(new URL('./parser.worker.ts', import.meta.url), { type: 'module' }));
                start = performance.now();
                const workbook = await parser.readFile(bytes, { styles: true, limits, timeoutMs: 120_000 });
                result.parseMs = Math.round(performance.now() - start);
                start = performance.now();
                new XlsxParser().toHTMLSheetPage(workbook, 0, { limits: { ...limits, maxHtmlLength: 2 * 1024 * 1024 } });
                result.renderMs = Math.round(performance.now() - start);
            } catch (error) {
                result.status = `error: ${error instanceof Error ? `${error.name}: ${error.message}` : String(error)}`;
            }
            state.running = null; state.results.push(result); save(state); render();
            // Let the browser reclaim the terminated worker before the next, larger case.
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        status.textContent = 'Done. Copy the JSON below into the results notes.';
        running = false;
    };

    button('Run default-limit cases', () => void run(CASES.filter(c => c.group === 'default')));
    button('Run headroom cases (raised limits)', () => void run(CASES.filter(c => c.group === 'headroom')));
    button('Clear results', () => { state.results = []; state.running = null; save(state); render(); });
    button('Copy JSON', () => void navigator.clipboard?.writeText(output.textContent ?? '').catch(() => undefined));

    root.append(
        element('h1', 'xlsx-to-js device memory check'),
        element('p', 'Runs the production parser worker with DEFAULT_LIMITS on generated workbooks. Run on each target device with other tabs closed. If the tab reloads, the case that was running is recorded as a crash.'),
        device, buttons, status, table, element('h2', 'Results JSON'), output,
    );
    render();
    const autorun = new URLSearchParams(location.search).get('cases');
    if (autorun) void run(CASES.filter(c => autorun.split(',').includes(c.id)));
}
