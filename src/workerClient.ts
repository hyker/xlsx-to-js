import { Workbook } from './core/workbook/types';
import { XlsxParserOptions, XlsxParseProgress, XlsxRenderOptions } from './types';
import { assertWorkbookTextBudget, resolveLimits } from './core/security';

export interface XlsxWorkerOptions extends XlsxParserOptions {
    /** Wall-clock timeout, including worker startup. Default: 30 seconds. */
    timeoutMs?: number;
}

export interface XlsxWorkerRenderOptions extends XlsxRenderOptions {
    signal?: AbortSignal;
    timeoutMs?: number;
}

export interface XlsxWorkerClientOptions {
    /** Reject excess jobs before creating workers or cloning inputs. Default: 1. */
    maxConcurrentJobs?: number;
}

/** Each job owns a disposable worker; cancellation interrupts XML or export work. */
export class XlsxWorkerParser {
    private activeJobs = 0;
    private maxConcurrentJobs: number;
    constructor(private createWorker: () => Worker, options: XlsxWorkerClientOptions = {}) {
        this.maxConcurrentJobs = options.maxConcurrentJobs ?? 1;
        if (!Number.isSafeInteger(this.maxConcurrentJobs) || this.maxConcurrentJobs <= 0) throw new Error('Invalid worker concurrency limit');
    }

    readFile(file: ArrayBuffer, options: XlsxWorkerOptions = {}): Promise<Workbook> {
        return this.run<Workbook>('workbook', options, () => {
            const { signal, onProgress, timeoutMs, ...parserOptions } = options;
            const limits = resolveLimits(parserOptions.limits);
            if (file.byteLength > limits.maxFileBytes) throw new Error('Archive exceeds file budget');
            return { type: 'parse', file, options: { ...parserOptions, limits } };
        });
    }

    /** Full export in a terminable worker; font metrics use the worker fallback. */
    toHTML(workbook: Workbook, options: XlsxWorkerRenderOptions = {}): Promise<string> {
        return this.run<string>('html', options, () => {
            const { signal, timeoutMs, ...renderOptions } = options;
            const limits = resolveLimits(renderOptions.limits);
            if (workbook.workSheets.length > limits.maxSheets) throw new Error('Sheet count exceeds resource limits');
            assertWorkbookTextBudget(workbook, limits);
            return { type: 'render', workbook, options: { ...renderOptions, limits } };
        });
    }

    private run<T>(resultKey: 'workbook' | 'html', options: {
        signal?: AbortSignal; timeoutMs?: number; onProgress?: (progress: XlsxParseProgress) => void;
    }, payload: () => unknown): Promise<T> {
        return new Promise((resolve, reject) => {
            const started = performance.now();
            const { signal, onProgress, timeoutMs = 30_000 } = options;
            if (signal?.aborted) { reject(new DOMException('Parsing cancelled', 'AbortError')); return; }
            if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0 || timeoutMs > 2147483647) throw new Error('Invalid parsing timeout');
            if (this.activeJobs >= this.maxConcurrentJobs) throw new Error('Worker concurrency limit exceeded');
            const message = payload();
            if (performance.now() - started >= timeoutMs) throw new Error('Parsing exceeded time limit');
            const worker = this.createWorker();
            this.activeJobs++;
            let settled = false;
            const finish = (error?: Error, result?: T) => {
                if (settled) return;
                settled = true;
                clearTimeout(timer);
                signal?.removeEventListener('abort', abort);
                this.activeJobs--;
                worker.terminate();
                if (error) reject(error); else resolve(result!);
            };
            const abort = () => finish(new DOMException('Parsing cancelled', 'AbortError'));
            const timer = setTimeout(() => finish(new Error('Parsing exceeded time limit')), Math.max(0, timeoutMs - (performance.now() - started)));
            signal?.addEventListener('abort', abort, { once: true });
            if (signal?.aborted) { abort(); return; }
            worker.onmessage = (event: MessageEvent<{ type: string; workbook?: Workbook; html?: string; error?: string; progress?: XlsxParseProgress }>) => {
                if (settled) return;
                if (performance.now() - started >= timeoutMs) { finish(new Error('Parsing exceeded time limit')); return; }
                const data = event.data;
                if (data.type === 'progress') {
                    try { onProgress?.(data.progress!); } catch (error) { finish(error instanceof Error ? error : new Error(String(error))); }
                } else if (data.type === 'result' && (resultKey === 'html' ? typeof data.html === 'string' : !!data.workbook)) finish(undefined, data[resultKey] as T);
                else finish(new Error(data.error || 'Invalid parser worker response'));
            };
            worker.onerror = event => finish(new Error(event.message || 'Parser worker failed'));
            worker.onmessageerror = () => finish(new Error('Could not receive parser worker result'));
            // Clone inputs so callers retain their bytes and workbook for retries.
            if (performance.now() - started >= timeoutMs) { finish(new Error('Parsing exceeded time limit')); return; }
            try { worker.postMessage(message); }
            catch (error) { finish(error instanceof Error ? error : new Error(String(error))); }
        });
    }
}
