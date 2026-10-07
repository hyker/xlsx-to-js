import { Workbook } from './core/workbook/types';
import { XlsxParserOptions, XlsxParseProgress } from './types';
import { resolveLimits } from './core/security';

export interface XlsxWorkerOptions extends XlsxParserOptions {
    /** Wall-clock timeout, including worker startup. Default: 30 seconds. */
    timeoutMs?: number;
}

/** Each read owns a disposable worker, so abort interrupts synchronous XML work. */
export class XlsxWorkerParser {
    constructor(private createWorker: () => Worker) {}

    readFile(file: ArrayBuffer, options: XlsxWorkerOptions = {}): Promise<Workbook> {
        return new Promise((resolve, reject) => {
            const { signal, onProgress, timeoutMs = 30_000, ...parserOptions } = options;
            if (signal?.aborted) { reject(new DOMException('Parsing cancelled', 'AbortError')); return; }
            const limits = resolveLimits(parserOptions.limits);
            if (file.byteLength > limits.maxFileBytes) throw new Error('Archive exceeds file budget');
            if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0) throw new Error('Invalid parsing timeout');
            const worker = this.createWorker();
            let settled = false;
            const finish = (error?: Error, workbook?: Workbook) => {
                if (settled) return;
                settled = true;
                clearTimeout(timer);
                signal?.removeEventListener('abort', abort);
                worker.terminate();
                if (error) reject(error); else resolve(workbook!);
            };
            const abort = () => finish(new DOMException('Parsing cancelled', 'AbortError'));
            const timer = setTimeout(() => finish(new Error('Parsing exceeded time limit')), timeoutMs);
            signal?.addEventListener('abort', abort, { once: true });
            worker.onmessage = (event: MessageEvent<{ type: string; workbook?: Workbook; error?: string; progress?: XlsxParseProgress }>) => {
                if (settled) return;
                const data = event.data;
                if (data.type === 'progress') {
                    try { onProgress?.(data.progress!); } catch (error) { finish(error instanceof Error ? error : new Error(String(error))); }
                } else if (data.type === 'result' && data.workbook) finish(undefined, data.workbook);
                else finish(new Error(data.error || 'Invalid parser worker response'));
            };
            worker.onerror = event => finish(new Error(event.message || 'Parser worker failed'));
            worker.onmessageerror = () => finish(new Error('Could not receive parser worker result'));
            // Clone bytes rather than transferring them: callers retain their input
            // for retries and exports. The archive budget bounds the copy.
            try { worker.postMessage({ file, options: { ...parserOptions, limits } }); }
            catch (error) { finish(error instanceof Error ? error : new Error(String(error))); }
        });
    }
}
