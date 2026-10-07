import { DOMParser as XmlDOMParser } from '@xmldom/xmldom';
import { XlsxParser } from './xlsxParser';
import { XlsxParserOptions } from './types';

// xmldom can recover from malformed XML. Escalate every diagnostic so the worker
// preserves the strict rejection semantics of native DOMParser.
class StrictDOMParser {
    parseFromString(source: string, type: string): Document {
        return new XmlDOMParser({ onError: (_level, message) => { throw new Error(`Invalid XML: ${message}`); } })
            .parseFromString(source, type as 'text/xml') as unknown as Document;
    }
}

globalThis.DOMParser = StrictDOMParser as unknown as typeof DOMParser;
const scope = globalThis as unknown as {
    onmessage: ((event: MessageEvent<{ file: ArrayBuffer; options: XlsxParserOptions }>) => void) | null;
    postMessage: (message: unknown) => void;
};
scope.onmessage = async event => {
    try {
        const workbook = await new XlsxParser().readFile(event.data.file, {
            ...event.data.options,
            onProgress: progress => scope.postMessage({ type: 'progress', progress }),
        });
        scope.postMessage({ type: 'result', workbook });
    } catch (error) {
        scope.postMessage({ type: 'error', error: error instanceof Error ? error.message : String(error) });
    }
};
