import { XlsxLimits } from '../types';
import { limitError, resolveLimits } from './security';

const NS: Record<string, readonly string[]> = {
    s: ['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main'],
    a: ['http://schemas.openxmlformats.org/drawingml/2006/main', 'http://purl.oclc.org/ooxml/drawingml/main'],
    xdr: ['http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing'],
    rel: ['http://schemas.openxmlformats.org/package/2006/relationships', 'http://purl.oclc.org/ooxml/package/relationships'],
};

export function matchesName(element: Element, name: string): boolean {
    const [group, local] = name.includes(':') ? name.split(':') : ['s', name];
    const namespaces = name === 'Relationships' || name === 'Relationship' ? NS.rel : NS[group];
    // Unqualified XML is accepted for compatibility; prefixed names use namespace URIs.
    return element.localName === local && (!element.namespaceURI || !!namespaces?.includes(element.namespaceURI));
}

export function parseXml(str: string, rootName: string, limits: XlsxLimits = resolveLimits()): Document {
    if (str.length > limits.maxEntryBytes) throw limitError(limits, 'maxEntryBytes', 'XML exceeds entry budget');
    if (/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(str)) throw new Error('DTD and entity declarations are prohibited');
    preflightXml(str, limits, rootName);
    const doc = new DOMParser().parseFromString(str, 'text/xml');
    if (!doc.documentElement || doc.getElementsByTagNameNS('*', 'parsererror').length || !matchesName(doc.documentElement, rootName)) {
        throw new Error(`Invalid ${rootName} XML`);
    }
    const stack: { node: Node; depth: number }[] = [{ node: doc.documentElement, depth: 1 }];
    let nodes = 0;
    while (stack.length) {
        const { node, depth } = stack.pop()!;
        if (++nodes > limits.maxXmlNodes) throw limitError(limits, 'maxXmlNodes', `XML complexity exceeds resource limits: ${rootName} part has too many XML nodes`);
        if (depth > limits.maxXmlDepth) throw limitError(limits, 'maxXmlDepth', `XML complexity exceeds resource limits: ${rootName} part is nested too deeply`);
        for (let child = node.firstChild; child; child = child.nextSibling) stack.push({ node: child, depth: depth + 1 });
    }
    return doc;
}

/** A lexical budget pass, not an XML validator. No DOM or token arrays are built.
 * Quoted attributes, comments, CDATA and PIs are skipped as indivisible regions;
 * the real XML parser remains responsible for well-formedness and namespaces.
 */
function preflightXml(xml: string, limits: XlsxLimits, rootName: string): void {
    let nodes = 0, depth = 0, attributes = 0, i = 0;
    const node = (atDepth: number) => {
        if (++nodes > limits.maxXmlNodes) throw limitError(limits, 'maxXmlNodes', `XML complexity exceeds resource limits: ${rootName} part has too many XML nodes`);
        if (atDepth > limits.maxXmlDepth) throw limitError(limits, 'maxXmlDepth', `XML complexity exceeds resource limits: ${rootName} part is nested too deeply`);
    };
    const region = (end: string, start: number) => {
        const close = xml.indexOf(end, start);
        if (close < 0) throw new Error('Invalid XML: unterminated markup');
        return close + end.length;
    };
    while (i < xml.length) {
        if (xml[i] !== '<') {
            const next = xml.indexOf('<', i);
            // Include document-level nodes conservatively; they allocate too.
            node(depth + 1);
            i = next < 0 ? xml.length : next;
        } else if (xml.startsWith('<!--', i)) {
            node(depth + 1); i = region('-->', i + 4);
        } else if (xml.startsWith('<![CDATA[', i)) {
            node(depth + 1); i = region(']]>', i + 9);
        } else if (xml.startsWith('<?', i)) {
            node(depth + 1); i = region('?>', i + 2);
        } else if (xml.startsWith('<!', i)) {
            throw new Error('Invalid XML declaration');
        } else if (xml.startsWith('</', i)) {
            const end = xml.indexOf('>', i + 2);
            if (end < 0 || --depth < 0) throw new Error('Invalid XML closing tag');
            i = end + 1;
        } else {
            node(++depth);
            let perElement = 0;
            i++;
            while (i < xml.length && xml[i] !== '>') {
                const ch = xml[i++];
                if (ch === '=') {
                    if (++attributes > limits.maxXmlAttributes) throw limitError(limits, 'maxXmlAttributes', `XML attribute complexity exceeds resource limits: ${rootName} part has too many attributes`);
                    if (++perElement > limits.maxXmlAttributesPerElement) throw limitError(limits, 'maxXmlAttributesPerElement', `XML attribute complexity exceeds resource limits: an element in the ${rootName} part has too many attributes`);
                } else if (ch === '"' || ch === "'") {
                    i = region(ch, i);
                } else if (ch === '<') throw new Error('Invalid XML tag');
            }
            if (i === xml.length) throw new Error('Invalid XML: unterminated tag');
            if (xml[i - 1] === '/') depth--;
            i++;
        }
    }
    if (depth !== 0) throw new Error('Invalid XML: unclosed tag');
}

export function relationshipId(element: Element): string {
    return element.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
        ?? element.getAttributeNS('http://purl.oclc.org/ooxml/officeDocument/relationships', 'id') ?? '';
}
