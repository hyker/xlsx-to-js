import { XlsxLimits } from '../types';
import { resolveLimits } from './security';

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
    if (str.length > limits.maxEntryBytes) throw new Error('XML exceeds entry budget');
    if (/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(str)) throw new Error('DTD and entity declarations are prohibited');
    const doc = new DOMParser().parseFromString(str, 'text/xml');
    if (!doc.documentElement || doc.getElementsByTagNameNS('*', 'parsererror').length || !matchesName(doc.documentElement, rootName)) {
        throw new Error(`Invalid ${rootName} XML`);
    }
    const stack: { node: Node; depth: number }[] = [{ node: doc.documentElement, depth: 1 }];
    let nodes = 0;
    while (stack.length) {
        const { node, depth } = stack.pop()!;
        if (++nodes > limits.maxXmlNodes || depth > limits.maxXmlDepth) throw new Error('XML complexity exceeds resource limits');
        for (let child = node.firstChild; child; child = child.nextSibling) stack.push({ node: child, depth: depth + 1 });
    }
    return doc;
}

export function relationshipId(element: Element): string {
    return element.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
        ?? element.getAttributeNS('http://purl.oclc.org/ooxml/officeDocument/relationships', 'id') ?? '';
}
