import JSZip from 'jszip';
import { XlsxLimits } from '../types';
import { getChildrenByName } from './utils';
import { parseXml } from './xml';
import { limitError } from './security';

interface Entry { size: number; crc: number }
interface ChunkStream {
    on(event: 'data', callback: (chunk: Uint8Array) => void): ChunkStream;
    on(event: 'error', callback: (error: Error) => void): ChunkStream;
    on(event: 'end', callback: () => void): ChunkStream;
    pause(): void;
    resume(): void;
}

function validPath(path: string): boolean {
    return !!path && !/[\\\u0000-\u001f\u007f:#?%]/.test(path) && !path.startsWith('/') &&
        !path.split('/').some(p => p === '.' || p === '..' || p === '');
}

/** Validate the directory before JSZip allocates entries or decompresses anything.
 * ZIP64, encrypted archives and multi-volume archives are deliberately unsupported.
 */
function inspectZip(file: ArrayBuffer, limits: XlsxLimits): Map<string, Entry> {
    if (file.byteLength > limits.maxFileBytes) throw limitError(limits, 'maxFileBytes', 'Archive exceeds file budget');
    const view = new DataView(file);
    let eocd = -1;
    for (let p = file.byteLength - 22; p >= Math.max(0, file.byteLength - 65557); p--) {
        if (view.getUint32(p, true) === 0x06054b50 && p + 22 + view.getUint16(p + 20, true) === file.byteLength) { eocd = p; break; }
    }
    if (eocd < 0) throw new Error('Invalid ZIP directory');
    const count = view.getUint16(eocd + 10, true);
    const length = view.getUint32(eocd + 12, true);
    const offset = view.getUint32(eocd + 16, true);
    if (view.getUint16(eocd + 4, true) || view.getUint16(eocd + 6, true) ||
        view.getUint16(eocd + 8, true) !== count || count === 0xffff || length === 0xffffffff || offset === 0xffffffff) {
        throw new Error('Unsupported ZIP64 or multi-volume archive');
    }
    if (count > limits.maxEntries) throw limitError(limits, 'maxEntries', 'Archive entry count exceeds resource limits');
    if (offset + length !== eocd) throw new Error('Invalid ZIP directory bounds');
    const entries = new Map<string, Entry>();
    const names = new Set<string>();
    const decoder = new TextDecoder('utf-8', { fatal: true });
    let total = 0, p = offset;
    for (let i = 0; i < count; i++) {
        if (p + 46 > eocd || view.getUint32(p, true) !== 0x02014b50) throw new Error('Invalid ZIP entry');
        const flags = view.getUint16(p + 8, true), method = view.getUint16(p + 10, true);
        const compressed = view.getUint32(p + 20, true), size = view.getUint32(p + 24, true);
        const nameLen = view.getUint16(p + 28, true), extraLen = view.getUint16(p + 30, true), commentLen = view.getUint16(p + 32, true);
        const local = view.getUint32(p + 42, true);
        const next = p + 46 + nameLen + extraLen + commentLen;
        if (next > eocd || local + 30 > offset || view.getUint32(local, true) !== 0x04034b50) throw new Error('Invalid ZIP entry bounds');
        if ((flags & 1) || ![0, 8].includes(method) || view.getUint16(p + 34, true) ||
            compressed === 0xffffffff || size === 0xffffffff || local === 0xffffffff) throw new Error('Unsupported ZIP entry');
        const rawName = new Uint8Array(file, p + 46, nameLen);
        const name = decoder.decode(rawName);
        const directory = name.endsWith('/');
        if (!validPath(directory ? name.slice(0, -1) : name) || names.has(name)) throw new Error('Unsafe or duplicate ZIP path');
        names.add(name);
        const localNameLen = view.getUint16(local + 26, true), localExtraLen = view.getUint16(local + 28, true);
        if (local + 30 + localNameLen + localExtraLen + compressed > offset || localNameLen !== nameLen ||
            view.getUint16(local + 6, true) !== flags || view.getUint16(local + 8, true) !== method) throw new Error('Inconsistent ZIP entry');
        const localName = new Uint8Array(file, local + 30, localNameLen);
        if (rawName.some((byte, j) => byte !== localName[j])) throw new Error('Inconsistent ZIP entry name');
        if (size > limits.maxEntryBytes) throw limitError(limits, 'maxEntryBytes', `Archive expansion exceeds resource limits: ${name} expands to ${size} bytes`);
        if ((total += size) > limits.maxTotalBytes) throw limitError(limits, 'maxTotalBytes', 'Archive expansion exceeds resource limits');
        // Unicode path overrides can change the name JSZip sees. Reject overrides
        // rather than allowing directory validation to be bypassed.
        for (let q = p + 46 + nameLen; q < p + 46 + nameLen + extraLen;) {
            if (q + 4 > p + 46 + nameLen + extraLen) throw new Error('Invalid ZIP extra field');
            const tag = view.getUint16(q, true), len = view.getUint16(q + 2, true);
            if (tag === 0x0001 || tag === 0x7075) throw new Error('Unsupported ZIP name override or ZIP64 entry');
            q += 4 + len;
            if (q > p + 46 + nameLen + extraLen) throw new Error('Invalid ZIP extra field');
        }
        if (!directory) entries.set(name, { size, crc: view.getUint32(p + 16, true) });
        p = next;
    }
    if (p !== eocd) throw new Error('Invalid ZIP directory length');
    return entries;
}

const crcTable = Uint32Array.from({ length: 256 }, (_, i) => {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
});

export class Archive {
    private consumed = 0;
    private constructor(private zip: JSZip, private entries: Map<string, Entry>, private limits: XlsxLimits) {}

    static async open(file: ArrayBuffer, limits: XlsxLimits): Promise<Archive> {
        const entries = inspectZip(file, limits);
        const zip = await JSZip.loadAsync(file);
        for (const name of entries.keys()) if (!zip.file(name)) throw new Error('Inconsistent archive path');
        return new Archive(zip, entries, limits);
    }

    has(path: string): boolean { return this.entries.has(path); }

    async bytes(path: string): Promise<Uint8Array> {
        const entry = this.entries.get(path), file = this.zip.file(path);
        if (!entry || !file) throw new Error('Missing required package part');
        // JSZip's documented internalStream API is absent from its TS declarations.
        const stream = (file as JSZip.JSZipObject & { internalStream(type: string): ChunkStream }).internalStream('uint8array');
        return new Promise((resolve, reject) => {
            const chunks: Uint8Array[] = [];
            let size = 0, crc = 0xffffffff, settled = false;
            const fail = (error: Error) => { if (!settled) { settled = true; stream.pause(); chunks.length = 0; reject(error); } };
            stream.on('data', chunk => {
                if (settled) return;
                size += chunk.byteLength;
                this.consumed += chunk.byteLength;
                // Exceeding the declared size means forged ZIP metadata, not a budget choice.
                if (size > entry.size) { fail(new Error('Decompressed data exceeds declared ZIP size')); return; }
                if (size > this.limits.maxEntryBytes) { fail(limitError(this.limits, 'maxEntryBytes', 'Decompressed data exceeds resource limits')); return; }
                if (this.consumed > this.limits.maxTotalBytes) { fail(limitError(this.limits, 'maxTotalBytes', 'Decompressed data exceeds resource limits')); return; }
                for (let i = 0; i < chunk.length; i++) crc = crcTable[(crc ^ chunk[i]) & 255] ^ (crc >>> 8);
                chunks.push(chunk);
            }).on('error', fail).on('end', () => {
                if (settled) return;
                if (size !== entry.size || ((crc ^ 0xffffffff) >>> 0) !== entry.crc) { fail(new Error('ZIP size or checksum mismatch')); return; }
                settled = true;
                const result = new Uint8Array(size);
                let offset = 0;
                for (const chunk of chunks) { result.set(chunk, offset); offset += chunk.length; }
                resolve(result);
            });
            stream.resume();
        });
    }

    async text(path: string): Promise<string> {
        return new TextDecoder('utf-8', { fatal: true }).decode(await this.bytes(path));
    }
}

export interface Relationship { id: string; type: string; target: string; external: boolean }

export function relationshipPart(path: string): string {
    const slash = path.lastIndexOf('/');
    return `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
}

export function resolvePart(source: string, target: string): string {
    if (!target || /[\\\u0000-\u001f\u007f:#?%]/.test(target)) throw new Error('Unsafe relationship target');
    const parts = target.startsWith('/') ? [] : source.split('/').slice(0, -1);
    for (const segment of target.replace(/^\//, '').split('/')) {
        if (segment === '..') { if (!parts.length) throw new Error('Relationship escapes package'); parts.pop(); }
        else if (segment !== '.') { if (!segment) throw new Error('Invalid relationship target'); parts.push(segment); }
    }
    const path = parts.join('/');
    if (!validPath(path)) throw new Error('Invalid relationship target');
    return path;
}

export async function readRelationships(archive: Archive, source: string, limits: XlsxLimits): Promise<Map<string, Relationship>> {
    const path = source ? relationshipPart(source) : '_rels/.rels';
    const result = new Map<string, Relationship>();
    if (!archive.has(path)) return result;
    const doc = parseXml(await archive.text(path), 'Relationships', limits);
    for (const element of getChildrenByName(doc.documentElement, 'Relationship')) {
        const id = element.getAttribute('Id') ?? '', type = element.getAttribute('Type') ?? '', target = element.getAttribute('Target') ?? '';
        const mode = element.getAttribute('TargetMode') ?? 'Internal';
        if (!id || !type || !target || result.has(id) || !['Internal', 'External'].includes(mode)) throw new Error('Invalid package relationship');
        const external = mode === 'External';
        result.set(id, { id, type, external, target: external ? target : resolvePart(source, target) });
    }
    return result;
}

export function relationshipType(rel: Relationship, kind: string): boolean {
    return rel.type === `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}` ||
        rel.type === `http://purl.oclc.org/ooxml/officeDocument/relationships/${kind}`;
}

export function toBase64(bytes: Uint8Array): string {
    let binary = '';
    for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
    return btoa(binary);
}
