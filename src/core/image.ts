import { XlsxLimits } from '../types';
import { limitError } from './security';

export interface ImageInfo { mime: string; pixels: number }

/** Read raster headers without decoding pixels. Animated GIF/APNG are rejected:
 * their frame buffers would otherwise escape a static-image pixel budget.
 */
export function inspectImage(data: Uint8Array | string, limits: XlsxLimits): ImageInfo {
    const length = data.length;
    const byte = (i: number): number => typeof data === 'string' ? data.charCodeAt(i) : data[i];
    const fail = (): never => { throw new Error('Invalid image data'); };
    const has = (offset: number, text: string): boolean => {
        if (offset + text.length > length) return false;
        for (let i = 0; i < text.length; i++) if (byte(offset + i) !== text.charCodeAt(i)) return false;
        return true;
    };
    const u16 = (i: number, little = false) => little ? byte(i) + byte(i + 1) * 256 : byte(i) * 256 + byte(i + 1);
    const u32 = (i: number) => byte(i) * 16777216 + byte(i + 1) * 65536 + byte(i + 2) * 256 + byte(i + 3);
    const dimensions = (width: number, height: number): number => {
        if (!width || !height) fail();
        const pixels = width * height;
        if (width > limits.maxImageDimension || height > limits.maxImageDimension) throw limitError(limits, 'maxImageDimension', `Image pixels exceed resource limits: ${width}x${height} image`);
        if (pixels > limits.maxImagePixels) throw limitError(limits, 'maxImagePixels', `Image pixels exceed resource limits: ${width}x${height} image`);
        return pixels;
    };
    if (length > limits.maxEntryBytes) fail();
    if (has(0, '\x89PNG\r\n\x1a\n')) {
        if (length < 33 || u32(8) !== 13 || !has(12, 'IHDR')) fail();
        const pixels = dimensions(u32(16), u32(20));
        let p = 8, imageData = false;
        while (p + 12 <= length) {
            const size = u32(p), next = p + 12 + size;
            if (next > length) fail();
            if (has(p + 4, 'acTL') || has(p + 4, 'fcTL') || has(p + 4, 'fdAT')) throw new Error('Animated images are unsupported');
            if (p !== 8 && has(p + 4, 'IHDR')) fail();
            if (has(p + 4, 'IDAT')) imageData = true;
            if (has(p + 4, 'IEND')) {
                if (size || next !== length || !imageData) fail();
                return { mime: 'image/png', pixels };
            }
            p = next;
        }
        fail();
    }
    if (has(0, 'GIF87a') || has(0, 'GIF89a')) {
        if (length < 13) fail();
        const width = u16(6, true), height = u16(8, true), pixels = dimensions(width, height);
        let p = 13 + ((byte(10) & 128) ? 3 * (1 << ((byte(10) & 7) + 1)) : 0), frames = 0;
        const blocks = () => {
            while (p < length) {
                const size = byte(p++);
                if (!size) return;
                if (p + size > length) fail();
                p += size;
            }
            fail();
        };
        while (p < length) {
            const marker = byte(p++);
            if (marker === 0x3b) {
                if (!frames || p !== length) fail();
                return { mime: 'image/gif', pixels };
            }
            if (marker === 0x21) {
                if (p >= length) fail();
                p++; blocks();
            } else if (marker === 0x2c) {
                if (++frames > 1) throw new Error('Animated images are unsupported');
                if (p + 9 > length) fail();
                const left = u16(p, true), top = u16(p + 2, true), w = u16(p + 4, true), h = u16(p + 6, true);
                dimensions(w, h);
                if (left + w > width || top + h > height) fail();
                const packed = byte(p + 8);
                p += 9 + ((packed & 128) ? 3 * (1 << ((packed & 7) + 1)) : 0);
                if (p >= length || byte(p) < 2 || byte(p) > 8) fail();
                p++; blocks();
            } else fail();
        }
        fail();
    }
    if (length >= 4 && byte(0) === 0xff && byte(1) === 0xd8) {
        let p = 2, pixels: number | undefined, inScan = false;
        while (p < length) {
            if (inScan) {
                while (p < length && byte(p) !== 0xff) p++;
                if (p === length) fail();
            }
            if (byte(p++) !== 0xff) fail();
            while (p < length && byte(p) === 0xff) p++;
            if (p >= length) fail();
            const marker = byte(p++);
            if (inScan && (marker === 0 || (marker >= 0xd0 && marker <= 0xd7))) continue;
            inScan = false;
            if (marker === 0xd9) {
                if (pixels === undefined || p !== length) fail();
                return { mime: 'image/jpeg', pixels: pixels! };
            }
            if (marker === 0 || marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7)) fail();
            if (marker === 1) continue;
            if (p + 2 > length) fail();
            const size = u16(p), next = p + size;
            if (size < 2 || next > length) fail();
            if (marker === 0xda) {
                if (pixels === undefined || size < 6) fail();
                inScan = true;
            }
            // Only the common single-frame baseline, extended and progressive
            // Huffman modes are supported. Hierarchical/multi-frame modes fail.
            if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
                if (![0xc0, 0xc1, 0xc2].includes(marker) || pixels !== undefined || size < 11 || byte(p + 2) !== 8) fail();
                const components = byte(p + 7);
                if (!components || components > 4 || size !== 8 + 3 * components) fail();
                pixels = dimensions(u16(p + 5), u16(p + 3));
            }
            if (marker === 0xdc || marker === 0xde || marker === 0xdf) fail();
            p = next;
        }
        fail();
    }
    throw new Error('Unsupported image format');
}

export function inspectBase64Image(base64: string, limits: XlsxLimits): ImageInfo {
    if (base64.length % 4 !== 0 || base64.length > Math.ceil(limits.maxEntryBytes / 3) * 4) throw new Error('Invalid image data');
    let end = base64.length;
    if (base64.charCodeAt(end - 1) === 61) end--;
    if (base64.charCodeAt(end - 1) === 61) end--;
    if (base64.length / 4 * 3 - (base64.length - end) > limits.maxEntryBytes) throw new Error('Invalid image data');
    for (let i = 0; i < end; i++) {
        const c = base64.charCodeAt(i);
        if (!((c >= 65 && c <= 90) || (c >= 97 && c <= 122) || (c >= 48 && c <= 57) || c === 43 || c === 47)) throw new Error('Invalid image data');
    }
    return inspectImage(atob(base64), limits);
}

export function chargeImagePixels(pixels: number, consumed: number, limits: XlsxLimits): number {
    const total = consumed + pixels;
    if (total > limits.maxTotalImagePixels) throw limitError(limits, 'maxTotalImagePixels', 'Total image pixels exceed resource limits');
    return total;
}
