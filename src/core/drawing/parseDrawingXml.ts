import { parseXml } from '../xml';
import { XlsxLimits } from '../../types';
import { boundedNumber, limitError, resolveLimits, safeColor } from '../security';
import { Relationship, relationshipType } from '../archive';
import { argbToHex, getElementByName, getElementsByName, hexToRgb, rgbToHex } from "../utils";
import { Drawing, DrawingObjectType, DrawingPosition, MediaFile } from "./types";
import { Theme } from "../theme/types";

export const parseDrawingXml = (drawingStr: string, rels: Map<string, Relationship>, media: MediaFile[], themes: Theme[] = [], limits: XlsxLimits = resolveLimits()): Drawing[] => {
    const drawings: Drawing[] = [];
    const xmlDoc = parseXml(drawingStr, 'xdr:wsDr', limits);
    const drawingElement = xmlDoc.documentElement;
    const drawingsArray = [
        ...getElementsByName(drawingElement, 'xdr:twoCellAnchor'),
        ...getElementsByName(drawingElement, 'xdr:oneCellAnchor'),
        ...getElementsByName(drawingElement, 'xdr:absoluteAnchor'),
    ];
    if (drawingsArray.length > limits.maxDrawings) throw limitError(limits, 'maxDrawings', 'Drawing count exceeds resource limits');
    const emu = (value: string | null | undefined) => boundedNumber(value ?? '0', 'drawing geometry', limits.maxDrawingPixels * 9525, 0, true);

    if (drawingsArray) {
        drawingsArray.forEach(anchor => {
            const anchorType = anchor.tagName.includes('oneCellAnchor')
                ? 'oneCell'
                : (anchor.tagName.includes('absoluteAnchor') ? 'absolute' : 'twoCell');
            const elementTypeMap: { [tag: string]: DrawingObjectType } = {
                'xdr:pic': 'image',
                'xdr:sp': 'shape',
                'xdr:cxnSp': 'connector',
                'xdr:grpSp': 'group'
            };

            let drawingType: DrawingObjectType = 'unknown';
            let containerElement: Element | undefined;

            for (const [tag, detectedType] of Object.entries(elementTypeMap)) {
                const foundElement = getElementByName(anchor, tag);
                if (foundElement) {
                    drawingType = detectedType;
                    containerElement = foundElement;

                    if (drawingType === 'shape' && getElementByName(containerElement, 'xdr:txBody')) {
                        drawingType = 'textbox';
                    }
                    break;
                }
            }

            // Fallback: explicit textbox element (xdr:txbx)
            if (!containerElement) {
                const txbx = getElementByName(anchor, 'xdr:txbx');
                if (txbx) {
                    drawingType = 'textbox';
                    containerElement = txbx;
                }
            }

            // Obtener metadata del dibujo
            const cNvPr = getElementByName(anchor, 'xdr:cNvPr');
            const id = cNvPr?.getAttribute('id') || '';
            const name = cNvPr?.getAttribute('name') || '';
            const title = cNvPr?.getAttribute('title') || '';
            const description = cNvPr?.getAttribute('descr') || '';

            // Obtener la imagen asociada
            const blip = getElementByName(anchor, 'a:blip');
            const embedId = blip?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') ?? blip?.getAttributeNS('http://purl.oclc.org/ooxml/officeDocument/relationships', 'embed') ?? '';
            if (blip?.hasAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'link') || blip?.hasAttributeNS('http://purl.oclc.org/ooxml/officeDocument/relationships', 'link')) throw new Error('External drawing images are unsupported');
            const rel = rels.get(embedId);
            if (drawingType === 'image' && (!rel || rel.external || !relationshipType(rel, 'image'))) throw new Error('Invalid drawing image relationship');
            const mediaFile = media.find(m => m.name === rel?.target);
            if (drawingType === 'image' && !mediaFile) throw new Error('Missing drawing image');

            // Posiciones
            const from = getElementByName(anchor, 'xdr:from');
            const to = getElementByName(anchor, 'xdr:to');
            const ext = getElementByName(anchor, 'xdr:ext'); // size for oneCell/absolute
            const pos = getElementByName(anchor, 'xdr:pos'); // absolute pos

            if ((anchorType !== 'absolute' && !from) || (anchorType === 'twoCell' && !to) || (anchorType !== 'twoCell' && !ext) || (anchorType === 'absolute' && !pos)) throw new Error('Incomplete drawing anchor');

            const getPosition = (posElement: Element | undefined) => ({
                col: boundedNumber(getElementByName(posElement, 'xdr:col')?.textContent ?? '0', 'drawing column', limits.maxColumns - 1, 0, true),
                colOff: emu(getElementByName(posElement, 'xdr:colOff')?.textContent),
                row: boundedNumber(getElementByName(posElement, 'xdr:row')?.textContent ?? '0', 'drawing row', limits.maxRows - 1, 0, true),
                rowOff: emu(getElementByName(posElement, 'xdr:rowOff')?.textContent),
            });

            const position: DrawingPosition = {
                from: getPosition(from),
                to: to ? getPosition(to) : getPosition(from),
            };

            if (position.to.col < position.from.col || position.to.row < position.from.row) throw new Error('Reversed drawing anchor');

            let sizeEMU = (anchorType === 'oneCell' || anchorType === 'absolute') && ext
                ? {
                    cx: emu(ext.getAttribute('cx')),
                    cy: emu(ext.getAttribute('cy')),
                }
                : undefined;
            const absEMU = anchorType === 'absolute' && pos
                ? {
                    x: emu(pos.getAttribute('x')),
                    y: emu(pos.getAttribute('y')),
                }
                : undefined;

            // Try to override size from shape transform if present
            if (containerElement) {
                const xfrm = getElementByName(containerElement, 'a:xfrm');
                const xfrmExt = getElementByName(xfrm, 'a:ext');
                if (xfrmExt) {
                    const cx = emu(xfrmExt.getAttribute('cx'));
                    const cy = emu(xfrmExt.getAttribute('cy'));
                    if (cx > 0 || cy > 0) {
                        // prefer explicit transform ext when available
                        sizeEMU = { cx: cx || 0, cy: cy || 0 };
                    }
                }
            }

            const properties = extractPropertiesByType(drawingType, containerElement, themes);

            drawings.push({
                id,
                name,
                title,
                description,
                type: drawingType,
                base64: mediaFile?.base64 || '',
                anchorType,
                sizeEMU,
                absEMU,
                properties,
                position,
            });
        });
    }

    return drawings;
};

function extractPropertiesByType(type: DrawingObjectType, container: Element | undefined, themes: Theme[]): Record<string, any> {
    if (!container) return {};

    const colorFrom = (root?: Element): string => {
        if (!root) return '';
        const srgb = getElementByName(root, 'a:srgbClr');
        if (srgb) {
            const color = safeColor(`#${srgb.getAttribute('val')}`);
            if (!color) throw new Error('Invalid drawing color');
            return color;
        }
        const scheme = getElementByName(root, 'a:schemeClr');
        if (scheme) {
            const name = scheme.getAttribute('val') || '';
            const t = themes.find(th => th.name === name);
            let hex = t?.val ? argbToHex(t.val) : '';
            const shade = getElementByName(scheme, 'a:shade')?.getAttribute('val');
            const tint = getElementByName(scheme, 'a:tint')?.getAttribute('val');
            if (hex) {
                const rgb = hexToRgb(hex);
                if (shade) {
                    const f = boundedNumber(shade, 'drawing shade', 100000, 0, true) / 100000;
                    rgb.r = Math.round(rgb.r * f);
                    rgb.g = Math.round(rgb.g * f);
                    rgb.b = Math.round(rgb.b * f);
                }
                if (tint) {
                    const f = 1 - boundedNumber(tint, 'drawing tint', 100000, 0, true) / 100000;
                    rgb.r = Math.round(rgb.r * (1 - f) + 255 * f);
                    rgb.g = Math.round(rgb.g * (1 - f) + 255 * f);
                    rgb.b = Math.round(rgb.b * (1 - f) + 255 * f);
                }
                hex = rgbToHex(rgb.r, rgb.g, rgb.b);
            }
            return hex;
        }
        const scrgb = getElementByName(root, 'a:scrgbClr');
        if (scrgb) {
            const r = boundedNumber(scrgb.getAttribute('r') ?? '0', 'drawing color', 100000, 0, true);
            const g = boundedNumber(scrgb.getAttribute('g') ?? '0', 'drawing color', 100000, 0, true);
            const b = boundedNumber(scrgb.getAttribute('b') ?? '0', 'drawing color', 100000, 0, true);
            const to255 = (v: number) => Math.round(v * 255 / 100000);
            return rgbToHex(to255(r), to255(g), to255(b));
        }
        return '';
    };

    const shapeStyles = (sp: Element) => {
        const spPr = getElementByName(sp, 'xdr:spPr');
        const solidFill = getElementByName(spPr, 'a:solidFill');
        let fillColor = colorFrom(solidFill);
        const ln = getElementByName(spPr, 'a:ln');
        const lnFill = getElementByName(ln, 'a:solidFill');
        let lineColor = colorFrom(lnFill);
        const lineWidth = ln?.hasAttribute('w') ? boundedNumber(ln.getAttribute('w'), 'drawing line width', 12700000, 0, true) : 0;
        const style = getElementByName(sp, 'xdr:style');
        if (!fillColor) {
            const fillRef = getElementByName(style, 'a:fillRef');
            fillColor = colorFrom(fillRef);
        }
        if (!lineColor) {
            const lnRef = getElementByName(style, 'a:lnRef');
            lineColor = colorFrom(lnRef);
        }
        return { fillColor, lineColor, lineWidth };
    };

    switch (type) {
        case 'image':
            {
            const blip = getElementByName(container, 'a:blip');
            return {
                embedId: blip?.getAttribute('r:embed') || ''
            };
        }
        case 'textbox':
            {
            const textElements = getElementsByName(container, 'a:t');
            const text = textElements.map(el => el.textContent).join('\n');
            const rPr = getElementByName(container, 'a:rPr');
            const sz = rPr?.getAttribute('sz');
            const fontPt = sz ? boundedNumber(sz, 'drawing font size', 40900, 0, true) / 100 : undefined;
            const bold = rPr?.getAttribute('b') === '1' || rPr?.getAttribute('b') === 'true';
            const italic = rPr?.getAttribute('i') === '1' || rPr?.getAttribute('i') === 'true';
            let txtColor = colorFrom(getElementByName(rPr, 'a:solidFill'));
            // Fallback to style fontRef when run color is not set
            if (!txtColor) {
                const style = getElementByName(container, 'xdr:style');
                const fontRef = getElementByName(style, 'a:fontRef');
                txtColor = colorFrom(fontRef);
            }
            const pPr = getElementByName(container, 'a:pPr');
            const algn = pPr?.getAttribute('algn') || '';
            const textAlign = algn === 'ctr' ? 'center' : (algn === 'r' ? 'right' : 'left');
            const { fillColor, lineColor, lineWidth } = shapeStyles(container);
            return { text, fontPt, bold, italic, textColor: txtColor, textAlign, fillColor, lineColor, lineWidth };
            }
        case 'shape':
            {
            const prstGeom = getElementByName(container, 'a:prstGeom');
            const shapeType = prstGeom?.getAttribute('prst') || 'custom';
            const { fillColor, lineColor, lineWidth } = shapeStyles(container);
            return { fillColor, shapeType, lineColor, lineWidth };
            }
        case 'connector':
            {
            const line = getElementByName(container, 'a:ln');
            const lineColor = colorFrom(getElementByName(line, 'a:solidFill'));
            const lineWidth = line?.hasAttribute('w') ? boundedNumber(line.getAttribute('w'), 'drawing line width', 12700000, 0, true) : 0;
            return { lineColor, lineWidth };
            }
        case 'group':
            return {};

        default:
            return {};
    }
}
