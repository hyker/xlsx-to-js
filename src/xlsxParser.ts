import { XlsxParserOptions, XlsxRenderOptions, XlsxPageOptions, XlsxSheetPage } from './types';
import { parseThemeXml, Theme } from './core/theme';
import { parseStylesXml, StyleSheet } from './core/style';
import { parseWorkbookXml, Workbook } from './core/workbook';
import { parseSharedStringsXml } from './core/sharedString';
import { parseWorksheetXml } from './core/worksheet';
import { Drawing, MediaFile } from './core/drawing/types';
import { parseDrawingXml } from './core/drawing';
import { Archive, readRelationships, Relationship, relationshipType, toBase64 } from './core/archive';
import { assertWorkbookTextBudget, limitError, parseRange, resolveLimits } from './core/security';
import { renderWorkbook, renderSheetPage } from './core/render';
import { chargeImagePixels, inspectImage } from './core/image';

export class XlsxParser {
    /** Bounded preview with original coordinates and merges clipped to the page. */
    toHTMLSheetPage(workbook: Workbook, sheetIndex: number, options: XlsxPageOptions = {}): XlsxSheetPage {
        return renderSheetPage(workbook, sheetIndex, options);
    }
    /** Render visible sheets using the same layout and validation as toHTMLSheet. */
    toHTML(workbook: Workbook, options: XlsxRenderOptions = {}): string {
        return renderWorkbook(workbook, undefined, options);
    }

    /** Sheet indexes refer to the original workbook order, including hidden sheets. */
    toHTMLSheet(workbook: Workbook, sheetIndex: number, options: XlsxRenderOptions = {}): string {
        if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= workbook.workSheets.length) return '';
        return renderWorkbook(workbook, sheetIndex, options);
    }

    async readFile(file: ArrayBuffer, options: XlsxParserOptions = {}): Promise<Workbook> {
        const checkAbort = () => { if (options.signal?.aborted) throw new DOMException('Parsing cancelled', 'AbortError'); };
        checkAbort();
        const limits = resolveLimits(options.limits);
        if (options.locale !== undefined) Intl.getCanonicalLocales(options.locale);
        options.onProgress?.({ phase: 'archive', completedSheets: 0, totalSheets: 0 });
        checkAbort();
        const archive = await Archive.open(file, limits);
        checkAbort();
        const rootRels = await readRelationships(archive, '', limits);
        const root = [...rootRels.values()].filter(rel => relationshipType(rel, 'officeDocument'));
        if (root.length !== 1 || root[0].external) throw new Error('Invalid workbook relationship');
        const workbookPath = root[0].target;
        const workbook = parseWorkbookXml(await archive.text(workbookPath), limits);
        options.onProgress?.({ phase: 'metadata', completedSheets: 0, totalSheets: workbook.workSheets.length });
        const relationships = await readRelationships(archive, workbookPath, limits);
        const findPart = (kind: string): Relationship | undefined => {
            const rels = [...relationships.values()].filter(rel => relationshipType(rel, kind));
            if (rels.length > 1 || rels[0]?.external) throw new Error(`Invalid ${kind} relationship`);
            return rels[0];
        };
        let themes: Theme[] = [];
        const theme = findPart('theme');
        if (theme && (options.styles || options.drawings)) themes = parseThemeXml(await archive.text(theme.target), limits);
        let style: StyleSheet = { fonts: [], fills: [], borders: [], cells: [] };
        const styles = findPart('styles');
        if (styles) style = parseStylesXml(await archive.text(styles.target), themes, limits);
        const strings = findPart('sharedStrings');
        const sharedStrings = strings ? parseSharedStringsXml(await archive.text(strings.target), limits) : [];
        const mediaCache = new Map<string, MediaFile>();
        const imagePixels = new Map<string, number>();
        const drawingCache = new Map<string, Drawing[]>();
        const sheetParts = new Set<string>();
        let cells = 0, merges = 0, drawingCount = 0, totalImagePixels = 0;
        let completedSheets = 0;
        for (const sheet of workbook.workSheets) {
            checkAbort();
            options.onProgress?.({ phase: 'worksheet', completedSheets: completedSheets++, totalSheets: workbook.workSheets.length, sheetName: sheet.name });
            const rel = relationships.get(sheet.relationshipId!);
            if (!rel || rel.external || !relationshipType(rel, 'worksheet') || sheetParts.has(rel.target)) throw new Error('Invalid worksheet relationship');
            sheetParts.add(rel.target);
            const xml = await archive.text(rel.target);
            checkAbort();
            const parsed = parseWorksheetXml(xml, style, sharedStrings, options.dense ?? false, options.skipHiddenRows ?? false, limits, workbook.date1904, options.styles ?? false,
                { cells: limits.maxCells - cells, merges: limits.maxMergedCells - merges }, options.locale);
            if (parsed.dimention) {
                const { end } = parseRange(parsed.dimention, limits);
                cells += end.row * end.col;
            }
            for (const range of parsed.mergeCells) merges += parseRange(range, limits).area;
            if (cells > limits.maxCells) throw limitError(limits, 'maxCells', `Workbook grid exceeds resource limits at sheet "${sheet.name}"`);
            if (merges > limits.maxMergedCells) throw limitError(limits, 'maxMergedCells', `Workbook merged cells exceed resource limits at sheet "${sheet.name}"`);
            Object.assign(sheet, parsed, { id: sheet.id, name: sheet.name, state: sheet.state, relationshipId: sheet.relationshipId });
            if (!options.drawings) continue;
            if (parsed.drawingRelationshipId === undefined) continue;
            const sheetRels = await readRelationships(archive, rel.target, limits);
            const drawingRel = sheetRels.get(parsed.drawingRelationshipId);
            if (!drawingRel || drawingRel.external || !relationshipType(drawingRel, 'drawing')) throw new Error('Invalid drawing relationship');
            let drawings = drawingCache.get(drawingRel.target);
            if (!drawings) {
                const imageRels = await readRelationships(archive, drawingRel.target, limits);
                const media: MediaFile[] = [];
                for (const imageRel of imageRels.values()) {
                    if (!relationshipType(imageRel, 'image')) continue;
                    if (imageRel.external) throw new Error('External drawing images are unsupported');
                    let image = mediaCache.get(imageRel.target);
                    if (!image) {
                        const bytes = await archive.bytes(imageRel.target);
                        const info = inspectImage(bytes, limits);
                        image = { name: imageRel.target, base64: toBase64(bytes) };
                        imagePixels.set(image.base64, info.pixels);
                        mediaCache.set(imageRel.target, image);
                    }
                    media.push(image);
                }
                drawings = parseDrawingXml(await archive.text(drawingRel.target), imageRels, media, themes, limits);
                drawingCache.set(drawingRel.target, drawings);
            }
            drawingCount += drawings.length;
            if (drawingCount > limits.maxDrawings) throw limitError(limits, 'maxDrawings', 'Workbook drawings exceed resource limits');
            for (const drawing of drawings) if (drawing.type === 'image') {
                totalImagePixels = chargeImagePixels(imagePixels.get(drawing.base64)!, totalImagePixels, limits);
            }
            sheet.drawings = drawings;
        }
        checkAbort();
        assertWorkbookTextBudget(workbook, limits);
        options.onProgress?.({ phase: 'complete', completedSheets, totalSheets: workbook.workSheets.length });
        return workbook;
    }
}
