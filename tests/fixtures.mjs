import JSZip from 'jszip';

export const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
export const DRAW = 'http://schemas.openxmlformats.org/drawingml/2006/';
export const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const esc = value => String(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
export const relationships = rels => `<Relationships xmlns="${PKG}">${rels.map(r => `<Relationship Id="${r.id}" Type="${REL}${r.kind}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
export const workbookXml = (sheets = [{name:'First',id:1,rel:'rId1'}], date1904 = false) => `<workbook xmlns="${NS}" xmlns:r="${REL.slice(0,-1)}"><workbookPr date1904="${date1904 ? 1 : 0}"/><sheets>${sheets.map(s => `<sheet name="${esc(s.name)}" sheetId="${s.id}" r:id="${s.rel}"${s.state ? ` state="${s.state}"` : ''}/>`).join('')}</sheets></workbook>`;
export const worksheetXml = (data='<row r="1"><c r="A1" s="0"><v>42</v></c></row>', dimension='A1', extras='') => `<worksheet xmlns="${NS}" xmlns:r="${REL.slice(0,-1)}">${dimension === null ? '' : `<dimension ref="${esc(dimension)}"/>`}<sheetData>${data}</sheetData>${extras}</worksheet>`;
export const stylesXml = ({font='Arial', horizontal='left', vertical='bottom', fontSize='11', color='FF000000', flags='', numFmtId=0}={}) => `<styleSheet xmlns="${NS}"><fonts><font><name val="${esc(font)}"/><sz val="${esc(fontSize)}"/>${color === 'theme' ? '<color theme="0"/>' : `<color rgb="${esc(color)}"/>`}${flags}</font></fonts><fills><fill><patternFill patternType="none"/></fill></fills><borders><border/></borders><cellXfs><xf fontId="0" fillId="0" borderId="0" numFmtId="${numFmtId}"><alignment horizontal="${esc(horizontal)}" vertical="${esc(vertical)}"/></xf></cellXfs></styleSheet>`;
export const themeXml = `<a:theme xmlns:a="${DRAW}main"><a:themeElements><a:clrScheme><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme></a:themeElements></a:theme>`;
export function drawingXml({color='336699', extent='95250', row='0', col='0', offset='0', image=false, text='', absolute=false}={}) {
    const object = image ? '<xdr:pic><xdr:blipFill><a:blip r:embed="imageRel"/></xdr:blipFill></xdr:pic>' : `<xdr:sp><xdr:spPr><a:solidFill><a:srgbClr val="${esc(color)}"/></a:solidFill></xdr:spPr>${text ? `<xdr:txBody><a:p><a:r><a:t>${esc(text)}</a:t></a:r></a:p></xdr:txBody>` : ''}</xdr:sp>`;
    return `<xdr:wsDr xmlns:xdr="${DRAW}spreadsheetDrawing" xmlns:a="${DRAW}main" xmlns:r="${REL.slice(0,-1)}"><xdr:${absolute ? 'absoluteAnchor' : 'oneCellAnchor'}>${absolute ? `<xdr:pos x="${offset}" y="${offset}"/>` : `<xdr:from><xdr:col>${col}</xdr:col><xdr:row>${row}</xdr:row><xdr:colOff>${offset}</xdr:colOff><xdr:rowOff>0</xdr:rowOff></xdr:from>`}<xdr:ext cx="${extent}" cy="95250"/>${object}</xdr:${absolute ? 'absoluteAnchor' : 'oneCellAnchor'}></xdr:wsDr>`;
}
export async function zip(parts) {
    const z = new JSZip();
    for (const [path, value] of parts) z.file(path, value);
    return z.generateAsync({type:'arraybuffer',compression:'DEFLATE'});
}
export async function fixture({sheet=worksheetXml(), style=stylesXml(), theme, shared, drawing, drawingRelationships, extra=[], sheets, date1904=false, workbook, workbookRelationships, order}={}) {
    const parts = new Map([
        ['_rels/.rels', relationships([{id:'root',kind:'officeDocument',target:'xl/workbook.xml'}])],
        ['xl/workbook.xml', workbook ?? workbookXml(sheets, date1904)],
        ['xl/_rels/workbook.xml.rels', workbookRelationships ?? relationships([
            {id:'rId1',kind:'worksheet',target:'worksheets/sheet1.xml'},
            ...(style === null ? [] : [{id:'styles',kind:'styles',target:'styles.xml'}]),
            ...(theme ? [{id:'theme',kind:'theme',target:'theme/theme1.xml'}] : []),
            ...(shared ? [{id:'strings',kind:'sharedStrings',target:'sharedStrings.xml'}] : []),
        ])],
        ['xl/worksheets/sheet1.xml', sheet],
        ...(style === null ? [] : [['xl/styles.xml', style]]),
        ...(theme ? [['xl/theme/theme1.xml',theme]] : []),
        ...(shared ? [['xl/sharedStrings.xml',shared]] : []),
        ...(drawing ? [
            ['xl/drawings/drawing1.xml',drawing],
            ['xl/worksheets/_rels/sheet1.xml.rels',relationships([{id:'drawingRel',kind:'drawing',target:'../drawings/drawing1.xml'}])],
            ['xl/drawings/_rels/drawing1.xml.rels',drawingRelationships ?? relationships([])],
        ] : []),
        ...extra,
    ]);
    const ordered = order ? [...order.map(path => [path,parts.get(path)]), ...[...parts].filter(([path])=>!order.includes(path))] : [...parts];
    return zip(ordered);
}
export function centralEntries(bytes) {
    const view = new DataView(bytes), decoder = new TextDecoder();
    const entries = [];
    for (let p=0;p<bytes.byteLength-46;p++) if(view.getUint32(p,true)===0x02014b50) {
        const len=view.getUint16(p+28,true);
        entries.push({offset:p,name:decoder.decode(new Uint8Array(bytes,p+46,len)),local:view.getUint32(p+42,true)});
        p += 45+len+view.getUint16(p+30,true)+view.getUint16(p+32,true);
    }
    return entries;
}
