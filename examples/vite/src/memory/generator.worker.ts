// Builds benchmark workbooks off the main thread and transfers the bytes back,
// so generation memory is released before the parser worker starts.
import JSZip from 'jszip';

export type Shape = 'numbers' | 'strings' | 'formulas' | 'worst-case';
export interface GenerateRequest { shape: Shape; cells: number }

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const COLUMNS = 10;

const column = (n: number) => { let s = ''; for (; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
const rels = (items: [string, string, string][]) =>
    `<Relationships xmlns="${PKG}">${items.map(([id, kind, target]) => `<Relationship Id="${id}" Type="${REL}/${kind}" Target="${target}"/>`).join('')}</Relationships>`;

function sheetXml({ shape, cells }: GenerateRequest): string {
    if (shape === 'worst-case') {
        // Just under the default 400,000 node and attribute budgets: the
        // heaviest DOM a file can produce while still passing preflight.
        return `<worksheet xmlns="${NS}"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><x>${'<y a="1"/>'.repeat(399_980)}</x></worksheet>`;
    }
    const rows = Math.ceil(cells / COLUMNS), parts: string[] = [];
    for (let r = 1; r <= rows; r++) {
        let row = `<row r="${r}">`;
        for (let c = 1; c <= COLUMNS; c++) {
            const ref = `${column(c)}${r}`;
            row += shape === 'numbers' ? `<c r="${ref}" s="1"><v>${r * c}.25</v></c>`
                : shape === 'strings' ? `<c r="${ref}" s="0" t="s"><v>${(r * COLUMNS + c) % 5000}</v></c>`
                : `<c r="${ref}"><f>A${r}*2</f><v>${r}</v></c>`;
        }
        parts.push(row + '</row>');
    }
    return `<worksheet xmlns="${NS}"><dimension ref="A1:${column(COLUMNS)}${rows}"/><sheetData>${parts.join('')}</sheetData></worksheet>`;
}

self.onmessage = async (event: MessageEvent<GenerateRequest>) => {
    const zip = new JSZip();
    zip.file('_rels/.rels', rels([['root', 'officeDocument', 'xl/workbook.xml']]));
    zip.file('xl/workbook.xml', `<workbook xmlns="${NS}" xmlns:r="${REL}"><sheets><sheet name="Benchmark" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', rels([['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'styles', 'styles.xml'], ['rId3', 'sharedStrings', 'sharedStrings.xml']]));
    zip.file('xl/styles.xml', `<styleSheet xmlns="${NS}"><fonts><font><name val="Calibri"/><sz val="11"/></font></fonts><fills><fill><patternFill patternType="none"/></fill></fills><borders><border/></borders><cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="4" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`);
    zip.file('xl/sharedStrings.xml', `<sst xmlns="${NS}">${Array.from({ length: 5000 }, (_, i) => `<si><t>Customer name ${i}</t></si>`).join('')}</sst>`);
    zip.file('xl/worksheets/sheet1.xml', sheetXml(event.data));
    const bytes = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
    (self as unknown as Worker).postMessage(bytes, [bytes]);
};
