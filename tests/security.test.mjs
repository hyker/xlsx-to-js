import test from 'node:test';
import assert from 'node:assert/strict';
import { JSDOM, VirtualConsole } from 'jsdom';
import { transformSync } from 'esbuild';
import { readFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { XlsxParser, DEFAULT_LIMITS, PREVIEW_LIMITS, getRangeDetails } from '../dist/index.js';
import { NS, REL, DRAW, esc, fixture, zip, worksheetXml, stylesXml, themeXml, workbookXml, relationships, drawingXml, gifWithComment, pngImage, centralEntries } from './fixtures.mjs';

const dom = new JSDOM('<main></main>', {runScripts:'dangerously',virtualConsole:new VirtualConsole()});
if (process.env.XLSX_XML_PARSER === 'xmldom') await import('../dist/worker.js');
else globalThis.DOMParser = dom.window.DOMParser;
const parser = new XlsxParser();
const read = async (options={}, parserOptions={styles:true}) => parser.readFile(await fixture(options),parserOptions);
function inspect(html) {
    const root = dom.window.document.querySelector('main'); root.innerHTML=html;
    for(const el of root.querySelectorAll('*')) for(const attr of el.attributes) {
        assert.ok(!/^on/i.test(attr.name),`Unexpected event handler: ${attr.name}`);
        if (['src','href'].includes(attr.name)) assert.ok(attr.value.startsWith('data:image/'));
    }
    assert.equal(root.querySelectorAll('script,iframe,object,embed,link').length,0);
    return root;
}
const both = w => [parser.toHTML(w),parser.toHTMLSheet(w,0)];

test('tiny archives with repeated shared strings fail before reporting completion', async () => {
    const shared = `<sst xmlns="${NS}"><si><t>${'x'.repeat(32767)}</t></si></sst>`;
    const data = Array.from({ length: 1000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>0</v></c></row>`).join('');
    const file = await fixture({ shared, sheet: worksheetXml(data, 'A1:A1000') });
    assert.ok(file.byteLength < 10000);
    for (const limits of [DEFAULT_LIMITS, PREVIEW_LIMITS]) {
        const phases = [];
        await assert.rejects(parser.readFile(file, { limits, onProgress: p => phases.push(p.phase) }), /Workbook text/);
        assert.ok(!phases.includes('complete'));
    }
    const w = await parser.readFile(file, { limits: { maxWorkbookTextLength: 34 * 1024 * 1024 } });
    assert.equal(w.workSheets[0].data[999][0].value.length, 32767);
});

test('workbook text budget includes formulas and aggregates across hidden sheets', async () => {
    const sheet = worksheetXml(`<row r="1"><c r="A1" t="inlineStr"><is><t>${'x'.repeat(2000)}</t></is><f>${'y'.repeat(2000)}</f></c></row>`);
    await read({ sheet }, { limits: { maxWorkbookTextLength: 5000 } });
    await assert.rejects(read({ sheet }, { limits: { maxWorkbookTextLength: 3000 } }), /Workbook text/);
    await assert.rejects(read({
        sheet,
        sheets: [{ name: 'First', id: 1, rel: 'rId1' }, { name: 'Second', id: 2, rel: 'rId2', state: 'hidden' }],
        workbookRelationships: relationships([
            { id: 'rId1', kind: 'worksheet', target: 'worksheets/sheet1.xml' },
            { id: 'rId2', kind: 'worksheet', target: 'worksheets/sheet2.xml' },
        ]),
        extra: [['xl/worksheets/sheet2.xml', sheet]],
    }, { limits: { maxWorkbookTextLength: 5000 } }), /Workbook text/);
});

test('workbook text limits use UTF-16 lengths, allow the exact boundary and ignore omitted cells', async () => {
    const sheet = worksheetXml('<row r="1"><c r="A1" t="inlineStr"><is><t>😀</t></is><f>X</f></c></row>');
    // First + rId1 + visible + A1:A1 + A1 + emoji + formula = 26 UTF-16 units.
    const w = await read({ sheet, style: null }, { limits: { maxWorkbookTextLength: 26 } });
    assert.equal(w.workSheets[0].data[0][0].value, '😀');
    await assert.rejects(read({ sheet, style: null }, { limits: { maxWorkbookTextLength: 25 } }), /Workbook text/);
    await read({ sheet: sheet.replace('<row r="1">', '<row r="1" hidden="1">'), style: null },
        { skipHiddenRows: true, limits: { maxWorkbookTextLength: 21 } });
    for (const value of [0, -1, Infinity, NaN, 1.5, undefined]) {
        await assert.rejects(read({}, { limits: { maxWorkbookTextLength: value } }), /Invalid resource limit/);
    }
});

test('workbook text budget counts shared styles once but bounds repeated image strings', async () => {
    const w = await read({
        sheet: worksheetXml('<row r="1"><c r="A1" s="0"><v>1</v></c><c r="B1" s="0"><v>2</v></c></row>', 'A1:B1'),
        style: stylesXml({ font: 'x'.repeat(2000) }),
    }, { styles: true, limits: { maxWorkbookTextLength: 5000 } });
    assert.equal(w.workSheets[0].data[0][0].style, w.workSheets[0].data[0][1].style);
    const drawing = drawingXml({ image: true });
    const anchor = drawing.slice(drawing.indexOf('<xdr:oneCellAnchor>'), drawing.indexOf('</xdr:wsDr>'));
    await assert.rejects(read({
        sheet: worksheetXml(undefined, 'A1', '<drawing r:id="drawingRel"/>'),
        drawing: drawing.replace(anchor, anchor.repeat(10)),
        drawingRelationships: relationships([{ id: 'imageRel', kind: 'image', target: '../media/image.gif' }]),
        extra: [['xl/media/image.gif', gifWithComment(1000)]],
    }, { drawings: true, limits: { maxWorkbookTextLength: 5000 } }), /Workbook text/);
});

test('multi-megabyte embedded GIF renders in full and paginated HTML without regex stack overflow', async () => {
    const w = await read({
        sheet: worksheetXml(undefined, 'A1', '<drawing r:id="drawingRel"/>'),
        drawing: drawingXml({ image: true }),
        drawingRelationships: relationships([{ id: 'imageRel', kind: 'image', target: '../media/image.gif' }]),
        extra: [['xl/media/image.gif', gifWithComment()]],
    }, { drawings: true });
    for (const html of [...both(w), parser.toHTMLSheetPage(w, 0).html]) {
        assert.ok(html.includes(`src="data:image/gif;base64,${w.workSheets[0].drawings[0].base64}"`));
    }
    const image = w.workSheets[0].drawings[0];
    image.base64 = image.base64.slice(0, -4) + 'AA=A';
    assert.throws(() => parser.toHTML(w), /Invalid image data/);
});

test('image validation preserves legal padding and rejects malformed or oversized data', async () => {
    const w = await read({
        sheet: worksheetXml(undefined, 'A1', '<drawing r:id="drawingRel"/>'),
        drawing: drawingXml({ image: true }),
        drawingRelationships: relationships([{ id: 'imageRel', kind: 'image', target: '../media/image.gif' }]),
        extra: [['xl/media/image.gif', gifWithComment(0)]],
    }, { drawings: true });
    const image = w.workSheets[0].drawings[0];
    for (const padding of [0, 1, 2]) {
        const bytes = gifWithComment(padding), size = bytes.length;
        image.base64 = bytes.toString('base64');
        assert.ok(parser.toHTML(w, { limits: { maxEntryBytes: size } }).includes('data:image/gif;base64,'));
        assert.throws(() => parser.toHTML(w, { limits: { maxEntryBytes: size - 1 } }), /Invalid image data/);
    }
    for (const suffix of ['=', '===', 'AA=A', 'A===', 'AAAA\n', 'AAA-', 'AAA_', 'AAAé']) {
        image.base64 = 'R0lGODlh' + suffix;
        assert.throws(() => parser.toHTMLSheetPage(w, 0), /Invalid image data/);
    }
});

test('font and alignment payloads cannot escape HTML or add CSS declarations',async()=>{
    for(const payload of [
        'Arial;"><img src=x onerror="window.pwned=1"><span style="',
        'Arial;position:fixed;inset:0;background-image:url(https://example.invalid/leak)',
        'Arial\\";background:url(https://example.invalid/leak)',
        'Arial\n;position:fixed',
    ]) {
        const w=await read({style:stylesXml({font:payload,horizontal:payload,vertical:payload})});
        for(const html of both(w)) {
            const root=inspect(html), td=root.querySelector('td');
            const computed=dom.window.getComputedStyle(td);
            assert.equal(root.querySelectorAll('img').length,0);
            assert.equal(dom.window.getComputedStyle(td).position,'static');assert.equal(dom.window.getComputedStyle(td).backgroundImage,'none');
            assert.equal(computed.textAlign,'');assert.equal(computed.verticalAlign,'bottom');
        }
    }
});

test('cell values, formulas, sheet names and drawing text remain inert',async()=>{
    const payload='<img src=x onerror="window.pwned=1">';
    const w=await read({sheet:worksheetXml(`<row r="1"><c r="A1" t="inlineStr"><is><t>${esc(payload)}</t></is><f>${esc(payload)}</f></c></row>`)});
    w.workSheets[0].name=payload;
    for(const html of both(w)) {const root=inspect(html);assert.equal(root.querySelector('td').textContent,payload);assert.equal(root.querySelectorAll('img').length,0);}
    const d=await read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({text:payload})},{drawings:true});
    for(const html of both(d)) {const root=inspect(html);assert.equal(root.querySelector('.xl-textbox').textContent,payload);}
});

test('drawing color payload is rejected even when cell styles are disabled',async()=>{
    const color='000000;"><img src=x onerror="window.pwned=1"><span style="';
    await assert.rejects(read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({color})},{drawings:true}),/Invalid drawing color/);
});

test('renderers revalidate mutable workbook colors, image data and geometry',async()=>{
    const w=await read();
    const style=w.workSheets[0].data[0][0].style;
    style.fontColor='red;position:fixed';style.fgColor='red;"><img onerror=x>';
    style.border.top={style:'thin',color:'red;background:url(https://example.invalid)'};
    for(const html of both(w)){const td=inspect(html).querySelector('td');assert.equal(dom.window.getComputedStyle(td).position,'static');assert.equal(dom.window.getComputedStyle(td).backgroundImage,'none');}
    const drawing={type:'shape',position:{from:{col:0,row:0,colOff:0,rowOff:0},to:{col:0,row:0,colOff:0,rowOff:0}},sizeEMU:{cx:9525,cy:9525},properties:{fillColor:'red;"><img onerror=x>',lineColor:'red;position:fixed'}};
    w.workSheets[0].drawings=[drawing];for(const html of both(w))inspect(html);
    drawing.sizeEMU.cx=Infinity;for(const fn of [()=>parser.toHTML(w),()=>parser.toHTMLSheet(w,0)])assert.throws(fn,/drawing geometry/);
    drawing.sizeEMU.cx=9525;drawing.type='image';drawing.base64='iVBORw0KGgo=" onerror="window.pwned=1';
    assert.throws(()=>parser.toHTML(w),/Invalid image data/);
});

for(const dimension of ['A0','A-1','A1:','A1:A0','B2:A1','A1:B2:C3','A1:Infinity','a1','A1junk']) {
    test(`dimension rejected before allocation: ${dimension}`,async()=>{
        for(const dense of [true,false]) await assert.rejects(read({sheet:worksheetXml('',dimension)},{dense}),/reference|range|budget|limits/i);
    });
}

test('tiny sparse and dense allocation attacks are rejected in a killable subprocess',()=>{
    for(const dimension of ['A1:ZZZZZZZZZZZZ1','A1:XFD1048576']) for(const dense of ['true','false']) {
        const child=spawnSync(process.execPath,['--max-old-space-size=128',fileURLToPath(new URL('./resource-child.mjs',import.meta.url)),dimension,dense],{timeout:5000,encoding:'utf8'});
        assert.equal(child.error,undefined,`Resource regression did not terminate: ${child.error}`);
        assert.equal(child.status,0,child.stderr);
    }
});

test('grid and merge budgets apply across sheets and to mutable renderer input',async()=>{
    await assert.rejects(read({sheet:worksheetXml('', 'A1:B2')},{limits:{maxCells:3}}),/budget/);
    const w=await read();w.workSheets.push(structuredClone(w.workSheets[0]));
    assert.throws(()=>parser.toHTML(w,{limits:{maxCells:1}}),/budget/);
    w.workSheets[0].mergeCells=['A1:A10000'];
    assert.throws(()=>parser.toHTMLSheet(w,0,{limits:{maxMergedCells:2}}),/Merged cells/);
    await assert.rejects(read({sheet:worksheetXml('', 'A1:B2','<mergeCells><mergeCell ref="A1:B2"/></mergeCells>')},{limits:{maxMergedCells:3}}),/Merged cells/);
});

test('invalid, out-of-dimension and duplicate cells fail explicitly',async()=>{
    for(const data of [
        '<row r="2"><c r="A2"><v>1</v></c></row>',
        '<row r="1"><c r="A2"><v>1</v></c></row>',
        '<row r="1"><c r="A1"/><c r="A1"/></row>',
        '<row r="1"/><row r="1"/>',
        '<row r="1"><c r="A1" s="999"/></row>',
        '<row r="1"><c r="A1" t="s"><v>99</v></c></row>',
        '<row r="1"><c r="A1" t="n"><v>Infinity</v></c></row>',
        `<row r="1"><c r="${esc('<img onerror=x>')}"/></row>`,
    ]) await assert.rejects(read({sheet:worksheetXml(data)}),e=>e instanceof Error && !(e instanceof TypeError));
});

test('malformed XML, wrong roots, DTDs and entity declarations are rejected',async()=>{
    for(const sheet of ['<worksheet><sheetData>','<wrong/>',`<!DOCTYPE worksheet [<!ENTITY x "boom">]><worksheet xmlns="${NS}"><sheetData/></worksheet>`,`<worksheet xmlns="${NS}"/>`]) await assert.rejects(read({sheet}),/XML|DTD|sheetData/);
    await assert.rejects(read({style:'<styleSheet><fonts>'}),/XML/);
    await assert.rejects(read({shared:'<sst><si>',sheet:worksheetXml()}),/XML/);
});

test('XML node and depth budgets are enforced',async()=>{
    await assert.rejects(read({sheet:worksheetXml('')},{limits:{maxXmlNodes:2}}),/complexity/);
    await assert.rejects(read({sheet:worksheetXml('<a>'.repeat(80)+'</a>'.repeat(80))}),/complexity/);
});

test('alternate namespace prefixes work and foreign namespaces cannot spoof workbook elements',async()=>{
    const xml=workbookXml().replace('<workbook ','<s:workbook ').replace('xmlns="','xmlns:s="').replace('</workbook>','</s:workbook>').replaceAll('<sheets>','<s:sheets>').replaceAll('</sheets>','</s:sheets>').replace('<sheet ','<s:sheet ').replace('<workbookPr ','<s:workbookPr ');
    const w=await read({workbook:xml});assert.equal(w.workSheets[0].data[0][0].value,'42');
    await assert.rejects(read({workbook:xml.replace(NS,'https://example.invalid/spoof')}),/XML/);
});

test('archive file size, entry count, per-entry size and total expansion have independent budgets',async()=>{
    const bytes=await fixture();
    await assert.rejects(parser.readFile(bytes,{limits:{maxFileBytes:100}}),/file budget/);
    await assert.rejects(parser.readFile(bytes,{limits:{maxEntries:1}}),/entry count/);
    const shared=`<sst xmlns="${NS}"><si><t>${'A'.repeat(1024*1024)}</t></si></sst>`;
    await assert.rejects(read({shared},{limits:{maxEntryBytes:1024}}),/expansion/);
    await assert.rejects(read({extra:[['unused/a.bin','A'.repeat(10000)],['unused/b.bin','B'.repeat(10000)]]},{limits:{maxEntryBytes:15000,maxTotalBytes:18000}}),/expansion/);
});

test('forged ZIP uncompressed size is caught during streamed decompression',async()=>{
    const bytes=await fixture({sheet:worksheetXml(' '.repeat(128*1024))});
    const entry=centralEntries(bytes).find(e=>e.name==='xl/worksheets/sheet1.xml');
    new DataView(bytes).setUint32(entry.offset+24,1,true);
    await assert.rejects(parser.readFile(bytes),/Decompressed data exceeds|size mismatch/);
});

test('corrupted ZIP checksum, truncated directory, traversal and duplicate paths fail',async()=>{
    let bytes=await fixture(), view=new DataView(bytes);
    const entry=centralEntries(bytes).find(e=>e.name==='xl/workbook.xml');view.setUint32(entry.offset+16,0,true);
    await assert.rejects(parser.readFile(bytes),/checksum/);
    bytes=await fixture();await assert.rejects(parser.readFile(bytes.slice(0,-1)),/ZIP/);
    await assert.rejects(parser.readFile(await zip([['../xl/workbook.xml',workbookXml()]])),/Unsafe/);
    bytes=await fixture({extra:[['xl/foo1.xml','a'],['xl/foo2.xml','b']]});view=new DataView(bytes);
    const duplicate=centralEntries(bytes).find(e=>e.name==='xl/foo2.xml');
    const name=new TextEncoder().encode('xl/foo1.xml');new Uint8Array(bytes,duplicate.offset+46,name.length).set(name);new Uint8Array(bytes,duplicate.local+30,name.length).set(name);
    await assert.rejects(parser.readFile(bytes),/duplicate/);
});

test('unsupported ZIP encryption and ZIP64 are rejected before loading',async()=>{
    const bytes=await fixture(), view=new DataView(bytes), entry=centralEntries(bytes)[0];
    view.setUint16(entry.offset+8,1,true);await assert.rejects(parser.readFile(bytes),/Unsupported/);
    view.setUint16(entry.offset+8,0,true);view.setUint32(entry.offset+24,0xffffffff,true);await assert.rejects(parser.readFile(bytes),/Unsupported/);
});

test('sheet order follows relationships, supports arbitrary part names and ignores decoy paths',async()=>{
    const sheets=[{name:'First',id:10,rel:'first'},{name:'Second',id:20,rel:'second'}];
    const w=await read({sheets, workbookRelationships:relationships([{id:'first',kind:'worksheet',target:'worksheets/custom.xml'},{id:'second',kind:'worksheet',target:'worksheets/sheet2.xml'}]),
        sheet:worksheetXml('<row r="1"><c r="A1"><v>999</v></c></row>'),
        extra:[['xl/worksheets/sheet2.xml',worksheetXml('<row r="1"><c r="A1"><v>222</v></c></row>')],['xl/worksheets/custom.xml',worksheetXml('<row r="1"><c r="A1"><v>111</v></c></row>')],['evil/xl/worksheets/sheet3.xml','bad']],
        order:['xl/worksheets/sheet2.xml','xl/worksheets/custom.xml']},{styles:false});
    assert.deepEqual(w.workSheets.map(s=>[s.name,s.data[0][0].value]),[['First','111'],['Second','222']]);
});

test('external, escaping, missing, ambiguous and wrong-type sheet relationships fail',async()=>{
    for(const rels of [
        [{id:'rId1',kind:'worksheet',target:'https://example.invalid/a.xlsx',external:true}],
        [{id:'rId1',kind:'worksheet',target:'../../outside.xml'}],
        [{id:'rId1',kind:'worksheet',target:'worksheets/missing.xml'}],
        [{id:'rId1',kind:'image',target:'worksheets/sheet1.xml'}],
        [{id:'rId1',kind:'worksheet',target:'worksheets/sheet1.xml'},{id:'rId1',kind:'worksheet',target:'worksheets/sheet1.xml'}],
        [],
    ]) await assert.rejects(read({workbookRelationships:relationships(rels)},{styles:false}),/relationship|package part|escapes/);
});

test('theme resolution is independent of ZIP order',async()=>{
    for(const order of [['xl/styles.xml','xl/theme/theme1.xml'],['xl/theme/theme1.xml','xl/styles.xml']]) {
        const w=await read({style:stylesXml({color:'theme'}),theme:themeXml,order});assert.equal(w.workSheets[0].data[0][0].style.fontColor,'#ffffff');
    }
});

test('hidden sheet state is preserved, HTML requires explicit opt-in',async()=>{
    const w=await read({sheets:[{name:'Secret',id:1,rel:'rId1',state:'veryHidden'}]});
    assert.equal(w.workSheets[0].state,'veryHidden');assert.equal(w.workSheets[0].data[0][0].value,'42');
    for(const html of both(w))assert.doesNotMatch(html,/>42<|Secret/);
    assert.match(parser.toHTML(w,{includeHiddenSheets:true}),/>42</);
    assert.match(parser.toHTMLSheet(w,0,{includeHiddenSheets:true}),/>42</);
});

test('hidden rows preserve coordinates and are omitted with skipHiddenRows',async()=>{
    const sheet=worksheetXml('<row r="1" hidden="1"><c r="A1"><v>99</v></c></row><row r="2"><c r="A2"><v>42</v></c></row>','A1:A2');
    const w=await read({sheet},{skipHiddenRows:true});assert.equal(w.workSheets[0].rowStyles[0].hidden,true);
    assert.equal(w.workSheets[0].data[1][0].value,'42');
    for(const html of both(w)){const root=inspect(html);assert.equal(root.querySelector('[data-row="1"]'),null);assert.equal(root.querySelector('[data-ref="A2"]').textContent,'42');}
    const full=await read({sheet},{skipHiddenRows:false});assert.match(parser.toHTML(full),/display:none/);
});

test('date systems and early 1900 serials are correct, explicit false font flags stay false',async()=>{
    const style=stylesXml({numFmtId:14,flags:'<b val="0"/><i val="false"/>'});
    for(const [date1904,year,month,day] of [[false,1900,1,11],[true,1904,1,12]]) {
        const w=await read({style,date1904});const cell=w.workSheets[0].data[0][0];
        assert.equal(cell.value,new Date(Date.UTC(year,month,day)).toLocaleDateString(undefined,{timeZone:'UTC'}));assert.equal(cell.style.bold,false);assert.equal(cell.style.italic,false);
    }
    const w=await read({style,sheet:worksheetXml('<row r="1"><c r="A1" s="0"><v>1</v></c></row>')});
    assert.equal(w.workSheets[0].data[0][0].value,new Date(Date.UTC(1900,0,1)).toLocaleDateString(undefined,{timeZone:'UTC'}));
});

test('sparse and dense modes render identically; dense empty objects are independent',async()=>{
    const options={sheet:worksheetXml('<row r="1"><c r="A1"><v>42</v></c></row>','A1:C2')};
    const sparse=await read(options,{dense:false}), dense=await read(options,{dense:true});
    assert.equal(parser.toHTML(sparse),parser.toHTML(dense));
    assert.notEqual(dense.workSheets[0].data[0][1],dense.workSheets[0].data[0][2]);
    assert.equal(getRangeDetails('A1').rowspan,1);assert.throws(()=>getRangeDetails('B2:A1'),/Reversed/);
});

test('all-sheet and single-sheet renderers preserve the same widths, heights, merges and drawings',async()=>{
    const w=await read({sheet:worksheetXml('<row r="1" ht="30"><c r="A1"><v>42</v></c></row>','A1:B2','<cols><col min="1" max="1" width="20"/></cols><mergeCells><mergeCell ref="A1:B1"/></mergeCells><drawing r:id="drawingRel"/>'),drawing:drawingXml()},{drawings:true});
    assert.equal(parser.toHTML(w),parser.toHTMLSheet(w,0));
    const root=inspect(parser.toHTML(w));assert.equal(root.querySelector('td').colSpan,2);assert.equal(root.querySelectorAll('.xl-shape').length,1);
    assert.match(root.querySelector('colgroup').innerHTML,/width:145px/);assert.match(root.querySelector('tbody tr').getAttribute('style'),/height:40px/);
});

for(const bad of ['Infinity','NaN','-1','1e999','1;position:fixed','999999999999999999999999999']) {
    test(`drawing geometry rejected: ${bad}`,async()=>{
        await assert.rejects(read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({extent:bad})},{drawings:true}),/geometry/);
    });
}

test('drawing count, coordinate, image relationships and expansion budgets are enforced',async()=>{
    await assert.rejects(read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({row:'10000'})},{drawings:true,limits:{maxRows:10000}}),/drawing row/);
    await assert.rejects(read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({image:true}),drawingRelationships:relationships([{id:'imageRel',kind:'image',target:'https://example.invalid/p.png',external:true}])},{drawings:true}),/External/);
    const w=await read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({extent:'9525000'})},{drawings:true});
    assert.throws(()=>parser.toHTML(w,{limits:{maxColumns:2}}),/expansion/);
    w.workSheets[0].drawings.push(structuredClone(w.workSheets[0].drawings[0]));assert.throws(()=>parser.toHTML(w,{limits:{maxDrawings:1}}),/Drawing count/);
});

test('PNG images use only embedded data URLs; drawing/media matching uses complete paths',async()=>{
    const png=pngImage();
    const w=await read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({image:true}),drawingRelationships:relationships([{id:'imageRel',kind:'image',target:'../media/a/image.png'}]),extra:[['xl/media/a/image.png',png],['xl/media/b/image.png','bad']]},{drawings:true});
    for(const html of both(w)){const img=inspect(html).querySelector('img');assert.match(img.src,/^data:image\/png;base64,/);}
});

test('HTML output budget and invalid limit overrides fail explicitly',async()=>{
    const w=await read();assert.throws(()=>parser.toHTML(w,{limits:{maxHtmlLength:100}}),/output budget/);
    for(const value of [0,-1,Infinity,NaN,1.5,undefined])await assert.rejects(read({}, {limits:{maxCells:value}}),/Invalid resource limit/);
    assert.ok(Object.isFrozen(DEFAULT_LIMITS));
});

test('demo status and error helpers display filenames and parser messages as text',async()=>{
    const source=await readFile(new URL('../examples/vite/src/storybook/status.ts',import.meta.url),'utf8');
    const code=transformSync(source,{loader:'ts',format:'esm'}).code;
    const {setStatus,setError}=await import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`);
    globalThis.document=dom.window.document;
    try {
        const payload='<img src=x onerror="window.pwned=1">';const root=document.querySelector('main');
        setStatus(root,payload,payload);assert.equal(root.querySelectorAll('img').length,0);assert.equal(root.querySelector('span').textContent,payload);
        setError(root,payload);assert.equal(root.querySelectorAll('img').length,0);assert.equal(root.querySelector('span').textContent,payload);
    } finally {delete globalThis.document;}
});

test('overlapping merges are rejected during parsing as well as rendering',async()=>{
    await assert.rejects(read({sheet:worksheetXml('', 'A1:C3','<mergeCells><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/></mergeCells>')}),/Overlapping/);
    const w=await read({sheet:worksheetXml('', 'A1:C3')});w.workSheets[0].mergeCells=['A1:B2','B2:C3'];
    for(const fn of [()=>parser.toHTML(w),()=>parser.toHTMLSheet(w,0)])assert.throws(fn,/Overlapping/);
});

test('missing dimensions are inferred across namespaces without changing cell positions',async()=>{
    const sheet=worksheetXml('<row r="2"><c r="B2"><v>42</v></c></row>',null);
    const w=await read({sheet});assert.equal(w.workSheets[0].dimention,'A1:B2');assert.equal(w.workSheets[0].data[1][1].value,'42');
});

test('workbook-wide parse grid, merge, drawing and sheet budgets are enforced',async()=>{
    const sheets=[{name:'First',id:1,rel:'rId1'},{name:'Second',id:2,rel:'rId2'}];
    const rels=relationships([{id:'rId1',kind:'worksheet',target:'worksheets/sheet1.xml'},{id:'rId2',kind:'worksheet',target:'worksheets/sheet2.xml'}]);
    const base={sheets,workbookRelationships:rels,sheet:worksheetXml('<row r="1"><c r="A1"><v>42</v></c></row>','A1:A2'),extra:[['xl/worksheets/sheet2.xml',worksheetXml('<row r="1"><c r="A1"><v>42</v></c></row>','A1:A2')]]};
    await assert.rejects(read(base,{limits:{maxCells:3}}),/Workbook grid/);
    await assert.rejects(read(base,{limits:{maxSheets:1}}),/Sheet count/);
    const merged=worksheetXml('', 'A1:A2','<mergeCells><mergeCell ref="A1:A2"/></mergeCells>');
    await assert.rejects(read({...base,sheet:merged,extra:[['xl/worksheets/sheet2.xml',merged]]},{limits:{maxMergedCells:3}}),e=>e.name==='XlsxLimitError'&&e.limit==='maxMergedCells'&&/Workbook merged cells/.test(e.message));
    const drawn=worksheetXml('<row r="1"><c r="A1"><v>42</v></c></row>','A1','<drawing r:id="drawingRel"/>');
    await assert.rejects(read({...base,sheet:drawn,drawing:drawingXml(),extra:[['xl/worksheets/sheet2.xml',drawn],['xl/worksheets/_rels/sheet2.xml.rels',relationships([{id:'drawingRel',kind:'drawing',target:'../drawings/drawing1.xml'}])]]},{drawings:true,limits:{maxDrawings:1}}),/Workbook drawings/);
});

test('default-hidden rows retain coordinates and zero-width columns remain hidden',async()=>{
    const sheet=worksheetXml('<row r="2" hidden="0"><c r="B2"><v>42</v></c></row>','A1:B2','<sheetFormatPr zeroHeight="1" defaultRowHeight="15"/><cols><col min="1" max="1" width="10" hidden="1"/></cols>');
    const w=await read({sheet},{skipHiddenRows:true});const root=inspect(parser.toHTML(w));
    assert.equal(root.querySelector('[data-row="1"]'),null);assert.equal(root.querySelector('[data-row="2"]').textContent,'2');
    assert.equal(root.querySelector('[data-col="1"]').style.display,'none');
});

test('shared rich strings do not shift indexes or include phonetic annotations',async()=>{
    const shared=`<sst xmlns="${NS}"><si><r><t>one</t></r><r><t>two</t></r><rPh><t>phonetic</t></rPh></si><si><t>three</t></si></sst>`;
    const sheet=worksheetXml('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>','A1:B1');
    const w=await read({shared,sheet});assert.deepEqual(w.workSheets[0].data[0].map(c=>c.value),['onetwo','three']);
});

test('package root relationship supports a workbook at a nonstandard path',async()=>{
    const bytes=await zip([
        ['_rels/.rels',relationships([{id:'root',kind:'officeDocument',target:'custom/book.xml'}])],
        ['custom/book.xml',workbookXml()],
        ['custom/_rels/book.xml.rels',relationships([{id:'rId1',kind:'worksheet',target:'../sheets/data.xml'}])],
        ['sheets/data.xml',worksheetXml('<row r="1"><c r="A1"><v>42</v></c></row>')],
    ]);
    const w=await parser.readFile(bytes);assert.equal(w.workSheets[0].data[0][0].value,'42');
});

test('SVG and arbitrary binary image content cannot be rendered as active documents',async()=>{
    await assert.rejects(read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({image:true}),drawingRelationships:relationships([{id:'imageRel',kind:'image',target:'../media/image.svg'}]),extra:[['xl/media/image.svg','<svg xmlns="http://www.w3.org/2000/svg" onload="alert(1)"/>']]},{drawings:true}),/Unsupported image format/);
    const w=await read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({image:true}),drawingRelationships:relationships([{id:'imageRel',kind:'image',target:'../media/image.gif'}]),extra:[['xl/media/image.gif',gifWithComment(0)]]},{drawings:true});
    for(const data of ['<svg/>','not an image']) {
        w.workSheets[0].drawings[0].base64=btoa(data);
        assert.throws(()=>parser.toHTML(w),/Unsupported image format/);assert.throws(()=>parser.toHTMLSheet(w,0),/Unsupported image format/);
    }
});

test('font sizes, style references and drawing line widths are validated',async()=>{
    await assert.rejects(read({style:stylesXml({fontSize:'1e999'})}),/font size/);
    await assert.rejects(read({style:stylesXml().replace('fontId="0"','fontId="999"')}),/fontId/);
    const drawing=drawingXml().replace('</xdr:spPr>','<a:ln w="Infinity"/></xdr:spPr>');
    await assert.rejects(read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing},{drawings:true}),/line width/);
    await assert.rejects(read({}, {limits:JSON.parse('{"__proto__":1}')}),/Unknown resource limit/);
});

test('the bundled demo workbook parses and renders with all supported options enabled',async()=>{
    const code=transformSync(await readFile(new URL('../examples/vite/src/storybook/sampleWorkbook.ts',import.meta.url),'utf8'),{loader:'ts',format:'esm'}).code;
    const {sampleWorkbookBase64}=await import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`);
    const bytes=Uint8Array.from(Buffer.from(sampleWorkbookBase64,'base64'));
    const w=await parser.readFile(bytes.buffer,{styles:true,drawings:true,dense:true,skipHiddenRows:true});
    assert.deepEqual(w.workSheets.map(s=>s.name),['Overview','Dataset']);
    assert.ok(inspect(parser.toHTML(w)).querySelectorAll('td').length>0);
});

test('a missing or external package root cannot select a workbook implicitly',async()=>{
    await assert.rejects(parser.readFile(await zip([['xl/workbook.xml',workbookXml()]])),/workbook relationship/);
    await assert.rejects(read({extra:[['_rels/.rels',relationships([{id:'root',kind:'officeDocument',target:'https://example.invalid/book.xml',external:true}])]]}),/workbook relationship/);
    assert.deepEqual(getRangeDetails('A1:XFD1048576'),{origin:'A1',colspan:16384,rowspan:1048576});
});

test('XML complexity is rejected before native DOM allocation, including attribute floods', async () => {
    const Native = globalThis.DOMParser;
    let worksheetParses = 0;
    globalThis.DOMParser = class {
        parseFromString(xml, type) {
            if (xml.includes('<worksheet')) worksheetParses++;
            return new Native().parseFromString(xml, type);
        }
    };
    try {
        for (const [data, limits] of [
            ['<a/>'.repeat(41), { maxXmlNodes: 40 }],
            ['<a>'.repeat(65) + '</a>'.repeat(65), {}],
            [`<a ${Array.from({ length: 257 }, (_, i) => `p${i}="x"`).join(' ')}/>`, {}],
            ['<a p="x"/>'.repeat(30), { maxXmlAttributes: 20 }],
            ['<![CDATA[x]]>'.repeat(41), { maxXmlNodes: 40 }],
            ['<!-- x -->'.repeat(41), { maxXmlNodes: 40 }],
            ['<?work x?>'.repeat(41), { maxXmlNodes: 40 }],
        ]) {
            worksheetParses = 0;
            await assert.rejects(read({ sheet: worksheetXml(data) }, { limits }), /complexity/);
            assert.equal(worksheetParses, 0);
        }
    } finally { globalThis.DOMParser = Native; }
});

test('XML preflight respects quoted attributes, CDATA, comments and processing instructions', async () => {
    const sheet = worksheetXml('<row r="1" note="&lt;/row> = &quot; &apos;"><c r="A1" t="inlineStr"><is><t><![CDATA[<a attr="b">]]></t></is></c></row>');
    const w = await read({ sheet: sheet.replace('<sheetData>', '<?work <ignored>?><!-- <ignored/> --><sheetData>') });
    assert.equal(w.workSheets[0].data[0][0].value, '<a attr="b">');
});

const imageWorkbook = (bytes, drawing = drawingXml({ image: true }), extra = {}) => ({
    sheet: worksheetXml(undefined, 'A1', '<drawing r:id="drawingRel"/>'), drawing,
    drawingRelationships: relationships([{ id: 'imageRel', kind: 'image', target: '../media/image.bin' }]),
    extra: [['xl/media/image.bin', bytes]], ...extra,
});

test('a tiny archive with a 4096-square PNG is rejected before completion or rendering', async () => {
    const bytes = await fixture(imageWorkbook(pngImage(4096, 4096)));
    assert.ok(bytes.byteLength < 5000);
    const phases = [];
    await assert.rejects(parser.readFile(bytes, { drawings: true, onProgress: p => phases.push(p.phase) }), /Image pixels/);
    assert.ok(!phases.includes('complete'));
    const w = await read(imageWorkbook(pngImage()), { drawings: true });
    // Mutation cannot bypass the render-side checks, even off the selected page.
    w.workSheets[0].dimention = 'A1:CV1';
    w.workSheets[0].drawings[0].position.from.col = 80;
    w.workSheets[0].drawings[0].position.to.col = 80;
    w.workSheets[0].drawings[0].base64 = pngImage(4096, 4096).toString('base64');
    for (const render of [() => parser.toHTML(w), () => parser.toHTMLSheet(w, 0), () => parser.toHTMLSheetPage(w, 0)]) assert.throws(render, /Image pixels/);
});

test('per-image dimensions and total pixels count repeated media and shared sheets', async () => {
    const drawing = drawingXml({ image: true });
    const anchor = drawing.slice(drawing.indexOf('<xdr:oneCellAnchor>'), drawing.indexOf('</xdr:wsDr>'));
    const parts = imageWorkbook(pngImage(8, 8), drawing.replace(anchor, anchor.repeat(2)));
    await assert.rejects(read(parts, { drawings: true, limits: { maxImagePixels: 63 } }), /Image pixels/);
    await assert.rejects(read(parts, { drawings: true, limits: { maxImageDimension: 7 } }), /Image pixels/);
    await assert.rejects(read(parts, { drawings: true, limits: { maxTotalImagePixels: 127 } }), /Total image pixels/);
    const w = await read(parts, { drawings: true, limits: { maxImagePixels: 64, maxImageDimension: 8, maxTotalImagePixels: 128 } });
    for (const render of [() => parser.toHTML(w, { limits: { maxTotalImagePixels: 127 } }),
        () => parser.toHTMLSheetPage(w, 0, { limits: { maxTotalImagePixels: 127 } })]) assert.throws(render, /Total image pixels/);
    await assert.rejects(read({ ...parts, sheets: [{ name: 'A', id: 1, rel: 'rId1' }, { name: 'B', id: 2, rel: 'rId2' }],
        workbookRelationships: relationships([{ id: 'rId1', kind: 'worksheet', target: 'worksheets/sheet1.xml' },
            { id: 'rId2', kind: 'worksheet', target: 'worksheets/sheet2.xml' }, { id: 'styles', kind: 'styles', target: 'styles.xml' }]),
        extra: [...parts.extra, ['xl/worksheets/sheet2.xml', parts.sheet], ['xl/worksheets/_rels/sheet2.xml.rels',
            relationships([{ id: 'drawingRel', kind: 'drawing', target: '../drawings/drawing1.xml' }])]],
    }, { drawings: true, limits: { maxTotalImagePixels: 255 } }), /Total image pixels/);
    w.workSheets[0].drawings[0].base64 = pngImage(0, 1).toString('base64');
    assert.throws(() => parser.toHTML(w), /Invalid image/);
});

test('GIF logical and frame dimensions are bounded; GIF and PNG animation are rejected', async () => {
    const gif = Buffer.from('R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw==', 'base64');
    const huge = Buffer.from(gif); huge.writeUInt16LE(65535, 6); huge.writeUInt16LE(65535, 8);
    await assert.rejects(read(imageWorkbook(huge), { drawings: true }), /Image pixels/);
    const frame = gif.indexOf(0x2c), outside = Buffer.from(gif); outside.writeUInt16LE(2, frame + 5);
    await assert.rejects(read(imageWorkbook(outside), { drawings: true }), /Invalid image/);
    const animated = Buffer.concat([gif.subarray(0, -1), gif.subarray(frame, -1), Buffer.from([0x3b])]);
    await assert.rejects(read(imageWorkbook(animated), { drawings: true }), /Animated images/);
    const png = pngImage(), actl = Buffer.alloc(20); actl.writeUInt32BE(8); actl.write('acTL', 4);
    await assert.rejects(read(imageWorkbook(Buffer.concat([png.subarray(0, 33), actl, png.subarray(33)])), { drawings: true }), /Animated images/);
    for (const malformed of [gif.subarray(0, 12), gif.subarray(0, -2), png.subarray(0, 32), png.subarray(0, -5)]) {
        await assert.rejects(read(imageWorkbook(malformed), { drawings: true }), /Invalid image/);
    }
});

test('JPEG dimensions are bounded across frame headers and cannot be redefined in a scan', async () => {
    const jpeg = (width, height, marker = 0xc0, tail = []) => Buffer.from([
        0xff, 0xd8, 0xff, marker, 0, 11, 8, height >> 8, height & 255, width >> 8, width & 255, 1, 1, 0x11, 0,
        0xff, 0xda, 0, 8, 1, 1, 0, 0, 63, 0, ...tail, 0xff, 0xd9,
    ]);
    for (const marker of [0xc0, 0xc1, 0xc2]) {
        const w = await read(imageWorkbook(jpeg(8, 8, marker)), { drawings: true });
        assert.match(parser.toHTML(w), /data:image\/jpeg/);
        await assert.rejects(read(imageWorkbook(jpeg(4096, 4096, marker)), { drawings: true }), /Image pixels/);
    }
    for (const bytes of [jpeg(0, 8), jpeg(8, 8, 0xc3), jpeg(8, 8, 0xc0, [0xff, 0xdc, 0, 4, 0xff, 0xff]),
        jpeg(8, 8).subarray(0, -1), Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0xff, 0xff])]) {
        await assert.rejects(read(imageWorkbook(bytes), { drawings: true }), /Invalid image/);
    }
});
