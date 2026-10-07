import test from 'node:test';
import assert from 'node:assert/strict';
import { JSDOM, VirtualConsole } from 'jsdom';
import { transformSync } from 'esbuild';
import { readFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { XlsxParser, DEFAULT_LIMITS, getRangeDetails } from '../dist/index.js';
import { NS, REL, DRAW, esc, fixture, zip, worksheetXml, stylesXml, themeXml, workbookXml, relationships, drawingXml, centralEntries } from './fixtures.mjs';

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
    await assert.rejects(read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({row:'10000'})},{drawings:true}),/drawing row/);
    await assert.rejects(read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({image:true}),drawingRelationships:relationships([{id:'imageRel',kind:'image',target:'https://example.invalid/p.png',external:true}])},{drawings:true}),/External/);
    const w=await read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({extent:'9525000'})},{drawings:true});
    assert.throws(()=>parser.toHTML(w,{limits:{maxColumns:2}}),/expansion/);
    w.workSheets[0].drawings.push(structuredClone(w.workSheets[0].drawings[0]));assert.throws(()=>parser.toHTML(w,{limits:{maxDrawings:1}}),/Drawing count/);
});

test('PNG images use only embedded data URLs; drawing/media matching uses complete paths',async()=>{
    const png=Uint8Array.from([137,80,78,71,13,10,26,10,0,0,0,0]);
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
    await assert.rejects(read({...base,sheet:merged,extra:[['xl/worksheets/sheet2.xml',merged]]},{limits:{maxMergedCells:3}}),/Workbook grid/);
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
    const w=await read({sheet:worksheetXml(undefined,'A1','<drawing r:id="drawingRel"/>'),drawing:drawingXml({image:true}),drawingRelationships:relationships([{id:'imageRel',kind:'image',target:'../media/image.svg'}]),extra:[['xl/media/image.svg','<svg xmlns="http://www.w3.org/2000/svg" onload="alert(1)"/>']]},{drawings:true});
    assert.throws(()=>parser.toHTML(w),/Unsupported image format/);assert.throws(()=>parser.toHTMLSheet(w,0),/Unsupported image format/);
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
