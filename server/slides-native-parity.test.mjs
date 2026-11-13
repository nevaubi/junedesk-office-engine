import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { Module, createRequire } from 'node:module';
import { resolve } from 'node:path';
import { test } from 'node:test';
import { build } from 'esbuild';

const root = resolve(import.meta.dirname, '..');
const filename = resolve(root, 'slides-parity-test.cjs');
const bundle = await build({
  stdin: { contents: `
    export * from '@genoffice/pptx-engine';
    export * from './vendor/slides/src/main/ops/index.ts';
    export { listSlideAnimations } from './vendor/slides/src/main/ops/animation-ops.ts';
    export { readNativeSlideDetails } from './vendor/slides/src/main/native-details.ts';
    export { OP_DOCS } from './vendor/slides/src/shared/op-docs.ts';
  `, resolveDir: root, sourcefile: 'slides-parity-test.ts', loader: 'ts' },
  bundle: true, write: false, platform: 'node', format: 'cjs', target: 'node22',
  alias: {
    '@genoffice/pptx-engine': resolve(root, 'vendor/packages/pptx-engine/src/index.ts'),
    '@genoffice/docx-engine/math': resolve(root, 'vendor/packages/docx-engine/src/math.ts'),
  },
  plugins: [{ name: 'native-guides', setup(b) {
    b.onResolve({ filter: /\?raw$/ }, a => ({ path: resolve(a.resolveDir, a.path.slice(0, -4)), namespace: 'raw' }));
    b.onLoad({ filter: /.*/, namespace: 'raw' }, async a => ({ contents: await readFile(a.path, 'utf8'), loader: 'text' }));
  } }],
});
const compiled = new Module(filename);
compiled.filename = filename;
compiled.paths = Module._nodeModulePaths(root);
compiled._compile(bundle.outputFiles[0].text, filename);
const engine = compiled.exports;
const require = createRequire(import.meta.url);
const JSZip = require('jszip');
const names = ['addAnimation', 'addConnector', 'addSlideWithLayout', 'alignElements', 'distributeElements', 'insertEquation', 'removeAnimation', 'reorderAnimation', 'setShapeCustomGeometry', 'setSlideLayout', 'setTableStyle'];

function txn(opened, ...ops) {
  const result = engine.runTxn(opened, { ops });
  assert.equal(result.applied, true, JSON.stringify(result.failures));
  return result;
}
const target = el => ({ slide: 0, el });
async function fixture() {
  const opened = await engine.openPptx(await engine.createBlankPptx());
  txn(opened, ...[0, 1, 2].map(i => ({ op: 'addElement', target: {slide: 0}, kind: i === 2 ? 'ellipse' : 'rect', offset: {x: 500000 + i * 2200000, y: 500000 + i * 200000, cx: 1000000, cy: 700000}, paragraphs: [{runs: [{text: `Preserve ${i}`, fontSize: 24, color: '#172E4C'}]}] })));
  const stable = await engine.openPptx(await engine.savePptx(opened));
  return { opened: stable, ids: stable.deck.slides[0].elements.map(e => e.id) };
}
async function reopen(opened) { return engine.openPptx(await engine.savePptx(opened)); }
async function packageEntries(opened) {
  const zip = await JSZip.loadAsync(await engine.savePptx(opened));
  return new Map(await Promise.all(Object.values(zip.files).filter(f => !f.dir).map(async f => [f.name, await f.async('nodebuffer')])));
}

test('all eleven actual native handlers match callable browser guides; canonical engine copies cannot drift', async () => {
  for (const name of names) {
    assert.ok(engine.opNames().includes(name), name);
    assert.ok(engine.OP_DOCS[name] && engine.OP_DOCS[name].aiCallable !== false && !engine.OP_DOCS[name].pending, name);
  }
  for (const name of ['animation.ts', 'custgeom.ts', 'deck-text.ts', 'generate.ts', 'index.ts', 'notes.ts', 'parse.ts', 'table-edit.ts', 'table-style.ts', 'types.ts']) {
    const retained = await readFile(resolve(root, 'vendor/packages/pptx-engine/src', name), 'utf8');
    const canonical = await readFile(resolve(root, '../../src/writer/packages/pptx-engine/src', name), 'utf8');
    assert.equal(retained.replaceAll('\r\n', '\n'), canonical.replaceAll('\r\n', '\n'), name);
  }
  for (const group of ['text', 'element', 'insert', 'table', 'slide', 'deck']) {
    assert.equal((await readFile(resolve(root, 'vendor/slides/src/shared/prompts/ops', `${group}.md`), 'utf8')).replaceAll('\r\n', '\n'), (await readFile(resolve(root, '../../src/office/slides/src/shared/prompts/ops', `${group}.md`), 'utf8')).replaceAll('\r\n', '\n'));
  }
});

test('alignment/distribution and connected ellipse endpoints survive native save/reopen', async () => {
  const {opened, ids} = await fixture();
  const original = await packageEntries(opened);
  txn(opened, {op:'alignElements',target:{slide:0},els:ids,mode:'top'}, {op:'distributeElements',target:{slide:0},els:ids,axis:'horizontal'});
  txn(opened, {op:'addConnector',target:{slide:0},from:ids[0],to:ids[2],fromSide:'right',toSide:'left',kind:'elbow',arrow:'both',line:{color:'#172E4C',widthPt:2}});
  let next = await reopen(opened);
  const shapes = next.deck.slides[0].elements.slice(0,3);
  assert.deepEqual(shapes.map(s => s.transform.offset.y), [500000,500000,500000]);
  assert.equal(shapes[1].transform.offset.x - shapes[0].transform.offset.x, shapes[2].transform.offset.x - shapes[1].transform.offset.x);
  const connector = next.deck.slides[0].elements.find(e => e.connection);
  assert.equal(connector.connection.end.idx, 2, 'ellipse left connection is site 2');
  const beforeEnd = connector.transform.offset.x + connector.transform.offset.cx;
  txn(next, {op:'setTransform',target:target(shapes[2].id),box:{...shapes[2].transform.offset,x:7000000}});
  next = await reopen(next);
  const moved = next.deck.slides[0].elements.find(e => e.connection);
  assert.equal(moved.transform.offset.x + moved.transform.offset.cx, 7000000);
  assert.notEqual(beforeEnd,7000000);
  const result = await packageEntries(next);
  for (const [name, bytes] of original) if (name !== 'ppt/slides/slide1.xml') assert.deepEqual(result.get(name),bytes,name);
});

test('custom geometry is editable native XML and rejects non-finite/overflow paths atomically', async () => {
  const {opened, ids} = await fixture();
  const path = {w:1000000,h:700000,cmds:[{op:'M',pts:[0,0]},{op:'L',pts:[1000000,0]},{op:'Q',pts:[500000,700000,0,0]},{op:'Z',pts:[]}]};
  txn(opened,{op:'setShapeCustomGeometry',target:target(ids[0]),path});
  const next = await reopen(opened);
  assert.ok(next.deck.slides[0].elements[0].customGeometry);
  assert.match(next.deck.slides[0].elements[0].anchor.originalXml,/<a:custGeom>/);
  assert.match(next.deck.slides[0].elements[0].anchor.originalXml,/<a:quadBezTo>/);
  const original = await packageEntries(next);
  const invalid = engine.runTxn(next,{ops:[{op:'alignElements',target:{slide:0},els:next.deck.slides[0].elements.map(e=>e.id),mode:'bottom'},{op:'setShapeCustomGeometry',target:target(next.deck.slides[0].elements[0].id),path:{...path,w:1e100}}]});
  assert.equal(invalid.applied,false);
  assert.deepEqual(await packageEntries(next),original);
});

test('native equations survive save, reopen and an unrelated text-style edit, then explicit replacement works', async () => {
  const {opened, ids} = await fixture();
  txn(opened,{op:'insertEquation',target:target(ids[0]),latex:'\\frac{a}{b} + x^2'});
  let next=await reopen(opened);
  const shape=next.deck.slides[0].elements[0];
  assert.equal(shape.text.paragraphs.length,2);
  const equation=shape.text.paragraphs[1].runs[0];
  assert.match(equation.rawXml,/<m:f>/);
  assert.match(equation.rawXml,/<m:sSup>/);
  assert.ok(equation.text.includes('a'));
  txn(next,{op:'setFont',target:target(shape.id),font:{bold:true}});
  next=await reopen(next);
  assert.equal(next.deck.slides[0].elements[0].text.paragraphs[0].runs[0].bold,true);
  assert.equal(next.deck.slides[0].elements[0].text.paragraphs[1].runs[0].rawXml,equation.rawXml);
  txn(next,{op:'setText',target:target(next.deck.slides[0].elements[0].id),paragraphs:[{runs:[{text:'Explicit replacement'}]}]});
  next=await reopen(next);
  assert.doesNotMatch(next.deck.slides[0].elements[0].anchor.originalXml,/<a14:m>/);
  assert.match(next.deck.slides[0].elements[0].anchor.originalXml,/Explicit replacement/);
  txn(next,{op:'insertEquation',target:{slide:0},box:{x:1000000,y:1000000,cx:3000000,cy:800000},latex:'E = mc^2'});
  const appended=await reopen(next);
  assert.equal(appended.deck.slides[0].elements.length,4);
  assert.ok(appended.deck.slides[0].elements[3].text.paragraphs[0].runs[0].rawXml);
});

test('connector creation rejects unsupported connection geometry without guessing native endpoints', async()=>{
  const {opened,ids}=await fixture();
  txn(opened,{op:'setShapeGeometry',target:target(ids[0]),prst:'triangle'});
  const before=await packageEntries(opened);
  const result=engine.runTxn(opened,{ops:[{op:'addConnector',target:{slide:0},from:ids[0],to:ids[1]}]});
  assert.equal(result.applied,false); assert.match(result.failures[0].error,/connection geometry/);
  assert.deepEqual(await packageEntries(opened),before);
});

test('failed dependent timeline operation rolls back a newly materialized layout and every package entry', async()=>{
  const {opened}=await fixture();
  const before=await packageEntries(opened);
  const result=engine.runTxn(opened,{ops:[{op:'addSlideWithLayout',layout:'Title and Content'},{op:'removeAnimation',target:{slide:0},seq:22}]});
  assert.equal(result.applied,false); assert.equal(opened.deck.slides.length,1);
  assert.deepEqual(await packageEntries(opened),before);
});

test('adding an animation preserves an imported unknown effect instead of converting it to a guessed preset', async()=>{
  const {opened,ids}=await fixture();
  txn(opened,{op:'addAnimation',target:target(ids[0]),effect:'fade'});
  opened.deck.slides[0].bodySuffix=opened.deck.slides[0].bodySuffix.replace('presetID="10"','presetID="999"');
  const saved=await reopen(opened);
  assert.equal(engine.listSlideAnimations(saved.deck.slides[0])[0].custom,true);
  const originalEffect=engine.getSlideAnimations(saved.deck.slides[0])[0].presetXml;
  txn(saved,{op:'addAnimation',target:target(saved.deck.slides[0].elements[1].id),effect:'appear'});
  const next=await reopen(saved);
  assert.equal(engine.getSlideAnimations(next.deck.slides[0])[0].presetXml,originalEffect);
  assert.equal(engine.listSlideAnimations(next.deck.slides[0]).length,2);
  const before=await packageEntries(next);
  const unsafe=engine.runTxn(next,{ops:[{op:'reorderAnimation',target:{slide:0},seq:1,to:0}]});
  assert.equal(unsafe.applied,false); assert.match(unsafe.failures[0].error,/unmodeled timing/);
  assert.deepEqual(await packageEntries(next),before);
});

test('source files cannot forge the internal equation preservation marker', async()=>{
  const {opened}=await fixture();
  const slide=opened.deck.slides[0];
  slide.elements[0].anchor.originalXml=slide.elements[0].anchor.originalXml.replace('<a:r>','<a:r gxRaw="not-valid-base64!">');
  slide.structureDirty=true;
  const next=await reopen(opened);
  assert.equal(next.deck.slides[0].elements[0].text.paragraphs[0].runs[0].rawXml,undefined);
  assert.equal(next.deck.slides[0].elements[0].text.paragraphs[0].runs[0].text,'Preserve 0');
});

test('animation add/reorder/remove, direction, duration and native orphan seq IDs round-trip', async () => {
  const {opened, ids}=await fixture();
  txn(opened,{op:'addAnimation',target:target(ids[0]),effect:'flyIn',direction:'left',duration:600,trigger:'afterPrev'}, {op:'addAnimation',target:target(ids[1]),effect:'fade',duration:450});
  let next=await reopen(opened);
  let list=engine.listSlideAnimations(next.deck.slides[0]);
  assert.equal(list[0].direction,'left'); assert.equal(list[0].durationMs,600);
  txn(next,{op:'reorderAnimation',target:{slide:0},seq:1,to:0});
  next=await reopen(next); list=engine.listSlideAnimations(next.deck.slides[0]);
  assert.equal(list[0].effect,'fade'); assert.equal(list[1].effect,'flyIn');
  txn(next,{op:'removeAnimation',target:{slide:0},seq:0});
  next=await reopen(next); assert.deepEqual(engine.listSlideAnimations(next.deck.slides[0]).map(a=>a.effect),['flyIn']);
  const slide=next.deck.slides[0];
  const orphan={...engine.getSlideAnimations(slide)[0],spid:9999,effect:'fade'};
  engine.setSlideAnimations(slide,[orphan,...engine.getSlideAnimations(slide)]);
  const read=engine.readNativeSlideDetails(next,0);
  assert.equal(read.animations[0].el,null); assert.equal(read.animations[1].seq,1);
  const before=await packageEntries(next);
  const invalid=engine.runTxn(next,{ops:[{op:'removeAnimation',target:{slide:0},seq:88}]});
  assert.equal(invalid.applied,false); assert.deepEqual(await packageEntries(next),before);
});

test('layout names/indices and default appends create native linked layouts without changing original content', async () => {
  const {opened}=await fixture();
  const original=await packageEntries(opened);
  const details=engine.readNativeSlideDetails(opened,0);
  const title=details.layouts.find(l=>l.path.startsWith('builtin:') && l.name==='Title and Content');
  assert.ok(title, JSON.stringify(details.layouts));
  const dry=engine.runTxn(opened,{ops:[{op:'addSlideWithLayout',layout:title.name}],dryRun:true});
  assert.equal(dry.failures,undefined); assert.deepEqual(await packageEntries(opened),original);
  txn(opened,{op:'addSlideWithLayout',layout:title.name});
  let next=await reopen(opened);
  assert.equal(next.deck.slides.length,2); assert.ok(next.deck.slides[1].elements.length>=2);
  assert.deepEqual((await packageEntries(next)).get('ppt/slides/slide1.xml'),original.get('ppt/slides/slide1.xml'));
  const layout=engine.readNativeSlideDetails(next,1).layouts.find(l=>l.name==='Title Only');
  txn(next,{op:'setSlideLayout',target:{slide:1},layout:layout.index});
  next=await reopen(next);
  assert.ok(next.deck.slides[1].layoutPath); assert.equal(next.deck.slides.length,2);
});

test('table presets and named native styles/flags persist; malformed look flags and missing cells are rejected', async () => {
  const {opened}=await fixture();
  const r=txn(opened,{op:'addTable',target:{slide:0},rows:2,cols:2,offset:{x:500000,y:2000000,cx:5000000,cy:2000000}});
  const id=r.records[0].created[0];
  txn(opened,{op:'setTableStyle',target:target(id),styleName:'headerDarkBlue'});
  let next=await reopen(opened); let table=next.deck.slides[0].elements.find(e=>e.type==='table');
  assert.ok(next.archive.readText('ppt/tableStyles.xml'));
  txn(next,{op:'setTableStyle',target:target(table.id),styleId:'Medium Style 2 - Accent 1',firstRow:true,lastRow:true,bandCol:true,bandRow:false,keepFormatting:true});
  next=await reopen(next); table=next.deck.slides[0].elements.find(e=>e.type==='table');
  assert.match(table.anchor.originalXml,/lastRow="1"/); assert.match(table.anchor.originalXml,/bandCol="1"/);
  assert.match(table.anchor.originalXml,new RegExp(engine.resolveBuiltinTableStyleId('Medium Style 2 - Accent 1').replace(/[{}]/g,'\\$&')));
  for(const bad of [{firstRow:'false'},{borderWidthPt:'2'},{cells:[{row:100,col:0}],shadingColor:'#FFFFFF'}]) {
    const before=await packageEntries(next);
    assert.equal(engine.runTxn(next,{ops:[{op:'setTableStyle',target:target(table.id),...bad}]}).applied,false);
    assert.deepEqual(await packageEntries(next),before);
  }
});
