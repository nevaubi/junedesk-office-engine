import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { test } from 'node:test';
import { build } from 'esbuild';
import { resolve } from 'node:path';

// Bundle the retained engine entry modules with its actual installed Zod4,
// exactly the source tree used by build.mjs (not injected browser modules).
const root = resolve(import.meta.dirname, '..');
const result = await build({ stdin: { contents: `
  export { workbookChartEditSchema, workbookVisualAddSchema } from './vendor/sheets/src/shared/desktop-api.ts';
  export { buildChartXml } from './vendor/sheets/src/gateway/xlsx-drawing-add.ts';
  export { applyChartEdit } from './vendor/sheets/src/gateway/xlsx-chart.ts';
`, resolveDir: root, sourcefile: 'chart-cache-boundary.ts', loader: 'ts' }, bundle: true, write: false, platform: 'node', format: 'esm', target: 'node22' });
const engine = await import(`data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString('base64')}`);
const series = { name: 'Revenue', categories: ['2021', '2022', '2023'], values: [12, 0, 0], blanks: [1], valuesRef: "'Summary'!$B$2:$D$2", color: '#172E4C' };

test('actual retained engine admits sparse edit/add payloads and serializes gaps without losing real zero', () => {
  const add = engine.workbookVisualAddSchema.parse({ sheetId: 's', anchor: { fromRow: 0, fromColumn: 0, fromRowOffset: 0, fromColumnOffset: 0, toRow: 20, toColumn: 8, toRowOffset: 0, toColumnOffset: 0 }, chart: { chartType: 'line', title: 'Revenue', series: [series] } });
  const xml = engine.buildChartXml(add.chart);
  const edited = engine.applyChartEdit(xml, engine.workbookChartEditSchema.parse({ chartPath: 'xl/charts/chart1.xml', series: [{ index: 0, values: series.values, blanks: [1] }] }));
  for (const text of [xml, edited]) {
    const values = /<c:val>([\s\S]*?)<\/c:val>/.exec(text)[1];
    assert.match(values, /<c:ptCount val="3"\/>/);
    assert.doesNotMatch(values, /<c:pt idx="1">/);
    assert.match(values, /<c:pt idx="2"><c:v>0<\/c:v><\/c:pt>/);
    assert.ok(values.includes(series.valuesRef));
    assert.match(text, /172E4C/);
  }
});

test('browser and retained engine chart wire/native/cache modules cannot silently drift', async () => {
  for (const file of ['gateway/xlsx-chart.ts', 'gateway/xlsx-drawing-add.ts', 'gateway/xlsx-gateway.ts', 'domain/chart-cache.ts', 'domain/chart-visual.ts', 'shared/desktop-api.ts']) {
    const [browser, retained] = await Promise.all([
      readFile(resolve(root, '../../src/office/sheets/src', file), 'utf8'),
      readFile(resolve(root, 'vendor/sheets/src', file), 'utf8'),
    ]);
    assert.equal(retained.replaceAll('\r\n', '\n'), browser.replaceAll('\r\n', '\n'), file);
  }
});

test('retained engine accepts independent primary/secondary number formats and exports exact native axis XML', () => {
  const xml = engine.buildChartXml({ chartType: 'combo', title: 'Income and margin', series: [series, {...series,name:'Margin'}] });
  const edit = engine.workbookChartEditSchema.parse({ chartPath:'xl/charts/chart2.xml', valueAxisFormats:{primary:'"$"#,##0',secondary:'0.0%'} });
  const output = engine.applyChartEdit(xml,edit);
  const axes = [...output.matchAll(/<c:valAx>([\s\S]*?)<\/c:valAx>/g)].map(m=>m[1]);
  assert.match(axes.find(axis=>axis.includes('<c:axPos val="l"/>')),/formatCode="&quot;\$&quot;#,##0" sourceLinked="0"/);
  assert.match(axes.find(axis=>axis.includes('<c:axPos val="r"/>')),/formatCode="0.0%" sourceLinked="0"/);
  assert.equal(output.replace(/<c:numFmt\b[^>]*\/>/g,''),xml);
  assert.equal(engine.workbookChartEditSchema.safeParse({chartPath:'xl/charts/chart2.xml',valueAxisFormats:{}}).success,false);
});
