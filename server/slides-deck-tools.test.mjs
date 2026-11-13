// Deck-wide selector ops (findReplace across runs, formatText, replaceFont,
// replaceColor, matchElements, matchAcrossSlides), batch-start addressing,
// duplicate-keeps-notes, and the op docs for them, run through the retained
// engine's real executor on synthetic decks.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { Module } from "node:module";
import { resolve } from "node:path";
import { test } from "node:test";
import { build } from "esbuild";

const root = resolve(import.meta.dirname, "..");
const filename = resolve(root, "slides-deck-tools-test.cjs");
const bundle = await build({
  stdin: {
    contents: `
    export * from '@genoffice/pptx-engine';
    export * from './vendor/slides/src/main/ops/index.ts';
    export { OP_DOCS } from './vendor/slides/src/shared/op-docs.ts';
  `,
    resolveDir: root,
    sourcefile: "slides-deck-tools-test.ts",
    loader: "ts",
  },
  bundle: true,
  write: false,
  platform: "node",
  format: "cjs",
  target: "node22",
  alias: {
    "@genoffice/pptx-engine": resolve(root, "vendor/packages/pptx-engine/src/index.ts"),
    "@genoffice/docx-engine/math": resolve(root, "vendor/packages/docx-engine/src/math.ts"),
  },
  plugins: [
    {
      name: "native-guides",
      setup(b) {
        b.onResolve({ filter: /\?raw$/ }, (a) => ({
          path: resolve(a.resolveDir, a.path.slice(0, -4)),
          namespace: "raw",
        }));
        b.onLoad({ filter: /.*/, namespace: "raw" }, async (a) => ({
          contents: await readFile(a.path, "utf8"),
          loader: "text",
        }));
      },
    },
  ],
});
const compiled = new Module(filename);
compiled.filename = filename;
compiled.paths = Module._nodeModulePaths(root);
compiled._compile(bundle.outputFiles[0].text, filename);
const engine = compiled.exports;

function txn(opened, ...ops) {
  const result = engine.runTxn(opened, { ops });
  assert.equal(result.applied, true, JSON.stringify(result.failures));
  return result;
}
function refuse(opened, ops, pattern) {
  const result = engine.runTxn(opened, { ops });
  assert.equal(result.applied, false, "the transaction must be refused");
  assert.match(result.failures.map((f) => f.error).join("\n"), pattern);
  return result;
}
const reopen = async (opened) => engine.openPptx(await engine.savePptx(opened));
const textOf = (el) =>
  (el.text?.paragraphs ?? []).map((p) => p.runs.map((r) => r.text).join("")).join("\n");
const box = (i) => ({ x: 500000 + i * 2400000, y: 600000 + i * 300000, cx: 2000000, cy: 600000 });

/** A deck of `count` slides, each with the given text boxes (runs per box). */
async function deck(count, boxes) {
  const opened = await engine.openPptx(await engine.createBlankPptx());
  for (let s = 1; s < count; s++) txn(opened, { op: "addBlankSlide", target: { slide: s - 1 } });
  for (let s = 0; s < count; s++) {
    txn(
      opened,
      ...boxes(s).map((runs, i) => ({
        op: "addElement",
        target: { slide: s },
        kind: "textbox",
        offset: box(i),
        paragraphs: [{ runs }],
      })),
    );
  }
  return reopen(opened);
}
const els = (opened, s) =>
  opened.deck.slides[s].elements.filter((e) => e.type === "text" || e.type === "shape");

test("findReplace replaces a match spanning differently formatted runs and keeps the first run's format", async () => {
  const opened = await deck(1, () => [
    [
      { text: "Acme ", bold: true, fontSize: 20 },
      { text: "Corp", italic: true, fontSize: 20 },
      { text: " wins", fontSize: 20 },
    ],
  ]);
  const r = txn(opened, {
    op: "findReplace",
    find: "acme corp",
    replace: "Acme Corporation",
    expectedCount: 1,
  });
  assert.equal(r.records[0].after.count, 1);
  const next = await reopen(opened);
  const runs = els(next, 0)[0].text.paragraphs[0].runs.filter((x) => x.text);
  assert.equal(runs.map((x) => x.text).join(""), "Acme Corporation wins");
  assert.equal(runs[0].text, "Acme Corporation");
  assert.equal(runs[0].bold, true);
  assert.notEqual(runs[0].italic, true);
});

test("findReplace regex with captures, whole word, slide scope and an expected count that refuses on mismatch", async () => {
  const opened = await deck(2, (s) => [
    [{ text: `Filed 3/14/2024 and 12/1/2025 in cat catalog ${s}` }],
  ]);
  refuse(
    opened,
    [{ op: "findReplace", find: "cat", replace: "dog", wholeWord: true, expectedCount: 5 }],
    /matches 2 time/,
  );
  assert.match(textOf(els(opened, 0)[0]), /cat catalog/, "a refused count changes nothing");
  txn(opened, {
    op: "findReplace",
    find: "cat",
    replace: "dog",
    wholeWord: true,
    expectedCount: 2,
  });
  assert.match(textOf(els(opened, 1)[0]), /dog catalog/);
  const r = txn(opened, {
    op: "findReplace",
    find: "(\\d{1,2})/(\\d{1,2})/(\\d{4})",
    replace: "$3-$1-$2",
    regex: true,
    slides: { from: 1, to: 1 },
  });
  assert.equal(r.records[0].after.count, 2);
  assert.match(textOf(els(opened, 1)[0]), /Filed 2024-3-14 and 2025-12-1/);
  assert.match(textOf(els(opened, 0)[0]), /3\/14\/2024/, "slide 0 is outside the scope");
  refuse(
    opened,
    [{ op: "findReplace", find: "(a+)+", replace: "x", regex: true }],
    /nested quantifiers/,
  );
});

test("findReplace includeNotes rewrites speaker notes run by run", async () => {
  const opened = await deck(1, () => [[{ text: "Body mentions Acme" }]]);
  txn(opened, { op: "setNotes", target: { slide: 0 }, text: "Ask Acme about the Acme filing" });
  const without = txn(opened, { op: "findReplace", find: "Acme", replace: "Beta" });
  assert.equal(without.records[0].after.count, 1, "notes are excluded by default");
  const r = txn(opened, {
    op: "findReplace",
    find: "Acme",
    replace: "Gamma",
    includeNotes: true,
    expectedCount: 2,
  });
  assert.deepEqual(r.records[0].after.slides, [{ slide: 0, count: 2, notes: 2 }]);
  const next = await reopen(opened);
  assert.equal(
    engine.getSlideNotes(next.archive, next.deck.slides[0].path),
    "Ask Gamma about the Gamma filing",
  );
});

test("formatText applies one font patch to every element of a role on a slide range", async () => {
  const opened = await deck(3, (s) => [[{ text: `Box A ${s}` }], [{ text: `Box B ${s}` }]]);
  refuse(
    opened,
    [
      {
        op: "formatText",
        roles: ["text"],
        slides: { from: 1, to: 2 },
        font: { bold: true },
        expectedCount: 3,
      },
    ],
    /matches 4/,
  );
  const r = txn(opened, {
    op: "formatText",
    roles: ["text"],
    slides: { from: 1, to: 2 },
    font: { bold: true, fontSizePt: 28, color: "#172E4C" },
    paragraph: { align: "center" },
    expectedCount: 4,
  });
  assert.equal(r.records[0].after.count, 4);
  assert.deepEqual(r.records[0].after.byRole, { text: 4 });
  const next = await reopen(opened);
  for (const s of [1, 2])
    for (const el of els(next, s)) {
      const run = el.text.paragraphs[0].runs[0];
      assert.equal(run.bold, true);
      assert.equal(run.fontSize, 28);
      assert.equal(el.text.paragraphs[0].align, "center");
    }
  assert.notEqual(els(next, 0)[0].text.paragraphs[0].runs[0].bold, true, "slide 0 untouched");
  refuse(
    opened,
    [{ op: "formatText", roles: ["title"], font: { bold: true } }],
    /no element with text/,
  );
  refuse(opened, [{ op: "formatText", roles: ["headline"], font: { bold: true } }], /unknown role/);
});

test("replaceFont rewrites typeface references deck-wide and per slide scope", async () => {
  const opened = await deck(2, (s) => [
    [{ text: `Arial text ${s}`, fontFamily: "Arial" }],
    [{ text: `Other ${s}`, fontFamily: "Georgia" }],
  ]);
  const scoped = txn(opened, { op: "replaceFont", from: "Arial", to: "Calibri", slides: [1] });
  assert.equal(scoped.records[0].after.slideRefs, 1);
  let next = await reopen(opened);
  assert.equal(els(next, 1)[0].text.paragraphs[0].runs[0].fontFamily, "Calibri");
  assert.equal(els(next, 0)[0].text.paragraphs[0].runs[0].fontFamily, "Arial");
  const wide = txn(next, { op: "replaceFont", from: "Arial", to: "Calibri" });
  assert.ok(wide.records[0].after.slideRefs >= 1);
  assert.equal(wide.records[0].after.remainingRuns, 0);
  next = await reopen(next);
  for (const s of [0, 1]) {
    assert.equal(els(next, s)[0].text.paragraphs[0].runs[0].fontFamily, "Calibri");
    assert.equal(els(next, s)[1].text.paragraphs[0].runs[0].fontFamily, "Georgia");
  }
  refuse(next, [{ op: "replaceFont", from: "Arial", to: "Calibri" }], /not used/);
});

test("replaceColor rewrites explicit colors on slides with an exact count", async () => {
  const opened = await deck(2, (s) => [
    [{ text: `Blue ${s}`, color: "#1A73E8" }],
    [{ text: `Red ${s}`, color: "#EA4335" }],
  ]);
  refuse(
    opened,
    [{ op: "replaceColor", from: "#1A73E8", to: "#172E4C", expectedCount: 1 }],
    /matches 2/,
  );
  const r = txn(opened, {
    op: "replaceColor",
    map: [{ from: "#1a73e8", to: "#172E4C" }],
    expectedCount: 2,
  });
  assert.equal(r.records[0].after.replaced, 2);
  const next = await reopen(opened);
  for (const s of [0, 1]) {
    assert.equal(els(next, s)[0].text.paragraphs[0].runs[0].color.toUpperCase(), "#172E4C");
    assert.equal(els(next, s)[1].text.paragraphs[0].runs[0].color.toUpperCase(), "#EA4335");
  }
  refuse(next, [{ op: "replaceColor", from: "#123456", to: "#000000" }], /none of/);
});

test("matchElements and matchAcrossSlides copy size and position from a reference element", async () => {
  const opened = await deck(3, (s) => [[{ text: `Ref ${s}` }], [{ text: `Other ${s}` }]]);
  const [ref, other] = els(opened, 0);
  txn(opened, {
    op: "matchElements",
    target: { slide: 0, el: ref.id },
    els: [other.id],
    match: ["size", "top"],
  });
  let next = await reopen(opened);
  const [r0, o0] = els(next, 0);
  assert.equal(o0.transform.offset.cx, r0.transform.offset.cx);
  assert.equal(o0.transform.offset.y, r0.transform.offset.y);
  assert.notEqual(o0.transform.offset.x, r0.transform.offset.x, "left was not requested");
  // Two text boxes per slide: the role is ambiguous, so every slide is skipped
  refuse(
    next,
    [{ op: "matchAcrossSlides", target: { slide: 0, el: r0.id }, match: ["position"] }],
    /exactly one text/,
  );
  // One box per slide after deleting the second ones, then match across
  txn(
    next,
    ...[1, 2].map((s) => ({ op: "deleteElement", target: { slide: s, el: els(next, s)[1].id } })),
  );
  txn(next, {
    op: "setTransform",
    target: { slide: 0, el: r0.id },
    box: { x: 111111, y: 222222, cx: 3333333, cy: 444444 },
  });
  const m = txn(next, {
    op: "matchAcrossSlides",
    target: { slide: 0, el: r0.id },
    role: "text",
    match: ["position", "size"],
    slides: [1, 2],
    expectedCount: 2,
  });
  assert.equal(m.records[0].after.changed, 2);
  next = await reopen(next);
  for (const s of [1, 2])
    assert.deepEqual(els(next, s)[0].transform.offset, {
      x: 111111,
      y: 222222,
      cx: 3333333,
      cy: 444444,
    });
});

test("a multi-op batch resolves every target against the deck as it was at batch start", async () => {
  const opened = await deck(3, (s) => [[{ text: `Slide ${s}` }]]);
  const third = els(opened, 2)[0];
  // Old behavior: after deleting slide 0, "slide 2" was out of range (or another page)
  const r = txn(
    opened,
    { op: "deleteSlide", target: { slide: 0 } },
    {
      op: "setText",
      target: { slide: 2, el: third.id },
      paragraphs: [{ runs: [{ text: "Edited third" }] }],
    },
    { op: "findReplace", find: "Slide 1", replace: "Second", slideIndex: 1 },
  );
  assert.equal(r.records.length, 3);
  assert.equal(r.records[1].op.target.slide, 2, "records keep the caller's addressing");
  const next = await reopen(opened);
  assert.deepEqual(
    next.deck.slides.map((_, i) => textOf(els(next, i)[0])),
    ["Second", "Edited third"],
  );
  // A batch naming a slide an earlier op deleted fails atomically with a guided error
  refuse(
    next,
    [
      { op: "deleteSlide", target: { slide: 0 } },
      {
        op: "setText",
        target: { slide: 0, el: els(next, 0)[0].id },
        paragraphs: [{ runs: [{ text: "x" }] }],
      },
    ],
    /no slide "s_/,
  );
  assert.equal(next.deck.slides.length, 2);
  // Single-op requests are untouched by pinning
  assert.equal(
    engine.pinBatchRefs([{ op: "deleteSlide", target: { slide: 0 } }], next)[0].target.slide,
    0,
  );
});

test("duplicateSlide copies the source notes into a part of the duplicate's own", async () => {
  const opened = await deck(1, () => [[{ text: "Body" }]]);
  txn(opened, { op: "setNotes", target: { slide: 0 }, text: "Speaker line one\nline two" });
  txn(opened, { op: "duplicateSlide", target: { slide: 0 } });
  let next = await reopen(opened);
  const [a, b] = next.deck.slides;
  const pa = engine.notesPathForSlide(next.archive, a.path);
  const pb = engine.notesPathForSlide(next.archive, b.path);
  assert.ok(pa && pb && pa !== pb, "each slide owns a notes part");
  assert.equal(engine.getSlideNotes(next.archive, b.path), "Speaker line one\nline two");
  const rels = next.archive.readText(pb.replace("notesSlides/", "notesSlides/_rels/") + ".rels");
  assert.match(rels, new RegExp(`Target="\\.\\./slides/${b.path.split("/").pop()}"`));
  txn(next, { op: "setNotes", target: { slide: 1 }, text: "Changed copy" });
  next = await reopen(next);
  assert.equal(
    engine.getSlideNotes(next.archive, next.deck.slides[0].path),
    "Speaker line one\nline two",
  );
  assert.equal(engine.getSlideNotes(next.archive, next.deck.slides[1].path), "Changed copy");
  // A blank layout copy (clearText) gets no notes
  txn(next, { op: "duplicateSlide", target: { slide: 0 }, clearText: true });
  next = await reopen(next);
  assert.equal(engine.getSlideNotes(next.archive, next.deck.slides[1].path), "");
});

test("the new selector ops are documented, callable, and every doc example validates", async () => {
  const opened = await deck(2, () => [[{ text: "Text" }], [{ text: "Shape" }]]);
  const [t, s] = els(opened, 0);
  const sub = (json) => json.replaceAll("e_TEXT", t.id).replaceAll("e_SHAPE", s.id);
  for (const name of [
    "findReplace",
    "formatText",
    "replaceFont",
    "replaceColor",
    "matchElements",
    "matchAcrossSlides",
  ]) {
    assert.ok(engine.opNames().includes(name), name);
    const doc = engine.OP_DOCS[name];
    assert.ok(doc && doc.aiCallable !== false && !doc.pending, name);
    assert.ok(doc.examples.length > 0, name);
    for (const example of doc.examples) {
      const dry = engine.runTxn(opened, { ops: [JSON.parse(sub(example))], dryRun: true });
      assert.equal(
        dry.failures,
        undefined,
        `${name}: ${example} -> ${JSON.stringify(dry.failures)}`,
      );
    }
  }
  for (const name of engine.opNames())
    assert.ok(engine.OP_DOCS[name], `registered op ${name} has docs`);
});
