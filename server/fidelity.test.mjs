// Package-fidelity gate: one surgical edit per fixture must leave every other
// package entry byte-identical. Run after `node build.mjs` (needs .build/fidelity.cjs).
import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { createRequire } from "node:module";
import { resolve } from "node:path";
import { test } from "node:test";

const require = createRequire(import.meta.url);
const bundle = resolve(import.meta.dirname, "../.build/fidelity.cjs");

test("fidelity bundle is built", () => {
  assert.ok(existsSync(bundle), "run `node build.mjs` first");
});

const { CORPUS, verifyCase, PPTX_CORPUS, verifyPptxCase, verifyStyledSheetCreation } = existsSync(bundle)
  ? require(bundle)
  : { CORPUS: [], verifyCase: null, PPTX_CORPUS: [], verifyPptxCase: null };

test("first styled multi-sheet save allocates unique, resolvable relationships", async () => {
  const result = await verifyStyledSheetCreation();
  assert.equal(new Set(result.ids).size, result.ids.length, "relationship IDs must be unique");
  assert.equal(result.sheets.length, 3);
  for (const sheet of result.sheets) {
    assert.equal(sheet.matches, 1, sheet.name);
    assert.match(sheet.type, /\/worksheet$/);
    assert.equal(sheet.exists, true, sheet.name);
  }
});

for (const entry of CORPUS) {
  test(`untouched entries survive byte-identical: ${entry.name}`, async () => {
    const report = await verifyCase(entry);
    assert.equal(report.error, undefined, report.error);
    assert.deepEqual(report.unexpectedChanges, [], "unexpected package changes");
    assert.deepEqual(report.removedEntries, [], "package entries vanished");
    assert.ok(report.changedEntries.length > 0, "the edit did not change the worksheet part");
    assert.ok(report.preservedEntryCount > 0);
  });
}

for (const entry of PPTX_CORPUS) {
  test(`untouched entries survive byte-identical: ${entry.name}`, async () => {
    const report = await verifyPptxCase(entry);
    assert.equal(report.error, undefined, report.error);
    assert.deepEqual(report.unexpectedChanges, [], "unexpected package changes");
    assert.deepEqual(report.removedEntries, [], "package entries vanished");
    assert.deepEqual(report.changedEntries, [report.editedPart], "only the edited slide part may change");
    assert.ok(report.passed, "the edit did not read back from the saved deck");
    assert.ok(report.preservedEntryCount > 0);
  });
}
