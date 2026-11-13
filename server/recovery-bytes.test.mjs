import assert from "node:assert/strict";
import { test } from "node:test";
import { operationKind, recoveryResult, RECOVERY_MAX_BYTES } from "./boundary.mjs";

test("recovery-bytes is a read operation: it never advances the editing sequence or marks the document dirty", () => {
  assert.equal(operationKind("sheets", "workbook:recovery-bytes"), "read");
  assert.equal(operationKind("slides", "slides:recovery-bytes"), "read");
});

test("recovery-bytes is not reachable across document kinds", () => {
  assert.throws(() => operationKind("pdf", "pdf:recovery-bytes"));
  assert.throws(() => operationKind("sheets", "slides:recovery-bytes"));
  assert.throws(() => operationKind("slides", "workbook:recovery-bytes"));
});

test("a recovery package travels only when it fits the inline cap", () => {
  const small = new Uint8Array(RECOVERY_MAX_BYTES);
  const big = new Uint8Array(RECOVERY_MAX_BYTES + 1);
  assert.deepEqual(Object.keys(recoveryResult(small)).sort(), ["bytes", "dirty", "found", "size"]);
  const over = recoveryResult(big);
  assert.equal(over.found, true);
  assert.equal(over.size, RECOVERY_MAX_BYTES + 1);
  assert.equal("bytes" in over, false);
});

test("no recovery copy reports found false and clean", () => {
  assert.deepEqual(recoveryResult(null), { found: false, dirty: false });
});
