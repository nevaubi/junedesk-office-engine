// Package-fidelity gate for presentations (bundled with tests/fidelity.ts into
// .build/fidelity.cjs). Build a deck with pptx-engine, apply one text edit
// through the same in-memory model the Slides engine uses, save, and assert
// that only the edited slide part changed: every other ZIP entry (layouts,
// masters, theme, rels, content types, the untouched slide) survives
// byte-identical and none disappears.
import JSZip from "jszip";
import { createHash } from "node:crypto";

import {
  addElement,
  createBlankPptx,
  duplicateSlide,
  openPptx,
  savePptx,
} from "../vendor/packages/pptx-engine/src/index";
import type { TextElement } from "../vendor/packages/pptx-engine/src/types";

/** Re-exported so ad-hoc probes can inspect fixtures with the bundled engine. */
export { openPptx, savePptx };

export type PptxFixtureCase = {
  name: string;
  build: () => Promise<Uint8Array>;
  /** Slide index that receives the edit; every other part must not change. */
  slideIndex: number;
  before: string;
  after: string;
};

const EMU_IN = 914400;

/** Two slides; slide 1 carries a text box with a known run. */
async function buildTwoSlideDeck(): Promise<Uint8Array> {
  const opened = await openPptx(await createBlankPptx());
  const first = opened.deck.slides[0]!;
  addElement(first, {
    kind: "textbox",
    offset: { x: EMU_IN, y: EMU_IN, cx: 6 * EMU_IN, cy: EMU_IN },
    paragraphs: [{ runs: [{ text: "Old" }] }],
  });
  first.structureDirty = true;
  duplicateSlide(opened, 0);
  return savePptx(opened);
}

export const PPTX_CORPUS: readonly PptxFixtureCase[] = [
  { name: "two-slide-textbox.pptx", build: buildTwoSlideDeck, slideIndex: 0, before: "Old", after: "Verified" },
];

export type PptxFixtureReport = {
  fixture: string;
  passed: boolean;
  error?: string;
  editedPart: string;
  changedEntries: string[];
  removedEntries: string[];
  unexpectedChanges: string[];
  preservedEntryCount: number;
};

async function entryHashes(bytes: Uint8Array): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(bytes);
  const out = new Map<string, string>();
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    out.set(path, createHash("sha256").update(await file.async("uint8array")).digest("hex"));
  }
  return out;
}

export async function verifyPptxCase(entry: PptxFixtureCase, source?: Uint8Array): Promise<PptxFixtureReport> {
  const bytes = source ?? (await entry.build());
  try {
    const opened = await openPptx(bytes);
    const slide = opened.deck.slides[entry.slideIndex];
    if (!slide) throw new Error(`fixture has no slide ${entry.slideIndex + 1}`);
    const target = slide.elements.find(
      (el): el is TextElement =>
        (el.type === "text" || el.type === "shape") &&
        (el.text?.paragraphs ?? []).some((p) => p.runs.some((r) => r.text === entry.before)),
    );
    if (!target) {
      const seen = slide.elements
        .map((el) => `${el.type}:${"text" in el ? (el.text?.paragraphs ?? []).map((p) => p.runs.map((r) => r.text).join("")).join("|") : ""}`)
        .join(", ");
      throw new Error(`no run reads "${entry.before}" on slide ${entry.slideIndex + 1} (elements: ${seen || "none"})`);
    }
    for (const p of target.text!.paragraphs) for (const r of p.runs) if (r.text === entry.before) r.text = entry.after;
    target.dirty = true;
    const saved = await savePptx(opened);

    const before = await entryHashes(bytes);
    const after = await entryHashes(saved);
    const editedPart = slide.path;
    const changedEntries = [...after].filter(([p, h]) => before.get(p) !== h).map(([p]) => p);
    const removedEntries = [...before.keys()].filter((p) => !after.has(p));
    const unexpectedChanges = [...changedEntries, ...removedEntries].filter((p) => p !== editedPart);
    const reopened = await openPptx(saved);
    const readsBack = reopened.deck.slides[entry.slideIndex]!.elements.some(
      (el) =>
        (el.type === "text" || el.type === "shape") &&
        (el.text?.paragraphs ?? []).some((p) => p.runs.some((r) => r.text === entry.after)),
    );
    return {
      fixture: entry.name,
      passed: unexpectedChanges.length === 0 && changedEntries.includes(editedPart) && readsBack,
      editedPart,
      changedEntries,
      removedEntries,
      unexpectedChanges,
      preservedEntryCount: after.size - changedEntries.length,
    };
  } catch (error) {
    return {
      fixture: entry.name,
      passed: false,
      error: error instanceof Error ? error.message : String(error),
      editedPart: "",
      changedEntries: [],
      removedEntries: [],
      unexpectedChanges: [],
      preservedEntryCount: 0,
    };
  }
}
