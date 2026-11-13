// A worksheet part with no <sheetData> element. The schema requires one, but some
// generators (and sheets added by tools) omit it when the sheet has no cells. Every
// save-time patcher (sheet view, column attributes, filters, protection) anchors on
// <sheetFormatPr>, <cols> or <sheetData> and throws "Worksheet has no sheetData
// element" when none exists, which failed the whole workbook save with a 500.
// Reading each worksheet through ensureSheetData gives those patchers an anchor.

/// Elements the schema places AFTER sheetData, in schema order. The first one
/// present marks where an empty sheetData belongs.
const AFTER_SHEET_DATA = [
  "sheetCalcPr",
  "sheetProtection",
  "protectedRanges",
  "scenarios",
  "autoFilter",
  "sortState",
  "dataConsolidate",
  "customSheetViews",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "drawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];
const AFTER_PATTERN = new RegExp(`<(?:${AFTER_SHEET_DATA.join("|")})\\b`);

/** Returns the worksheet XML with an empty <sheetData/> added when it has none. */
export function ensureSheetData(worksheetXml: string): string {
  if (/<sheetData\b/.test(worksheetXml)) return worksheetXml;
  const root = /<worksheet\b[^>]*?(\/?)>/.exec(worksheetXml);
  if (!root) return worksheetXml;
  if (root[1] === "/") {
    const open = root[0].slice(0, -2).trimEnd();
    return (
      worksheetXml.slice(0, root.index) +
      `${open}><sheetData/></worksheet>` +
      worksheetXml.slice(root.index + root[0].length)
    );
  }
  const bodyStart = root.index + root[0].length;
  const later = AFTER_PATTERN.exec(worksheetXml.slice(bodyStart));
  const at = later
    ? bodyStart + later.index
    : worksheetXml.lastIndexOf("</worksheet>") >= bodyStart
      ? worksheetXml.lastIndexOf("</worksheet>")
      : bodyStart;
  return worksheetXml.slice(0, at) + "<sheetData/>" + worksheetXml.slice(at);
}
