// Helpers the HTTP layer needs from the retained engine, bundled without Electron.
import { blankXlsxBuffer, csvToXlsxBuffer, parseCsv } from "../vendor/sheets/src/gateway/csv-import";

export const makeBlank = () => blankXlsxBuffer("Sheet1");
export { csvToXlsxBuffer, parseCsv };
export { createBlankPptx } from "../vendor/packages/pptx-engine/src/blank";

const fail = (status: number, message: string) => Object.assign(new Error(message), { status });

/** CSV imports become editable workbooks; the user's source file is never written. */
export async function importCsv(bytes: Uint8Array, name: string): Promise<{ bytes: Uint8Array; name: string }> {
  if (bytes.byteLength > 16 * 1024 * 1024) throw fail(413, "CSV exceeds the 16 MB limit.");
  let text: string;
  try {
    text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    throw fail(422, "CSV must use UTF-8 encoding.");
  }
  if (text.includes("\0")) throw fail(422, "The upload is not ordinary CSV text.");
  let quotes = false;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === '"') {
      if (quotes && text[i + 1] === '"') i++;
      else quotes = !quotes;
    }
  }
  if (quotes) throw fail(422, "CSV has an unterminated quoted field.");
  const rows = parseCsv(text);
  let cells = 0;
  for (const row of rows) {
    cells += row.length;
    if (row.length > 1024 || row.some((c) => c.length > 32767)) throw fail(413, "CSV has an oversized row or cell.");
  }
  if (rows.length > 100_000 || cells > 250_000) throw fail(413, "CSV exceeds the row/cell limit (100,000 rows / 250,000 cells).");
  return { bytes: new Uint8Array(await csvToXlsxBuffer(text)), name: name.replace(/\.csv$/i, ".xlsx") };
}
