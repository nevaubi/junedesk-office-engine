/// Excel stores post-2007 function names in file formulas behind an
/// `_xlfn.` marker (worksheet-scope dynamic-array functions behind
/// `_xlfn._xlws.`) and shows the plain name in the UI. The sidecar strips
/// the markers on read; this restores them when a formula is serialized
/// back into sheet XML — without the marker Excel repairs the call to
/// #NAME?.
///
/// The same pass stores LET names and LAMBDA parameters as `_xlpm.` names
/// and rewrites the `@` and `#` operators to `_xlfn.SINGLE()` and
/// `_xlfn.ANCHORARRAY()`. The conversion lives in `./formula-storage`
/// (pure, idempotent, unit tested) and the name lists in
/// `./function-registry` (byte copies of src/lib/sheets/*; the service build has no
/// `@/` alias, keep them in sync).
///
/// Modified by Seeger Weiss (lit-ai) from GenOffice (Apache-2.0): delegates to formula-storage.

import { toStorageFormula } from './formula-storage'
import { XLFN_FUNCTION_NAMES, XLWS_FUNCTION_NAMES } from './function-registry'

/// Functions Excel 2007 does not know, keyed by the marker they need.
export const XLWS_FUNCTIONS: ReadonlySet<string> = new Set(XLWS_FUNCTION_NAMES)
export const XLFN_FUNCTIONS: ReadonlySet<string> = new Set(XLFN_FUNCTION_NAMES)

/// Converts a display formula to its storage form: future-function calls get
/// their markers, LET/LAMBDA names get `_xlpm.`, `@x` and `ref#` become
/// SINGLE/ANCHORARRAY calls. String literals are left alone. Marked calls
/// store the canonical uppercase name. A leading `=` is kept as given.
export function withFutureFunctionMarkers(formula: string): string {
  const stored = toStorageFormula(formula).text
  return formula.startsWith('=') ? `=${stored}` : stored
}
