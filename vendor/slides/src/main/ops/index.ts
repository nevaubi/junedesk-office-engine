// Modified for lit-ai: registers the deck selector ops (deck-ops.ts) and exports pinBatchRefs.
import './core-ops'
import './text-ops'
import './element-ops'
import './insert-ops'
import './table-ops'
import './slide-ops'
import './arrange-ops'
import './animation-ops'
import './equation-ops'
import './deck-ops'
export { pinBatchRefs, runTxn, type TxnRequest, type TxnResult, type OpFailure } from './executor'
export {
  elementDurableId,
  GuidedError,
  opNames,
  register,
  resolveGroupChildId,
  slideDurableId,
  type Op,
  type OpRecord,
  type OpTarget,
} from './registry'
