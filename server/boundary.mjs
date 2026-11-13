import {reject} from './auth.mjs';
const words = s => new Set(s.trim().split(/\s+/));
const read = {
 slides: words('read-slide read-native-details get-render-slides get-slide-size get-layouts private-font-faces private-font-data font-catalog font-missing get-chart-data chart-color-schemes get-link get-slide-links get-run-links get-header-footer get-transition get-animations get-shape-keys get-sections get-notes get-comments has-slide-clipboard clipboard-probe is-dirty master-enter master-open master-close recovery-bytes'),
 sheets: words('read-range read-formulas read-media read-pivot-definition recalc recovery-bytes'),
 pdf: words('read-file validate-text-edits list-edit-fonts can-draw-text list-page-images list-static-form-fills page-image-png page-preview-png is-untitled get-username')
};
const write = {
 slides: words('edit-text set-element-font set-element-paragraph-format find-replace set-slide-layout set-slide-size edit-transform edit-connector-endpoints edit-picture-src-rect edit-picture-opacity edit-image-fill change-shape set-shape-adjust set-text-anchor set-text-body-props set-effects group-elements ungroup-element batch-edit-transform add-element delete-element add-slide add-blank-slide add-slide-with-layout edit-fill edit-stroke flip-elements edit-background copy-slide paste-slide repaste-slide delete-slide reorder-element edit-table-cell table-structure table-merge set-table-col-width set-table-row-height set-table-cell-anchor edit-table-style edit-chart copy-elements paste-elements duplicate-elements add-table add-ink add-chart add-smartart add-image-bytes replace-picture-bytes set-link apply-header-footer apply-theme set-transition set-advance-times set-animations set-hidden set-sections add-section rename-section remove-section move-section move-slide set-notes add-comment delete-comment history-batch-begin history-batch-end apply-edit-script apply-txn ai-snapshot-restore undo redo master-edit-text master-edit-transform master-edit-fill master-edit-stroke master-delete-element local-page-generate land-generated-pages'),
 sheets: words('save-edits-begin save-edits-chunk save-edits-abort write-recovery'),
 pdf: words('insert-blank-page set-page-size crop-pages merge-pages split-pages')
};
const prefix={slides:'slides:',sheets:'workbook:',pdf:'pdf:'};
// Largest unsaved-work package handed to the platform as a recovery copy (matches the platform's inline recovery limit).
export const RECOVERY_MAX_BYTES = 3 * 1024 * 1024;
// Shape of a recovery-bytes result: the package only travels when it fits; larger work reports its size so the browser can say so.
export function recoveryResult(bytes, dirty = true) {
 if (!bytes) return {found:false,dirty:false};
 return bytes.byteLength > RECOVERY_MAX_BYTES ? {found:true,dirty,size:bytes.byteLength} : {found:true,dirty,size:bytes.byteLength,bytes};
}
export function operationKind(app, channel) {
 if (!prefix[app] || typeof channel !== 'string' || !channel.startsWith(prefix[app])) reject(403,'This document operation is not available.');
 const suffix=channel.slice(prefix[app].length);
 if (read[app].has(suffix)) return 'read';
 if (write[app].has(suffix)) return 'write';
 if (suffix==='save') return 'save';
 reject(403,'This document operation is not available in the browser edition.');
}
export function safeValue(value) {
 const json=JSON.stringify(value); if(!json || Buffer.byteLength(json)>24*1024*1024) reject(413,'Operation exceeds the size limit.');
 const walk=(v,depth)=>{if(depth>64)reject(422,'Operation nesting exceeds the limit.');if(v&&typeof v==='object')for(const k of Object.keys(v)){if(['__proto__','prototype','constructor'].includes(k))reject(422,'Unsafe object property.');walk(v[k],depth+1);}};
 walk(value,0); return value;
}
const pathKeys = new Set(['path','targetPath','sourcePath','filePath','outPath','outputPath','folder','directory','paths']);
export function mapDocumentPaths(value, token, actual, depth=0) {
 if(value instanceof ArrayBuffer||value instanceof Uint8Array)return value;
 if(value===token)return actual;
 if(depth>64)reject(422,'Operation nesting is too deep.');
 if(Array.isArray(value))return value.map(v=>mapDocumentPaths(v,token,actual,depth+1));
 if(!value || typeof value!=='object')return value;
 const result={};
 for(const [key,item] of Object.entries(value)) {
  if(['__proto__','constructor','prototype'].includes(key))reject(422,'Unsafe object property.');
  if(pathKeys.has(key)&&item!==undefined&&item!==null){
   if(item!==token)reject(403,'A file path outside the current document is not authorized.');
   result[key]=actual;
  }else result[key]=mapDocumentPaths(item,token,actual,depth+1);
 }
 return result;
}
