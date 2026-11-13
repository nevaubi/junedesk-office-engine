import {sessions} from '../../vendor/slides/src/main/session-state';
import {savePptx} from '@genoffice/pptx-engine';
import {parentPort,workerData} from 'node:worker_threads';
import {registerSlidesIpc,sessionDirty} from '../../vendor/slides/src/main/slides-main';
import {sender,invokeNative,drainEvents,permitSaveTarget,handlers} from './electron';
import {operationKind,safeValue,mapDocumentPaths,recoveryResult} from '../boundary.mjs';
const {app,documentPath,documentToken}=workerData;
globalThis.fetch=async()=>{throw Error('External network access is disabled in the document engine.');};
registerSlidesIpc();
let initial:any;
const ready=invokeNative('slides:open-path',[documentPath,1100]).then(r=>{initial=r;});
function scrub(value:any):any{
 if(value instanceof Uint8Array||value instanceof ArrayBuffer)return value;
 if(typeof value==='string'&&value===documentPath)return documentToken;
 if(Array.isArray(value))return value.map(scrub);
 if(value&&typeof value==='object'){const result:any={};for(const[k,v]of Object.entries(value))result[k]=scrub(v);return result;}
 return value;
}
let chain=Promise.resolve();
parentPort!.on('message',message=>{chain=chain.then(async()=>{
 try{
  await ready;let result:any;
  if(message.channel==='host:close'){sender.emit('destroyed');result={closed:true};}
  else if(message.channel==='host:open')result=initial;
  // Unsaved work as a package, for the platform to keep as the recovery copy. Nothing is written to the document.
  else if(message.channel==='slides:recovery-bytes'){operationKind(app,message.channel);const current=sessions.get(sender.id);if(!current)throw Error('The active presentation is unavailable.');result=sessionDirty(current)?recoveryResult(new Uint8Array(await savePptx(current.opened))):{found:false,dirty:false};}
  else if(message.channel==='host:bytes'){const current=sessions.get(sender.id);if(!current)throw Error('The active presentation is unavailable.');result=new Uint8Array(await savePptx(current.opened));}
  else{
   operationKind(app,message.channel);safeValue(message.args);
   const args=mapDocumentPaths(message.args,documentToken,documentPath);
   const current=sessions.get(sender.id);if(current)current.path=documentPath;
   result=await invokeNative(message.channel,args);
   const updated=sessions.get(sender.id);if(updated)updated.path=documentPath;
   if(result&&typeof result==='object'&&'path' in result)result.path=documentPath;
  }
  parentPort!.postMessage({id:message.id,ok:true,result:scrub(result),events:scrub(await drainEvents())});
 }catch(error:any){parentPort!.postMessage({id:message.id,ok:false,error:error?.message||'Document operation failed.'});}
 }).catch(()=>{});});
