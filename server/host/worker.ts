import {parentPort,workerData} from 'node:worker_threads';
import {readFile} from 'node:fs/promises';
import {registerSheetsWebHost,stopSheetsSidecar,readWorkbookRecoveryCopy} from '../../vendor/sheets/src/main/sheets-main';
import {sender,invokeNative,drainEvents,permitSaveTarget,handlers} from './electron';
import {operationKind,safeValue,mapDocumentPaths,recoveryResult} from '../boundary.mjs';
const {app,documentPath,documentToken}=workerData;
globalThis.fetch=async()=>{throw Error('External network access is disabled in the document engine.');};
registerSheetsWebHost(sender,documentPath,workerData.binaryPath);
parentPort!.on('close',()=>stopSheetsSidecar());
let initial:any;
const ready=invokeNative('workbook:select',[]).then(r=>{initial=r;});
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
  if(message.channel==='host:close'){sender.emit('destroyed');stopSheetsSidecar();result={closed:true};}
  else if(message.channel==='host:open')result=initial;
  else if(message.channel==='host:bytes')result=new Uint8Array(await readFile(documentPath));
  // The copy written by workbook:write-recovery, for the platform to keep as the unsaved-work recovery copy.
  else if(message.channel==='workbook:recovery-bytes'){operationKind(app,message.channel);result=recoveryResult(await readWorkbookRecoveryCopy(documentPath));}
  else{
   if(message.channel==='workbook:save'&&message.args[0]?.mode==='save-as')permitSaveTarget(documentPath);
   operationKind(app,message.channel);safeValue(message.args);
   const args=mapDocumentPaths(message.args,documentToken,documentPath);
   result=await invokeNative(message.channel,args);
  }
  parentPort!.postMessage({id:message.id,ok:true,result:scrub(result),events:scrub(await drainEvents())});
 }catch(error:any){parentPort!.postMessage({id:message.id,ok:false,error:error?.message||'Document operation failed.'});}
 }).catch(()=>{});});
