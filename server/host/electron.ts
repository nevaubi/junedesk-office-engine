/** Node host for the retained document-operation registry. No Electron runtime. */
import {EventEmitter} from 'node:events';
import {workerData} from 'node:worker_threads';
import {mkdirSync} from 'node:fs';
import {join} from 'node:path';
export const handlers = new Map<string, (...args:any[])=>any>();
export const events: Array<{channel:string;args:any[]}> = [];
const unsupported = () => { throw new Error('This desktop-only operation is unavailable in the browser host.'); };
const root = workerData?.root ?? process.cwd();
const scratch = workerData?.scratch ?? join(root,'.local/worker-test');
export const app:any = Object.assign(new EventEmitter(), {
 getPath:(name:string)=>{const p=join(scratch,name==='temp'?'temp':'state');mkdirSync(p,{recursive:true});return p;},
 getAppPath:()=>root, getLocale:()=> 'en-US', getSystemLocale:()=> 'en-US',
 getVersion:()=> 'web-preview.1', getName:()=> 'Seeger Weiss Office', isPackaged:false,
 whenReady:()=>Promise.resolve(), isReady:()=>true, getAppMetrics:()=>[],
});
export const sender:any = Object.assign(new EventEmitter(), {
 id:1, isDestroyed:()=>false, getURL:()=> 'https://office.invalid/editor',
 send:(channel:string,...args:any[])=>events.push({channel,args}),
 executeJavaScript:unsupported, focus:()=>{}, isFocused:()=>true,
});
sender.mainFrame={url:sender.getURL()};
export const ipcMain:any = Object.assign(new EventEmitter(), {
 handle:(channel:string,fn:any)=>{if(handlers.has(channel))throw Error('Duplicate engine channel: '+channel);handlers.set(channel,fn);},
 removeHandler:(channel:string)=>handlers.delete(channel),
});
export const webContents:any={fromId:(id:number)=>id===1?sender:null,getAllWebContents:()=>[sender]};
export class BrowserWindow { constructor(){unsupported();} static fromWebContents(){return null;} static getAllWindows(){return [];} static getFocusedWindow(){return null;} }
export class WebContentsView { constructor(){unsupported();} }
let saveTarget:string|null=null;
export function permitSaveTarget(path:string){saveTarget=path;}
export const dialog:any={showOpenDialog:unsupported,showMessageBox:unsupported,showErrorBox:unsupported,
 showSaveDialog:async()=>{if(!saveTarget)unsupported();const p=saveTarget;saveTarget=null;return {canceled:false,filePath:p};}};
export const shell:any={openExternal:unsupported,openPath:unsupported,showItemInFolder:unsupported};
export const Menu:any={buildFromTemplate:(template:any)=>({items:template}),getApplicationMenu:()=>null,setApplicationMenu:()=>{}};
export const screen:any={getAllDisplays:()=>[],getPrimaryDisplay:()=>({bounds:{x:0,y:0,width:1280,height:900}})};
export const desktopCapturer:any={getSources:unsupported};
export const systemPreferences:any={getMediaAccessStatus:()=> 'denied',askForMediaAccess:async()=>false};
export const session:any={defaultSession:{setDisplayMediaRequestHandler:unsupported}};
export const net:any={fetch:unsupported};
export const nativeTheme:any=Object.assign(new EventEmitter(),{shouldUseDarkColors:false});
const clipboardData=new Map<string,any>();
export const clipboard:any={write:(v:any)=>{clipboardData.clear();for(const[k,x]of Object.entries(v))clipboardData.set(k,x);},
 writeBuffer:(k:string,v:any)=>clipboardData.set(k,v),readBuffer:(k:string)=>clipboardData.get(k)??Buffer.alloc(0),
 writeText:(v:string)=>clipboardData.set('text',v),readText:()=>clipboardData.get('text')??'',
 readHTML:()=>clipboardData.get('html')??'',writeHTML:(s:string)=>clipboardData.set('html',s),
 availableFormats:()=>[...clipboardData.keys()],clear:()=>clipboardData.clear(),
 readImage:unsupported,writeImage:unsupported};
export const nativeImage:any={createEmpty:()=>({isEmpty:()=>true}),createFromBuffer:unsupported,createFromDataURL:unsupported,createFromPath:unsupported};
export const webUtils:any={getPathForFile:unsupported};
export const contextBridge:any={exposeInMainWorld:unsupported};
export async function invokeNative(channel:string,args:any[]=[]){
 const fn=handlers.get(channel);if(!fn)throw Error('The retained engine does not provide '+channel);
 return await fn({sender,senderFrame:sender.mainFrame,reply:(c:string,...a:any[])=>sender.send(c,...a)},...args);
}
export async function drainEvents(){await new Promise<void>(r=>setImmediate(r));return events.splice(0);}
// Renderer-side bridge functions may be imported by a shared barrel, but are never executed in a Node document worker.
export const ipcRenderer:any={invoke:unsupported,send:unsupported,on:unsupported,removeListener:unsupported};
