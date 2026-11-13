/** Seeger Weiss local style storage and controlled desktop connection.
 * Modified from the Apache-licensed source; stock external service routes are not registered.
 */
import {app,ipcMain,type IpcMainInvokeEvent} from 'electron'
import {existsSync,mkdirSync,readdirSync,readFileSync,writeFileSync} from 'node:fs'
import {join,resolve} from 'node:path'
import {fileURLToPath} from 'node:url'
import {runtime,sessions} from './session-state'
import {registerCompanionAiIpc} from './sw-ipc'
function trust(e:IpcMainInvokeEvent){if(e.senderFrame!==e.sender.mainFrame)throw new Error('Untrusted window.');const u=e.senderFrame?.url;try{if(u?.startsWith('file:')&&resolve(fileURLToPath(u))===resolve(runtime.rendererFilePath!))return;if(runtime.rendererDevUrl&&u&&new URL(u).origin===new URL(runtime.rendererDevUrl).origin)return}catch{}throw new Error('Untrusted window.')}
function filename(value:unknown){if(typeof value!=='string'||!value.trim()||value.length>64||/[\\/\x00-\x1f:*?"<>|]/.test(value)||/^(con|prn|aux|nul|com\d|lpt\d)$/i.test(value)||value.endsWith('.'))throw new Error('Choose a short template name without file-path characters.');return value.trim()+'.json'}
function dataOf(raw:unknown){if(!raw||typeof raw!=='object')throw new Error('Invalid template.');const d=raw as Record<string,unknown>;if(typeof d.styleSkill!=='string'||d.styleSkill.length>40000||typeof d.topic!=='string'||d.topic.length>300)throw new Error('Template text is missing or too long.');return{topic:d.topic,styleSkill:d.styleSkill,createdAt:new Date().toISOString()}}
const dir=()=>join(app.getPath('userData'),'style-templates')
export function registerAiIpc(){registerCompanionAiIpc(()=>runtime.rendererFilePath!,()=>runtime.rendererDevUrl)}
export function registerSlidesOnlyAiIpc(){
 const blocked=(e:IpcMainInvokeEvent)=>{trust(e);return{ok:false,error:'External generation is not enabled. Use the local presentation library or insert a local file.'}}
 for(const name of ['ai:generate-image','ai:analyze-media','ai:insert-image-url','ai:replace-picture-url'])ipcMain.handle(name,blocked)
 ipcMain.handle('ai:log-run-failure',e=>{trust(e);return null})
 // Style instructions live in application data, not beside the user's presentation.
 ipcMain.handle('ai:save-sidecar',(e,raw)=>{trust(e);const session=sessions.get(e.sender.id);if(!session)return{ok:false};const d=dataOf(raw);mkdirSync(dir(),{recursive:true});writeFileSync(join(dir(),'last-style.json'),JSON.stringify(d),'utf8');return{ok:true}})
 ipcMain.handle('ai:save-style-template',(e,name,raw)=>{trust(e);try{const file=filename(name),data=dataOf(raw);mkdirSync(dir(),{recursive:true});writeFileSync(join(dir(),file),JSON.stringify({...data,name:name.trim()}),'utf8');return{ok:true}}catch(error){return{ok:false,error:error instanceof Error?error.message:'Template could not be saved.'}}})
 ipcMain.handle('ai:list-style-templates',e=>{trust(e);if(!existsSync(dir()))return[];return readdirSync(dir()).filter(f=>f.endsWith('.json')).slice(0,200).flatMap(f=>{try{const d=JSON.parse(readFileSync(join(dir(),f),'utf8'));return[{name:d.name??f.slice(0,-5),topic:d.topic??'',createdAt:d.createdAt??''}]}catch{return[]}})})
 ipcMain.handle('ai:load-style-template',(e,name)=>{trust(e);try{const d=dataOf(JSON.parse(readFileSync(join(dir(),filename(name)),'utf8')));return{ok:true,...d}}catch{return{ok:false,error:'The local style template could not be read.'}}})
}
