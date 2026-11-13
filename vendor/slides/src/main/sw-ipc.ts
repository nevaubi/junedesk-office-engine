/** Isolated spreadsheet-app IPC. Stock cloud routes are deliberately not registered. */
import { app, ipcMain, shell, type IpcMainInvokeEvent } from 'electron'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { AiSettings, AiStreamRequest, AiStreamChunk } from '@genoffice/ai-provider'
import { filterTools, modeName, profileName, publicPreferences, publicQuery, writerError, type WriterStatus } from '../shared/sw-policy'
import { validateConversation, publicResearchRequest } from './sw-conversation'
import { CONFIG_EXAMPLE, connectionPath, configuredProfile, loadConnection, searchPublic, streamWriter } from './sw-connection'

export function registerCompanionAiIpc(rendererFile:()=>string, rendererUrl:()=>string|undefined):void{
  ipcMain.handle('app:get-theme', (e) => { trust(e); return 'light' })
  const streams=new Map<string,AbortController>()
  const research=new Map<number,{query:string; expires:number; result?:ReturnType<typeof searchPublic>}>()
  const tested=new Map<string,string>()
  function trust(e:IpcMainInvokeEvent){
    const u=e.senderFrame?.url||''
    if(e.senderFrame!==e.sender.mainFrame)throw new Error('This operation is only available to the presentation window.')
    try{if(u.startsWith('file:')&&resolve(fileURLToPath(u))===resolve(rendererFile()))return
      const dev=rendererUrl();if(dev&&new URL(u).origin===new URL(dev).origin)return
    }catch{/* invalid source */}
    throw new Error('Untrusted window.')
  }
  const prefsPath=()=>join(app.getPath('userData'),'slides-preferences.json')
  function readPrefs(){try{return publicPreferences(JSON.parse(readFileSync(prefsPath(),'utf8')))}catch{return publicPreferences({})}}
  function settings():AiSettings{
    return {...readPrefs(),provider:'custom',providers:{custom:{apiKey:'',model:''}},gskToolsEnabled:false} as unknown as AiSettings
  }
  function status():WriterStatus{
    const configPath=connectionPath(app.getPath('userData'))
    try{const c=loadConnection(app.getPath('userData'));let configured=false;try{configuredProfile(c,'standard');configured=true}catch{}
      return{configured,researchConfigured:false,configPath,
        message:configured?'Connection configured. Use Test connection to check access.':'Set the Standard connection before using the writing assistant.',
        testedAt:tested.get(JSON.stringify(c))}}
    catch{return{configured:false,researchConfigured:false,configPath,message:'Local editing is ready. The writing connection has not been configured.'}}
  }
  ipcMain.handle('ai:get-settings',(e)=>{trust(e);return settings()})
  ipcMain.handle('ai:set-settings',(e,value)=>{trust(e);mkdirSync(dirname(prefsPath()),{recursive:true});writeFileSync(prefsPath(),JSON.stringify(publicPreferences(value),null,2),'utf8')})
  ipcMain.handle('sw:status',(e)=>{trust(e);return status()})
  ipcMain.handle('sw:open-config',async(e)=>{
    trust(e);const file=connectionPath(app.getPath('userData'));mkdirSync(dirname(file),{recursive:true})
    if(!existsSync(file))writeFileSync(file,JSON.stringify(CONFIG_EXAMPLE,null,2),'utf8')
    const error=await shell.openPath(file);if(error)throw new Error('The configuration could not be opened. Open the displayed file path in a text editor.');return status()
  })
  ipcMain.handle('sw:test',async(e)=>{
    trust(e);const controller=new AbortController();const timer=setTimeout(()=>controller.abort(),60000);let testedKey:string|undefined
    try{const c=loadConnection(app.getPath('userData'));testedKey=JSON.stringify(c);let text=''
      await streamWriter(c,'standard','Reply with the single word READY.',[{role:'user',text:'Connection test.'}],[],{onDelta:t=>{text+=t},onReasoning:()=>{},onToolCall:()=>{},onStopReason:()=>{}},controller.signal)
      if(!text.trim())throw new Error('No test response.')
      tested.set(JSON.stringify(c),new Date().toISOString());return{...status(),message:'Connection tested successfully. No presentation content was sent.'}
    }catch(err){if(testedKey)tested.delete(testedKey);return{...status(),message:writerError(err)}}finally{clearTimeout(timer)}
  })
  ipcMain.handle('sw:approve-research',(e)=>{trust(e);throw new Error('Research is not enabled in this companion yet.')})
  // Kept only as inert compatibility endpoints for inherited error helpers.
  ipcMain.handle('ai:gsk-status',(e)=>{trust(e);return{loggedIn:false}})
  ipcMain.handle('ai:gsk-login',(e)=>{trust(e);throw new Error('External account sign-in is not part of this application.')})
  ipcMain.handle('ai:stream',async(e,request:AiStreamRequest)=>{
    trust(e)
    const id=request?.requestId
    if(typeof id!=='string'||!/^[A-Za-z0-9:_-]{1,128}$/.test(id))throw new Error('Invalid request identifier.')
    const key=`${e.sender.id}:${id}`;const send=(chunk:Omit<AiStreamChunk,'requestId'>)=>{if(!e.sender.isDestroyed())e.sender.send('ai:stream-chunk',{requestId:id,...chunk})}
    if(streams.has(key)){send({type:'error',error:'This request is already running.'});return}
    const controller=new AbortController();streams.set(key,controller)
    const closed=()=>{controller.abort();research.delete(e.sender.id)};e.sender.once('destroyed',closed)
    let timedOut=false;const deadline=setTimeout(()=>{timedOut=true;controller.abort()},180000),ping=setInterval(()=>send({type:'ping'}),5000)
    try{
      validateConversation(request)
      const prefs=publicPreferences(request.settings),mode=modeName(prefs.swMode)
      if(mode==='research')throw new Error('Research is not enabled in this companion yet.')
      const c=loadConnection(app.getPath('userData'));let stopReason:string|undefined
      let payload={system:request.system,messages:request.messages,tools:filterTools(request.tools||[],mode)}

      await streamWriter(c,profileName(prefs.swProfile),payload.system,payload.messages,payload.tools,{
        onDelta:text=>send({type:'delta',text}),onReasoning:text=>send({type:'reasoning',text}),onToolCall:toolCall=>send({type:'tool-call',toolCall}),onStopReason:r=>{stopReason=r},
      },controller.signal)
      send(stopReason ? {type:'done',stopReason} : {type:'done'})
    }catch(error){if(timedOut)send({type:'error',error:'The request exceeded the time limit. Completed presentation changes can be restored from the response.',errorCode:'timeout'});else if(controller.signal.aborted)send({type:'done'});else send({type:'error',error:writerError(error)})}
    finally{clearTimeout(deadline);clearInterval(ping);streams.delete(key);e.sender.removeListener('destroyed',closed)}
  })
  ipcMain.handle('ai:stream-cancel',(e,id)=>{trust(e);if(typeof id==='string')streams.get(`${e.sender.id}:${id}`)?.abort()})
  ipcMain.handle('ai:chat',async(e,request)=>{
    trust(e);const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),120000)
    try{if(modeName(request?.settings?.swMode)==='research')throw new Error('Research uses only the approved public question.');if(typeof request?.system!=='string'||typeof request?.user!=='string'||request.system.length+request.user.length>300000)throw new Error('Invalid request.')
      const c=loadConnection(app.getPath('userData'));let content=''
      await streamWriter(c,profileName(request.settings?.swProfile),request.system,[{role:'user',text:request.user}],[],{onDelta:t=>{content+=t},onReasoning:()=>{},onToolCall:()=>{},onStopReason:()=>{}},controller.signal)
      return {ok:true,content}
    }catch(error){return{ok:false,error:writerError(error)}}finally{clearTimeout(timer)}
  })
  ipcMain.handle('ai:web-search',(e)=>{trust(e);return{results:[],method:'error',error:'Research is not enabled yet. No external fallback was used.'}})
  ipcMain.handle('ai:image-search',(e)=>{trust(e);return{images:[],method:'error',error:'Remote image search is disabled. Insert an image from your computer.'}})
  ipcMain.handle('ai:fetch-image',(e)=>{trust(e);return null})
  ipcMain.handle('sw:disabled-image-generation',(e)=>{trust(e);return{error:'Remote image generation is disabled in this document companion.'}})
}
