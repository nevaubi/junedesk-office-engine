/** Seeger Weiss integration: all inference stays in this privileged process. */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { BedrockRuntimeClient, ConverseStreamCommand, type Message } from '@aws-sdk/client-bedrock-runtime'
import { defaultProvider } from '@aws-sdk/credential-provider-node'
import { SignatureV4 } from '@smithy/signature-v4'
import { Sha256 } from '@aws-crypto/sha256-js'
import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { awsRegion, gatewayUrl, safeLink, type WriterProfile } from '../shared/sw-policy'
import { compatibleMessages, converseMessages, packReasoning, parseToolCall, readSse } from './sw-conversation'

export interface InferenceProfile {
  api: 'converse' | 'mantle'; region: string; modelId: string;
  maxTokens?: number; requestFields?: Record<string, unknown>
}
export interface WriterConnection {
  version: 1; awsProfile?: string;
  profiles: Partial<Record<WriterProfile, InferenceProfile>>;
  research?: { gatewayUrl: string; region: string; toolName: string; queryKey?: string; arguments?: Record<string, unknown> }
}
export const CONFIG_EXAMPLE = {
  version: 1, awsProfile: '', profiles: {
    standard: { api: 'converse', region: 'us-east-1', modelId: '', maxTokens: 8192 },
    thorough: { api: 'mantle', region: 'us-east-1', modelId: '', maxTokens: 16384 },
  },
  research: { gatewayUrl: '', region: 'us-east-1', toolName: '', queryKey: 'query', arguments: {} },
}
export function connectionPath(userData: string): string { return process.env.SW_SHEETS_CONFIG || join(userData, 'sheets-connection.json') }
export function loadConnection(userData: string): WriterConnection {
  let raw: string
  try { raw = readFileSync(connectionPath(userData),'utf8') }
  catch { throw new Error('Connection configuration is missing. Open Connection to create the configuration file.') }
  if(raw.length>50000)throw new Error('Connection configuration is too large.')
  const x=JSON.parse(raw) as WriterConnection
  if(x.version!==1 || !x.profiles || typeof x.profiles!=='object')throw new Error('Invalid connection configuration. Expected version 1 and profiles.')
  return x
}
export function configuredProfile(c:WriterConnection,name:WriterProfile):InferenceProfile {
  const p=c.profiles[name]
  if(!p || !['converse','mantle'].includes(p.api) || typeof p.modelId!=='string' || !p.modelId.trim() || p.modelId.length>2048)
    throw new Error(`The ${name==='standard'?'Standard':'Thorough'} connection has not been configured.`)
  awsRegion(p.region)
  if(p.requestFields && (Array.isArray(p.requestFields)||typeof p.requestFields!=='object'||JSON.stringify(p.requestFields).length>16000))throw new Error('Invalid additional request settings.')
  return p
}
export interface WriterStreamCallbacks {
  onDelta(text:string):void; onToolCall(call:AgentToolCall):void; onReasoning(text:string):void;
  onStopReason(reason:string):void
}
function credentials(c:WriterConnection){return defaultProvider(c.awsProfile?{profile:c.awsProfile}:{})}
function outputCap(p:InferenceProfile){return Math.min(32768,Math.max(512,Number(p.maxTokens)||8192))}

export async function streamWriter(c:WriterConnection,name:WriterProfile,system:string,messages:AgentMessage[],tools:AgentToolDef[],cb:WriterStreamCallbacks,signal:AbortSignal):Promise<void>{
  const p=configuredProfile(c,name)
  if(p.api==='mantle')return streamMantle(c,p,system,messages,tools,cb,signal)
  const client=new BedrockRuntimeClient({region:p.region,credentials:credentials(c),maxAttempts:2,
    endpoint:`https://bedrock-runtime.${p.region}.amazonaws.com`})
  try{
    const response=await client.send(new ConverseStreamCommand({
      modelId:p.modelId, system:[{text:system}], messages:converseMessages(messages) as Message[],
      inferenceConfig:{maxTokens:outputCap(p)},
      ...(p.requestFields?{additionalModelRequestFields:p.requestFields as never}:{}),
      ...(tools.length?{toolConfig:{tools:tools.map(t=>({toolSpec:{name:t.name,description:t.description,inputSchema:{json:t.inputSchema as never}}}))}}:{}),
    }),{abortSignal:signal})
    if(!response.stream)throw new Error('The service returned no response stream.')
    const calls=new Map<number,{id:string;name:string;json:string}>();const reasoning=new Map<number,{text:string;signature:string;redacted:Buffer[]}>()
    let complete=false, truncated=false
    for await(const e of response.stream){
      if(signal.aborted)throw new DOMException('Stopped','AbortError')
      if(e.contentBlockStart?.start?.toolUse){const t=e.contentBlockStart.start.toolUse;calls.set(e.contentBlockStart.contentBlockIndex!,{id:t.toolUseId!,name:t.name!,json:''})}
      if(e.contentBlockDelta){const {contentBlockIndex,delta}=e.contentBlockDelta; if(delta?.text)cb.onDelta(delta.text)
        if(delta?.toolUse?.input){const t=calls.get(contentBlockIndex!);if(t){t.json+=delta.toolUse.input;if(t.json.length>1000000)throw new Error('Tool output exceeded the size limit.')}}
        if(delta?.reasoningContent){const r=reasoning.get(contentBlockIndex!)||{text:'',signature:'',redacted:[]}; const d=delta.reasoningContent as any;r.text+=d.text||'';r.signature+=d.signature||'';if(d.redactedContent)r.redacted.push(Buffer.from(d.redactedContent));reasoning.set(contentBlockIndex!,r)}
      }
      if(e.messageStop){complete=true;truncated=e.messageStop.stopReason==='max_tokens';cb.onStopReason(e.messageStop.stopReason||'end_turn')}
      if(e.internalServerException || e.modelStreamErrorException || e.validationException || e.throttlingException)throw new Error('The AWS stream reported a service error.')
    }
    if(!complete)throw new Error('The response ended before completion.')
    const rb=[...reasoning.values()].map(r=>r.redacted.length?{redactedContent:Buffer.concat(r.redacted).toString('base64')}:{reasoningText:{text:r.text,signature:r.signature}})
    if(rb.length)cb.onReasoning(packReasoning(rb))
    for(const t of calls.values())cb.onToolCall({...parseToolCall(t.id,t.name,t.json,tools.map(x=>x.name)),...(truncated?{truncated:true}:{})})
  } finally { client.destroy() }
}
async function streamMantle(c:WriterConnection,p:InferenceProfile,system:string,messages:AgentMessage[],tools:AgentToolDef[],cb:WriterStreamCallbacks,signal:AbortSignal):Promise<void>{
  const token=process.env.AWS_BEARER_TOKEN_BEDROCK
  if(!token)throw new Error('Set the approved Bedrock bearer token in the launching process for this endpoint.')
  const url=`https://bedrock-mantle.${p.region}.api.aws/v1/chat/completions`
  const response=await fetch(url,{method:'POST',redirect:'error',signal,
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({...p.requestFields,model:p.modelId,messages:compatibleMessages(system,messages),stream:true,max_tokens:outputCap(p),
      ...(tools.length?{tools:tools.map(t=>({type:'function',function:{name:t.name,description:t.description,parameters:t.inputSchema}}))}:{}),store:false}),
  })
  if(!response.ok){const e=new Error('The AWS request failed.') as Error&{$metadata:{httpStatusCode:number}};e.$metadata={httpStatusCode:response.status};throw e}
  const calls=new Map<number,{id:string;name:string;json:string}>();let complete=false, truncated=false
  await readSse(response,(data)=>{
    if(data==='[DONE]'){complete=true;return}if(!data.trim())return
    const e=JSON.parse(data);if(e.error)throw new Error('The service returned a stream error.')
    for(const choice of e.choices??[]){const d=choice.delta||{};if(typeof d.content==='string')cb.onDelta(d.content)
      if(typeof d.reasoning_content==='string')cb.onReasoning(d.reasoning_content)
      for(const t of d.tool_calls??[]){const old=calls.get(t.index)||{id:'',name:'',json:''};if(t.id)old.id=t.id;if(t.function?.name)old.name=t.function.name;old.json+=t.function?.arguments||'';if(old.json.length>1000000)throw new Error('Tool output exceeded the size limit.');calls.set(t.index,old)}
      if(choice.finish_reason){complete=true;truncated=choice.finish_reason==='length';cb.onStopReason(choice.finish_reason==='length'?'max_tokens':choice.finish_reason)}
    }
  },signal)
  if(!complete)throw new Error('The response ended before completion.')
  for(const t of calls.values())cb.onToolCall({...parseToolCall(t.id,t.name,t.json,tools.map(x=>x.name)),...(truncated?{truncated:true}:{})})
}

/** Signed, bounded MCP invocation. Only the exact configured AWS host is contacted. */
export async function searchPublic(c:WriterConnection,query:string,signal:AbortSignal):Promise<{results:{title:string;url:string;snippet:string}[];answer?:string;method:string}>{
  const r=c.research;if(!r?.toolName||!/^[\w.-]+$/.test(r.toolName))throw new Error('Configure the approved research tool name.')
  const url=gatewayUrl(r.gatewayUrl);const region=awsRegion(r.region)
  if(!url.hostname.includes(`.${region}.amazonaws.com`))throw new Error('The research gateway region does not match its endpoint.')
  const qkey=r.queryKey||'query';if(!/^[A-Za-z_][\w-]*$/.test(qkey))throw new Error('Invalid research query field.')
  const signer=new SignatureV4({service:'bedrock-agentcore',region,credentials:credentials(c),sha256:Sha256})
  let sessionId='';const protocol='2025-11-25'
  async function rpc(method:string,params:unknown,id?:string):Promise<any>{
    const body=JSON.stringify({jsonrpc:'2.0',...(id?{id}:{}),method,params})
    let headers:Record<string,string>={'content-type':'application/json',accept:'application/json, text/event-stream','mcp-protocol-version':protocol,host:url.host,...(sessionId?{'mcp-session-id':sessionId}:{})}
    if(process.env.SW_RESEARCH_TOKEN)headers.authorization=`Bearer ${process.env.SW_RESEARCH_TOKEN}`
    else{const signed=await signer.sign({method:'POST',protocol:'https:',hostname:url.hostname,path:url.pathname,headers,body});headers=signed.headers}
    const res=await fetch(url,{method:'POST',headers,body,redirect:'error',signal})
    if(!res.ok)throw new Error(`Research service returned HTTP ${res.status}.`)
    if(res.headers.get('mcp-session-id'))sessionId=res.headers.get('mcp-session-id')!
    if(!id)return null
    let value:any
    if(res.headers.get('content-type')?.includes('text/event-stream')){
      await readSse(res,line=>{if(line&&line!=='[DONE]'){const x=JSON.parse(line);if(x.id===id)value=x}},signal)
    }else{const text=await res.text();if(text.length>1000000)throw new Error('Research output exceeded the limit.');value=JSON.parse(text)}
    if(value?.error||!value?.result)throw new Error('The research gateway did not return a successful result.')
    return value.result
  }
  await rpc('initialize',{protocolVersion:protocol,capabilities:{},clientInfo:{name:'seeger-weiss-writer',version:'0.1.0'}},'init')
  await rpc('notifications/initialized',{})
  const result=await rpc('tools/call',{name:r.toolName,arguments:{...(r.arguments||{}),[qkey]:query}},'search')
  if(result.isError)throw new Error('The approved search tool reported an error.')
  const raw=(result.content||[]).filter((x:any)=>x.type==='text').map((x:any)=>x.text).join('\n').slice(0,24000)
  let data:any=result.structuredContent
  if(!data)try{data=JSON.parse(raw)}catch{data={}}
  const rows=Array.isArray(data)?data:data.results||data.search_results||data.items||[]
  const results=(Array.isArray(rows)?rows:[]).slice(0,20).map((x:any)=>({title:String(x.title||x.name||'Source').slice(0,300),url:safeLink(x.url||x.uri||x.link),snippet:String(x.snippet||x.content||x.text||'').slice(0,4000)})).filter((x:any)=>x.url) as {title:string;url:string;snippet:string}[]
  return{results,...(!results.length&&raw?{answer:raw}:{}),method:'agentcore'}
}
