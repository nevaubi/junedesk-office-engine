/** Local presentation library. Content is validated before document changes. */
import type { PageSpec, SpecElement } from '../main/page-spec'
export interface SlideStyle {id:string;name:string;background:string;ink:string;muted:string;accent:string;wash:string;line:string}
export const SLIDE_STYLES: SlideStyle[] = [
 {id:'firm',name:'Seeger Weiss',background:'#FFFFFF',ink:'#182838',muted:'#526574',accent:'#0F4A74',wash:'#EFF4F7',line:'#D5E1E8'},
 {id:'plain',name:'Plain report',background:'#FFFFFF',ink:'#24282D',muted:'#606A73',accent:'#485967',wash:'#F4F5F6',line:'#DCE0E4'},
 {id:'navy',name:'Dark blue',background:'#102B41',ink:'#FFFFFF',muted:'#C7D8E4',accent:'#8FBAD9',wash:'#1B3B53',line:'#416078'},
 {id:'warm',name:'Warm paper',background:'#FAF8F4',ink:'#292C30',muted:'#696D71',accent:'#68643F',wash:'#F0EDE4',line:'#D9D4C8'},
 {id:'slate',name:'Slate',background:'#F6F7F9',ink:'#18212C',muted:'#556375',accent:'#405773',wash:'#E9EDF2',line:'#D4DAE2'},
 {id:'ink',name:'Ink and white',background:'#FFFFFF',ink:'#17191C',muted:'#65676B',accent:'#17191C',wash:'#F2F2F2',line:'#D8D8D8'},
]
export const SLIDE_LAYOUTS = [
 {id:'cover',name:'Cover',description:'Title and optional subtitle'},
 {id:'section',name:'Section',description:'A quiet section divider'},
 {id:'content',name:'Key points',description:'One to three short sections'},
 {id:'two-column',name:'Two columns',description:'Two aligned content blocks'},
 {id:'three-column',name:'Three columns',description:'Three aligned content blocks'},
 {id:'comparison',name:'Comparison',description:'Two alternatives, side by side'},
 {id:'timeline',name:'Timeline',description:'One to five dated entries'},
 {id:'process',name:'Process',description:'One to five ordered steps'},
 {id:'closing',name:'Closing',description:'A final point or next step'},
] as const
export type LayoutId = typeof SLIDE_LAYOUTS[number]['id']
export interface TemplateSlide {layout:LayoutId;title:string;subtitle?:string;footer?:string;items?:{heading:string;text:string}[]}
export interface TemplateDeck {title:string;styleId:string;mode:'append'|'replace';slides:TemplateSlide[]}
const limits:Record<LayoutId,{min:number;max:number;text:number;heading:number}>={
 cover:{min:0,max:0,text:0,heading:0},section:{min:0,max:0,text:0,heading:0},closing:{min:0,max:0,text:0,heading:0},
 content:{min:1,max:3,text:280,heading:70},'two-column':{min:2,max:2,text:360,heading:70},
 'three-column':{min:3,max:3,text:220,heading:70},comparison:{min:2,max:2,text:360,heading:70},timeline:{min:1,max:5,text:160,heading:42},process:{min:1,max:5,text:160,heading:42},
}
function text(value:unknown,name:string,max:number,optional=false):string|undefined{
 if(value===undefined&&optional)return undefined
 if(typeof value!=='string'||!value.trim()||value.length>max||/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value))throw new Error(`${name} must contain 1-${max} characters.`)
 return value.trim()
}
function slideOf(raw:unknown):TemplateSlide{
 if(!raw||typeof raw!=='object')throw new Error('A slide definition is required.')
 const x=raw as Record<string,unknown>,layout=x.layout as LayoutId,l=limits[layout]
 if(!l)throw new Error('Choose a layout from the slide library.')
 if(x.items!==undefined&&!Array.isArray(x.items))throw new Error('Slide items must be an array.')
 const items=(x.items??[]) as unknown[]
 if(items.length<l.min||items.length>l.max)throw new Error(`${layout} requires ${l.min===l.max?l.min:`${l.min}-${l.max}`} items. Split excess content into another slide.`)
 return{layout,title:text(x.title,'Title',110)!,...(x.subtitle!==undefined?{subtitle:text(x.subtitle,'Subtitle',180)!}:{}),...(x.footer!==undefined?{footer:text(x.footer,'Footer',100)!}:{}),...(items.length?{items:items.map((item,i)=>{if(!item||typeof item!=='object')throw new Error('Invalid slide item.');const a=item as Record<string,unknown>;return{heading:text(a.heading,`Item ${i+1} heading`,l.heading)!,text:text(a.text,`Item ${i+1} text`,l.text)!}})}:{})}
}
export function validateDeck(raw:unknown):TemplateDeck{
 if(!raw||typeof raw!=='object')throw new Error('A presentation definition is required.')
 const x=raw as Record<string,unknown>
 if(!SLIDE_STYLES.some(s=>s.id===x.styleId))throw new Error('Choose a style from the library.')
 if(x.mode!==undefined&&x.mode!=='append'&&x.mode!=='replace')throw new Error('Use append or replace.')
 if(!Array.isArray(x.slides)||x.slides.length<1||x.slides.length>20)throw new Error('Create 1-20 slides at a time.')
 const d:TemplateDeck={title:text(x.title,'Presentation title',180)!,styleId:x.styleId as string,mode:x.mode==='replace'?'replace':'append',slides:x.slides.map(slideOf)}
 d.slides.forEach((s,i)=>{const issues=layoutCapacityWarnings(buildTemplateSlide(s,d.styleId,i+1,d.slides.length));if(issues.length)throw new Error(`Slide ${i+1}: ${issues[0]} Shorten this text or split it across slides.`)});return d
}
export function buildTemplateSlide(raw:TemplateSlide,styleId:string,page=1,total=1):PageSpec{
 const s=slideOf(raw),t=SLIDE_STYLES.find(s=>s.id===styleId);if(!t)throw new Error('Choose a style from the library.')
 const els:SpecElement[]=[]
 const rect=(x:number,y:number,w:number,h:number,fill:string,shape='rect')=>els.push({type:'shape',shape,x,y,w,h,fill})
 const tx=(value:string|undefined,x:number,y:number,w:number,h:number,size:number,color=t.ink,bold=false)=>{if(!value)return;els.push({type:'text',x,y,w,h,paragraphs:[{runs:[{text:value,sizePt:size,font:'Arial',bold,color}],lineSpacingPct:115}]})}
 rect(64,55,48,5,t.accent)
 const opening=['cover','section','closing'].includes(s.layout)
 if(opening){rect(1004,120,212,484,t.wash);tx(s.title,64,200,874,194,42,t.ink,true);tx(s.subtitle,68,428,850,138,23,t.muted)}
 else{
  tx(s.title,64,85,1152,105,32,t.ink,true);tx(s.subtitle,66,187,1148,58,18,t.muted)
  const top=s.subtitle?267:222,items=s.items!,n=items.length
  if(s.layout==='content'){
   const rowH=Math.min(176,(633-top)/n)
   items.forEach((a,i)=>{const y=top+i*rowH;rect(64,y,5,rowH-25,t.accent);tx(a.heading,88,y,300,70,21,t.accent,true);tx(a.text,424,y,785,rowH-24,20)})
  }else if(s.layout==='timeline'){
   const rowH=(640-top)/n;rect(77,top+8,2,rowH*(n-1)+22,t.line)
   items.forEach((a,i)=>{const y=top+i*rowH;rect(68,y+3,20,20,t.accent,'ellipse');tx(a.heading,108,y-3,265,rowH-10,19,t.accent,true);tx(a.text,410,y-3,799,rowH-9,19)})
  }else if(s.layout==='process'){
   const cols=n<=3?n:3,rows=n<=3?1:2,gap=24,w=(1152-gap*(cols-1))/cols,h=(638-top-gap*(rows-1))/rows
   items.forEach((a,i)=>{const x=64+(i%cols)*(w+gap),y=top+Math.floor(i/cols)*(h+gap);rect(x,y,w,h,t.wash);tx(String(i+1).padStart(2,'0'),x+22,y+17,45,32,14,t.accent,true);tx(a.heading,x+75,y+17,w-99,rows===1?100:57,rows===1?21:19,t.ink,true);tx(a.text,x+22,y+(rows===1?153:85),w-44,h-(rows===1?168:99),rows===1?18:17)})
  }else{
   const gap=28,w=(1152-gap*(n-1))/n,h=638-top
   items.forEach((a,i)=>{const x=64+i*(w+gap);rect(x,top,w,h,t.wash);rect(x,top,w,4,t.accent);tx(a.heading,x+26,top+24,w-52,100,n===3?22:24,t.accent,true);tx(a.text,x+26,top+141,w-52,h-164,n===3?19:20)})
  }
 }
 rect(64,665,1152,1,t.line);tx(s.footer,64,675,1020,30,10,t.muted);tx(`${page} / ${total}`,1130,675,86,30,10,t.muted)
 return{background:t.background,elements:els}
}
export function exampleSlide(layout:LayoutId):TemplateSlide{
 const counts:Partial<Record<LayoutId,number>>={'content':3,'two-column':2,'three-column':3,'comparison':2,'timeline':4,'process':4}
 const n=counts[layout]??0
 return{layout,title:layout==='cover'?'Presentation title':layout==='closing'?'Next steps':layout==='section'?'Section title':SLIDE_LAYOUTS.find(x=>x.id===layout)?.name??'Slide title',...(n?{items:Array.from({length:n},(_,i)=>({heading:layout==='timeline'?`Date ${i+1}`:layout==='process'?`Step ${i+1}`:`Heading ${i+1}`,text:'Replace this text with your own concise, supported content.'}))}:{subtitle:'Add a short subtitle or leave this area blank.'})}
}
/** Conservative text estimate; final render-layer audits use actual font metrics. */
export function layoutCapacityWarnings(spec:PageSpec):string[]{
 const issues:string[]=[]
 for(const el of spec.elements){if(el.type!=='text')continue;const run=el.paragraphs[0]?.runs[0];if(!run)continue;const px=(run.sizePt??18)*96/72;const width=el.w-20,height=el.h-12;let lines=0;
 for(const para of run.text.split('\n')){let used=0;lines++;for(const word of para.split(/\s+/)){const w=[...word].reduce((n,c)=>n+px*(/[il.,:;!'|]/.test(c)?.25:/[MW@#]/.test(c)?.87:/[^\x00-\x7f]/.test(c)?1:.56),0);if(w>width){issues.push('A word is wider than its text box.');break}if(used&&used+w+px*.3>width){lines++;used=w}else used+=(used?px*.3:0)+w}}
 if(lines*px*1.15>height+5)issues.push(`Text may exceed its frame: "${run.text.slice(0,35)}".`)
 }return issues
}
