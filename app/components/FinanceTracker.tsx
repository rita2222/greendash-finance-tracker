//@ts-nocheck
import { useState, useEffect, useMemo, useRef } from "react"
import * as XLSX from 'xlsx'
import { createClient } from '@supabase/supabase-js'

const supabase = createClient(process.env.NEXT_PUBLIC_SUPABASE_URL || '', process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || '')

async function storageGet(key) {
  try {
    const { data, error } = await supabase.from('finance_data').select('value').eq('key', key).single()
    if (error && error.code !== 'PGRST116') {
      console.warn('⚠️ Supabase GET error:', error)
      const local = localStorage.getItem(key)
      return local ? { key, value: local } : null
    }
    return error ? null : { key, value: data?.value }
  } catch (e) { 
    console.error('❌ Supabase GET exception:', e)
    const local = localStorage.getItem(key)
    return local ? { key, value: local } : null
  }
}

async function storageSet(key, value) {
  try {
    await supabase.from('finance_data').delete().eq('key', key)
    const { error } = await supabase.from('finance_data').insert([{ key, value }])
    if (error) {
      console.warn('⚠️ Supabase SET error:', error)
      localStorage.setItem(key, value)
      return true
    }
    localStorage.setItem(key, value)
    return true
  } catch (e) { 
    console.error('❌ Supabase SET exception:', e)
    localStorage.setItem(key, value)
    return true
  }
}


const MONTHS=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
const today=new Date()
const todayStr=today.toISOString().slice(0,10)
const fmt=n=>`€${Number(n).toLocaleString('pt-PT',{minimumFractionDigits:2,maximumFractionDigits:2})}`
const fmtD=d=>d?new Date(d+'T12:00:00').toLocaleDateString('pt-PT'):'—'
const ym=d=>d?d.slice(0,7):''
const nowYM=`${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}`
const fmtBytes=b=>b<1024?`${b}B`:b<1048576?`${(b/1024).toFixed(1)}KB`:`${(b/1048576).toFixed(1)}MB`
const FICON=t=>!t?'📎':t.startsWith('image/')?'🖼️':t==='application/pdf'?'📕':t.includes('word')?'📝':t.includes('excel')||t.includes('sheet')?'📊':'📎';const openFile=a=>{const w=window.open();if(w)w.document.write(`<title>${a.name}</title><body style="margin:0"><iframe src="data:${a.mime};base64,${a.data}" style="border:none;width:100%;height:100vh"></iframe></body>`)}
const pct=(a,b)=>b===0?0:Math.min(100,(a/b)*100)
const ivaAmt=(amount,rate)=>Number(amount)*(Number(rate||0)/100)
const qDeadline=(q,y)=>{const mo=[4,7,10,1][q-1];return new Date(q===4?y+1:y,mo,20)}
const getLines=e=>e.ivaLines?.length?e.ivaLines:[{base:Number(e.amount||0),rate:Number(e.ivaRate||0)}]
const linesIva=ls=>ls.reduce((s,l)=>s+Number(l.base||0)*(Number(l.rate||0)/100),0)
const linesBase=ls=>ls.reduce((s,l)=>s+Number(l.base||0),0)
const grantBase=ge=>Number(ge.ivaRate||0)>0?Number(ge.amount||0)/(1+Number(ge.ivaRate)/100):Number(ge.amount||0)
const grantReim=ge=>grantBase(ge)*0.75
const dlAll=files=>{
  const f=files.filter(Boolean)
  if(!f.length)return
  f.forEach((a,i)=>setTimeout(()=>{const l=document.createElement('a');l.href=`data:${a.mime};base64,${a.data}`;l.download=a.name;document.body.appendChild(l);l.click();document.body.removeChild(l)},i*350))
}

const FREQUENCIES=[
  {v:'monthly',l:'Mensal',months:1},
  {v:'quarterly',l:'Trimestral',months:3},
  {v:'biannual',l:'Semestral',months:6},
  {v:'annual',l:'Anual',months:12},
]
const advanceDue=fe=>{
  if(fe.type!=='recurring')return fe
  const months=FREQUENCIES.find(f=>f.v===fe.frequency)?.months||1
  const d=new Date(fe.nextDue+'T12:00:00')
  d.setMonth(d.getMonth()+months)
  const maxDay=new Date(d.getFullYear(),d.getMonth()+1,0).getDate()
  d.setDate(Math.min(Number(fe.dayOfMonth)||1,maxDay))
  return{...fe,nextDue:d.toISOString().slice(0,10)}
}
const computeNextDue=(frequency,dayOfMonth)=>{
  const months=FREQUENCIES.find(f=>f.v===frequency)?.months||1
  const d=new Date(today.getFullYear(),today.getMonth(),Number(dayOfMonth)||1)
  if(d.toISOString().slice(0,10)<=todayStr)d.setMonth(d.getMonth()+months)
  return d.toISOString().slice(0,10)
}
const feTotal=fe=>Number(fe.amount||0)*(1+Number(fe.ivaRate||0)/100)
const feDueDate=fe=>fe.type==='once'?fe.date:fe.nextDue
const feIsOverdue=fe=>{const d=feDueDate(fe);return d&&d<todayStr}
const feIsDueToday=fe=>{const d=feDueDate(fe);return d===todayStr}

const PRR_CATS=[
  {id:'c1',label:'Equipamentos e componentes',sub:'Integração de IA nos processos existentes',budget:2500},
  {id:'c2',label:'Software – produtividade',sub:'Subscrição de software as a service',budget:1836},
  {id:'c3',label:'Software – negócio',sub:'Subscrição de software as a service',budget:2423.76},
  {id:'c4',label:'Contratação – produtividade',sub:'Técnicos/gestores de plataforma (24m)',budget:40000},
  {id:'c5',label:'Contratação – negócio',sub:'Técnicos/gestores de plataforma (24m)',budget:40000},
  {id:'c6',label:'Consultoria – produtividade',sub:'Aquisição de serviços/formação',budget:43500},
  {id:'c7',label:'Consultoria – negócio',sub:'Aquisição de serviços/formação',budget:40000},
  {id:'c8',label:'Contabilista – produtividade',sub:'Validação de pedidos de pagamento',budget:1250},
  {id:'c9',label:'Contabilista – negócio',sub:'Validação de pedidos de pagamento',budget:1250},
]
const PRR_BUDGET=172759.76,PRR_REIMB=129569.82,PRR_ADVANCE=38870.95
const IVA_RATES=[{v:0,l:'0% — Isento'},{v:6,l:'6%'},{v:13,l:'13%'},{v:23,l:'23%'}]

const blankInv=()=>({id:0,type:'venda',client:'',desc:'',amount:'',issued:todayStr,due:'',status:'unpaid',ivaRate:23,ivaLines:[],alertTag:false,createdAt:new Date().toISOString(),attachments:[]})
const blankExp=()=>({id:0,name:'',cat:'',amount:0,type:'one-off',day:1,date:todayStr,paid:false,ivaRate:23,ivaLines:[],ivaDeductible:true,notes:'',alertTag:false,createdAt:new Date().toISOString(),attachments:[]})
const blankGExp=()=>({id:0,grantId:'PRR',categoryId:'',name:'',supplier:'',amount:'',date:todayStr,ivaRate:23,invoiceFile:null,submittedDate:'',expectedSubmission:'',reimbursementDate:'',reimbursementAmount:'',paid:false,alertTag:false,notes:'',createdAt:new Date().toISOString()})
const blankFutureExp=()=>({id:0,name:'',cat:'',amount:'',ivaRate:23,ivaDeductible:true,type:'recurring',date:todayStr,frequency:'monthly',dayOfMonth:1,nextDue:'',isPRR:false,prrCategoryId:'',notes:'',createdAt:new Date().toISOString()})

const gexpSt=e=>{
  if(e.reimbursementDate)return{label:'Reembolsado',icon:'✅',color:'#16a34a',bg:'#dcfce7'}
  if(e.submittedDate)return{label:'Submetido',icon:'📤',color:'#2563eb',bg:'#dbeafe'}
  if(e.invoiceFile)return{label:'Pronto',icon:'📋',color:'#d97706',bg:'#fef3c7'}
  return{label:'Sem fatura',icon:'⚠️',color:'#dc2626',bg:'#fee2e2'}
}

// ── Shared UI ──────────────────────────────────────────────────────────────
function Field({label,value,onChange,type='text',placeholder='',required=false}){
  return(
    <div>
      <label style={{fontSize:12,fontWeight:600,color:'#64748b',display:'block',marginBottom:4}}>{label}{required&&<span style={{color:'#dc2626',marginLeft:3}}>*</span>}</label>
      <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
        style={{width:'100%',padding:'8px 12px',border:'1px solid #e2e8f0',borderRadius:8,fontSize:13,boxSizing:'border-box',outline:'none',fontFamily:'inherit'}}/>
    </div>
  )
}
function FileUploadZone({attachments,onAdd,onRemove}){
  const ref=useRef();const[drag,setDrag]=useState(false)
  const readFile=f=>new Promise((res,rej)=>{if(f.size>5242880){alert('Max 5MB');return rej()}const r=new FileReader();r.onload=e=>res({name:f.name,type:f.type,size:f.size,data:e.target.result.split(',')[1],mime:f.type});r.onerror=rej;r.readAsDataURL(f)})
  const add=async files=>{for(const f of Array.from(files)){try{onAdd(await readFile(f))}catch{}}}
  return(
    <div>
      <label style={{fontSize:12,fontWeight:600,color:'#64748b',display:'block',marginBottom:6}}>📎 Anexos</label>
      <div onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);add(e.dataTransfer.files)}} onClick={()=>ref.current.click()}
        style={{border:`2px dashed ${drag?'#16a34a':'#cbd5e1'}`,borderRadius:8,padding:'10px',textAlign:'center',cursor:'pointer',background:drag?'#f0fdf4':'#f8fafc',marginBottom:attachments.length?8:0}}>
        <div style={{fontSize:12,color:'#64748b'}}>📂 Arrastar ou <span style={{color:'#16a34a',fontWeight:600}}>escolher</span> · max 5MB</div>
        <input ref={ref} type="file" multiple style={{display:'none'}} onChange={e=>add(e.target.files)}/>
      </div>
      {attachments.map((a,i)=>(
        <div key={i} onClick={()=>openFile(a)} style={{display:'flex',alignItems:'center',gap:8,padding:'7px 10px',background:'#f8fafc',borderRadius:8,marginTop:4,border:'1px solid #e2e8f0',cursor:'pointer'}}>
          <span>{FICON(a.type)}</span>
          <div style={{flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{a.name}</div><div style={{fontSize:10,color:'#94a3b8'}}>{fmtBytes(a.size)}</div></div>
          <button onClick={e=>{e.stopPropagation();onRemove(i)}} style={{border:'none',background:'#fee2e2',color:'#dc2626',borderRadius:6,padding:'3px 7px',fontSize:11,cursor:'pointer'}}>✕</button>
        </div>
      ))}
    </div>
  )
}
function AttachmentChips({attachments}){
  if(!attachments?.length)return null
  return(<div style={{display:'flex',gap:4,flexWrap:'wrap',marginTop:5}}>{attachments.map((a,i)=>(
    <button key={i} onClick={()=>dlAll([a])} style={{display:'inline-flex',alignItems:'center',gap:4,background:'#f1f5f9',border:'1px solid #e2e8f0',borderRadius:6,padding:'3px 8px',fontSize:11,cursor:'pointer',color:'#475569',fontWeight:500}}>
      {FICON(a.type)}<span style={{maxWidth:100,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{a.name}</span>↓
    </button>
  ))}</div>)
}
function AiDocStrip({onExtracted,addAttachment,prompt=null}){
  const[aiFile,setAiFile]=useState(null);const[loading,setLoading]=useState(false);const[done,setDone]=useState(false)
  const ref=useRef()
  const readFile=f=>new Promise((res,rej)=>{if(f.size>5242880){alert('Max 5MB');return rej()}const r=new FileReader();r.onload=e=>res({name:f.name,type:f.type,size:f.size,data:e.target.result.split(',')[1],mime:f.type});r.onerror=rej;r.readAsDataURL(f)})
  const pick=async e=>{const f=e.target.files[0];if(!f)return;try{setAiFile(await readFile(f));setDone(false)}catch{}}
  const dfltPrompt='Extract from this document. Respond ONLY with valid JSON (no markdown): {"client":"","supplier":"","description":"","name":"","amount":0,"issued":"YYYY-MM-DD","due":"YYYY-MM-DD","date":"YYYY-MM-DD","ivaLines":[{"base":0,"rate":23}]}. List each VAT rate in ivaLines separately.'
  const extract=async()=>{
    if(!aiFile)return;setLoading(true)
    try{
      const ci=aiFile.type==='application/pdf'?{type:'document',source:{type:'base64',media_type:aiFile.mime,data:aiFile.data}}:{type:'image',source:{type:'base64',media_type:aiFile.mime,data:aiFile.data}}
      const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:600,messages:[{role:"user",content:[ci,{type:"text",text:prompt||dfltPrompt}]}]})})
      const data=await r.json();const txt=data.content?.find(b=>b.type==='text')?.text||'{}'
      const parsed=JSON.parse(txt.replace(/```json|```/g,'').trim())
      onExtracted(parsed);addAttachment(aiFile);setDone(true)
    }catch{alert('Não foi possível extrair — preencha manualmente')}
    setLoading(false)
  }
  return(
    <div style={{background:done?'#f0fdf4':'#f8fafc',border:`1px solid ${done?'#bbf7d0':'#e2e8f0'}`,borderRadius:10,padding:'10px 12px'}}>
      <div style={{fontSize:11,fontWeight:700,color:'#64748b',marginBottom:6}}>🤖 Preencher a partir de documento</div>
      {!aiFile?(
        <button onClick={()=>ref.current.click()} style={{width:'100%',border:'2px dashed #cbd5e1',borderRadius:8,padding:'8px',background:'white',cursor:'pointer',fontSize:12,color:'#64748b',fontWeight:600}}>📎 Carregar documento</button>
      ):(
        <div style={{display:'flex',alignItems:'center',gap:8}}>
          <span style={{flex:1,fontSize:11,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',color:'#475569'}}>{FICON(aiFile.type)} {aiFile.name}</span>
          {done?<span style={{fontSize:11,color:'#16a34a',fontWeight:700}}>✓ Preenchido</span>:
            <button onClick={extract} disabled={loading} style={{border:'none',background:'#16a34a',color:'white',borderRadius:7,padding:'5px 11px',fontSize:12,fontWeight:700,cursor:loading?'not-allowed':'pointer',whiteSpace:'nowrap'}}>{loading?'A analisar…':'🔍 Extrair'}</button>}
          <button onClick={()=>{setAiFile(null);setDone(false)}} style={{border:'none',background:'#fee2e2',color:'#dc2626',borderRadius:6,padding:'4px 7px',fontSize:11,cursor:'pointer'}}>✕</button>
        </div>
      )}
      <input ref={ref} type="file" accept="application/pdf,image/*" style={{display:'none'}} onChange={pick}/>
    </div>
  )
}
function IvaLinesEditor({lines,setLines,ivaDeductible,setIvaDeductible,showDeductible=false}){
  const updLine=(i,k,v)=>setLines(p=>p.map((l,j)=>j===i?{...l,[k]:v}:l))
  const addLine=()=>setLines(p=>[...p,{base:'',rate:23}])
  const removeLine=i=>setLines(p=>p.filter((_,j)=>j!==i))
  const totalBase=lines.reduce((s,l)=>s+Number(l.base||0),0)
  const totalIva=lines.reduce((s,l)=>s+Number(l.base||0)*(Number(l.rate||0)/100),0)
  return(
    <div style={{background:'#f8fafc',borderRadius:10,padding:12,border:'1px solid #e2e8f0',display:'flex',flexDirection:'column',gap:8}}>
      <div style={{fontSize:12,fontWeight:700,color:'#64748b'}}>💶 Valores e Taxas de IVA</div>
      {lines.map((l,i)=>(
        <div key={i} style={{display:'grid',gridTemplateColumns:'1fr 1fr auto',gap:6,alignItems:'flex-end'}}>
          <div>
            <label style={{fontSize:10,fontWeight:600,color:'#94a3b8',display:'block',marginBottom:3}}>Base s/ IVA (€)</label>
            <input type="number" value={l.base} onChange={e=>updLine(i,'base',e.target.value)} placeholder="0.00"
              style={{width:'100%',padding:'7px 10px',border:'1px solid #e2e8f0',borderRadius:7,fontSize:13,boxSizing:'border-box',fontFamily:'inherit',outline:'none'}}/>
          </div>
          <div>
            <label style={{fontSize:10,fontWeight:600,color:'#94a3b8',display:'block',marginBottom:3}}>Taxa IVA</label>
            <select value={l.rate} onChange={e=>updLine(i,'rate',Number(e.target.value))}
              style={{width:'100%',padding:'7px 10px',border:'1px solid #e2e8f0',borderRadius:7,fontSize:13,fontFamily:'inherit',background:'white'}}>
              {IVA_RATES.map(r=><option key={r.v} value={r.v}>{r.l}</option>)}
            </select>
          </div>
          <div style={{display:'flex',flexDirection:'column',alignItems:'center',gap:3,paddingBottom:1}}>
            {Number(l.base)>0&&Number(l.rate)>0&&<span style={{fontSize:10,color:'#7c3aed',fontWeight:700,whiteSpace:'nowrap'}}>{fmt(Number(l.base)*Number(l.rate)/100)}</span>}
            {lines.length>1&&<button onClick={()=>removeLine(i)} style={{border:'none',background:'#fee2e2',color:'#dc2626',borderRadius:5,padding:'4px 7px',fontSize:11,cursor:'pointer'}}>✕</button>}
          </div>
        </div>
      ))}
      <button onClick={addLine} style={{border:'none',background:'#f1f5f9',color:'#475569',borderRadius:7,padding:'5px 10px',fontSize:11,fontWeight:600,cursor:'pointer',alignSelf:'flex-start'}}>+ Adicionar taxa</button>
      {totalBase>0&&(
        <div style={{display:'flex',gap:14,padding:'7px 10px',background:'white',borderRadius:7,fontSize:12,flexWrap:'wrap'}}>
          <span style={{color:'#64748b'}}>Base: <strong>{fmt(totalBase)}</strong></span>
          <span style={{color:'#7c3aed'}}>IVA: <strong>{fmt(totalIva)}</strong></span>
          <span style={{color:'#1e293b',fontWeight:700}}>Total: <strong>{fmt(totalBase+totalIva)}</strong></span>
        </div>
      )}
      {showDeductible&&totalIva>0&&(
        <div style={{background:ivaDeductible===undefined?'#fef3c7':ivaDeductible?'#f0fdf4':'#fef2f2',border:`1px solid ${ivaDeductible===undefined?'#fde68a':ivaDeductible?'#bbf7d0':'#fecaca'}`,borderRadius:8,padding:'10px'}}>
          <div style={{fontSize:12,fontWeight:700,color:'#475569',marginBottom:7}}>💡 O IVA desta despesa é dedutível?</div>
          <div style={{display:'flex',gap:8}}>
            <button onClick={()=>setIvaDeductible(true)} style={{flex:1,padding:'7px',border:'none',borderRadius:8,background:ivaDeductible===true?'#16a34a':'#f1f5f9',color:ivaDeductible===true?'white':'#64748b',fontWeight:700,cursor:'pointer',fontSize:13}}>✓ Sim</button>
            <button onClick={()=>setIvaDeductible(false)} style={{flex:1,padding:'7px',border:'none',borderRadius:8,background:ivaDeductible===false?'#dc2626':'#f1f5f9',color:ivaDeductible===false?'white':'#64748b',fontWeight:700,cursor:'pointer',fontSize:13}}>✕ Não</button>
          </div>
        </div>
      )}
    </div>
  )
}
function FileDocRow({item,depth=2}){
  return(
    <div style={{display:'flex',alignItems:'center',gap:8,padding:'7px 16px',paddingLeft:16+depth*14,borderTop:'1px solid #f1f5f9',background:'white'}}>
      <span style={{fontSize:14}}>{FICON(item.file.type)}</span>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:12,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{item.file.name}</div>
        <div style={{fontSize:10,color:'#94a3b8'}}>{item.label}{item.file.size?` · ${fmtBytes(item.file.size)}`:''}</div>
      </div>
      <button onClick={()=>dlAll([item.file])} style={{border:'none',background:'#f0fdf4',color:'#16a34a',borderRadius:6,padding:'4px 8px',fontSize:11,cursor:'pointer',fontWeight:600,flexShrink:0}}>↓</button>
    </div>
  )
}
function DocsFolder({items,withCategory=false}){
  const[openMap,setOpenMap]=useState({})
  const tog=k=>setOpenMap(p=>({...p,[k]:!p[k]}))
  const tree=useMemo(()=>{
    const t={}
    items.forEach(item=>{
      if(!item.dateStr)return
      const d=new Date(item.dateStr+'T12:00:00');const y=d.getFullYear();const m=d.getMonth()
      const moKey=`${y}-${m}`
      if(!t[y])t[y]={}
      if(!t[y][moKey])t[y][moKey]={label:MONTHS[m],items:withCategory?{}:[]}
      if(withCategory){const cat=item.category||'Sem categoria';if(!t[y][moKey].items[cat])t[y][moKey].items[cat]=[];t[y][moKey].items[cat].push(item)}
      else t[y][moKey].items.push(item)
    })
    return t
  },[items])
  const cntY=months=>Object.values(months).reduce((s,mo)=>s+(withCategory?Object.values(mo.items).reduce((ss,f)=>ss+f.length,0):mo.items.length),0)
  const cntM=mo=>withCategory?Object.values(mo.items).reduce((s,f)=>s+f.length,0):mo.items.length
  if(!items.length){
    return(<div style={{padding:'20px',textAlign:'center',color:'#94a3b8',background:'white',borderRadius:12,fontSize:13}}>Sem documentos</div>)
  }
  return(
    <div style={{background:'white',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,0.07)',overflow:'hidden'}}>
      {Object.entries(tree).sort((a,b)=>Number(b[0])-Number(a[0])).map(([year,months])=>(
        <div key={year}>
          <div onClick={()=>tog(year)} style={{display:'flex',alignItems:'center',gap:8,padding:'10px 16px',cursor:'pointer',background:'#f8fafc',borderBottom:'1px solid #e2e8f0',userSelect:'none'}}>
            <span style={{fontSize:16}}>{openMap[year]?'📂':'📁'}</span>
            <span style={{fontWeight:700,fontSize:13,flex:1}}>{year}</span>
            <span style={{fontSize:11,color:'#94a3b8'}}>{cntY(months)} docs</span>
            <span style={{fontSize:11,color:'#94a3b8',marginLeft:4}}>{openMap[year]?'▾':'▸'}</span>
          </div>
          {openMap[year]&&Object.entries(months).sort((a,b)=>b[0].localeCompare(a[0])).map(([moKey,moData])=>(
            <div key={moKey}>
              <div onClick={()=>tog(moKey)} style={{display:'flex',alignItems:'center',gap:8,padding:'8px 16px',paddingLeft:28,cursor:'pointer',borderBottom:'1px solid #f1f5f9',userSelect:'none'}}>
                <span style={{fontSize:14}}>{openMap[moKey]?'📂':'📁'}</span>
                <span style={{fontWeight:600,fontSize:12,flex:1}}>{moData.label}</span>
                <span style={{fontSize:11,color:'#94a3b8'}}>{cntM(moData)} docs</span>
                <span style={{fontSize:11,color:'#94a3b8',marginLeft:4}}>{openMap[moKey]?'▾':'▸'}</span>
              </div>
              {openMap[moKey]&&(withCategory
                ?Object.entries(moData.items).map(([cat,files])=>(
                  <div key={cat}>
                    <div onClick={()=>tog(`${moKey}-${cat}`)} style={{display:'flex',alignItems:'center',gap:8,padding:'7px 16px',paddingLeft:40,cursor:'pointer',borderBottom:'1px solid #f1f5f9',userSelect:'none',background:'#fafafa'}}>
                      <span style={{fontSize:13}}>{openMap[`${moKey}-${cat}`]?'📂':'📁'}</span>
                      <span style={{fontWeight:500,fontSize:12,flex:1,color:'#475569'}}>{cat}</span>
                      <span style={{fontSize:11,color:'#94a3b8'}}>{files.length} docs</span>
                      <span style={{fontSize:11,color:'#94a3b8',marginLeft:4}}>{openMap[`${moKey}-${cat}`]?'▾':'▸'}</span>
                    </div>
                    {openMap[`${moKey}-${cat}`]&&files.map((item,i)=>(<FileDocRow key={i} item={item} depth={3}/>))}
                  </div>
                ))
                :moData.items.map((item,i)=>(<FileDocRow key={i} item={item} depth={2}/>))
              )}
            </div>
          ))}
        </div>
      ))}
    </div>
  )
}

// ── Modals ─────────────────────────────────────────────────────────────────
function InvoiceModal({inv,onSave,onClose}){
  const initLines=()=>inv.ivaLines?.length?inv.ivaLines.map(l=>({base:String(l.base),rate:l.rate})):inv.amount?[{base:String(inv.amount),rate:inv.ivaRate??23}]:[{base:'',rate:23}]
  const[f,setF]=useState({...inv,attachments:[...(inv.attachments||[])]})
  const[lines,setLines]=useState(initLines)
  const set=(k,v)=>setF(p=>({...p,[k]:v}))
  const totalBase=lines.reduce((s,l)=>s+Number(l.base||0),0)
  const valid=f.client&&totalBase>0&&f.issued
  const onExtracted=p=>{
    if(p.client)set('client',p.client)
    if(p.description||p.name)set('desc',p.description||p.name)
    if(p.issued)set('issued',p.issued)
    if(p.due)set('due',p.due)
    if(p.ivaLines?.length)setLines(p.ivaLines.map(l=>({base:String(l.base||0),rate:Number(l.rate||0)})))
    else if(p.amount)setLines([{base:String(p.amount),rate:p.ivaRate??23}])
  }
  const handleSave=()=>{
    if(!valid)return
    const cleanLines=lines.filter(l=>Number(l.base||0)>0).map(l=>({base:Number(l.base),rate:Number(l.rate||0)}))
    onSave({...f,amount:totalBase,ivaLines:cleanLines,ivaRate:cleanLines[0]?.rate??0})
  }
  return(
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.45)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:200,padding:16}}>
      <div style={{background:'white',borderRadius:16,padding:24,width:'100%',maxWidth:480,maxHeight:'92vh',overflowY:'auto',boxShadow:'0 20px 60px rgba(0,0,0,0.2)'}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:18}}>
          <div style={{fontWeight:700,fontSize:16}}>💰 {inv.id?'Editar Venda':'Nova Venda'}</div>
          <button onClick={onClose} style={{border:'none',background:'#f1f5f9',borderRadius:8,width:32,height:32,cursor:'pointer',fontSize:16}}>✕</button>
        </div>
        <div style={{display:'flex',flexDirection:'column',gap:13}}>
          <AiDocStrip onExtracted={onExtracted} addAttachment={a=>set('attachments',[...f.attachments,a])}/><div style={{display:'flex',gap:8}}><button onClick={()=>set('type','venda')} style={{flex:1,padding:'8px',border:`2px solid ${(f.type||'venda')==='venda'?'#16a34a':'#e2e8f0'}`,borderRadius:8,background:(f.type||'venda')==='venda'?'#f0fdf4':'white',cursor:'pointer',fontSize:13,fontWeight:600,color:(f.type||'venda')==='venda'?'#16a34a':'#64748b'}}>💰 Venda</button><button onClick={()=>set('type','premio')} style={{flex:1,padding:'8px',border:`2px solid ${f.type==='premio'?'#16a34a':'#e2e8f0'}`,borderRadius:8,background:f.type==='premio'?'#f0fdf4':'white',cursor:'pointer',fontSize:13,fontWeight:600,color:f.type==='premio'?'#16a34a':'#64748b'}}>🏆 Prémio</button></div>
          <Field label="Cliente" value={f.client} onChange={v=>set('client',v)} required/>
          <Field label="Descrição" value={f.desc} onChange={v=>set('desc',v)}/>
          <IvaLinesEditor lines={lines} setLines={setLines} showDeductible={false}/>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
            <Field label="Data de emissão" value={f.issued} onChange={v=>set('issued',v)} type="date" required/>
            <Field label="Data de vencimento" value={f.due} onChange={v=>set('due',v)} type="date"/>
          </div>
          <div>
            <label style={{fontSize:12,fontWeight:600,color:'#64748b',display:'block',marginBottom:4}}>Estado</label>
            <select value={f.status} onChange={e=>set('status',e.target.value)} style={{width:'100%',padding:'8px 12px',border:'1px solid #e2e8f0',borderRadius:8,fontSize:13,fontFamily:'inherit',background:'white'}}>
              <option value="unpaid">⏳ Por receber</option><option value="paid">✓ Recebida</option>
</select>
</div>
<label style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',fontSize:13,background:f.alertTag?'#fef2f2':'#f8fafc',padding:'8px 10px',borderRadius:8,border:`1px solid ${f.alertTag?'#fecaca':'#e2e8f0'}`}}><input type="checkbox" checked={!!f.alertTag} onChange={e=>set('alertTag',e.target.checked)} style={{width:15,height:15}}/>🚩 Marcar com tag de alerta</label>
<FileUploadZone attachments={f.attachments} onAdd={a=>set('attachments',[...f.attachments,a])} onRemove={i=>set('attachments',f.attachments.filter((_,j)=>j!==i))}/>
        </div>
        <div style={{display:'flex',gap:8,marginTop:20}}>
          <button onClick={onClose} style={{flex:1,padding:'10px',border:'1px solid #e2e8f0',borderRadius:8,background:'white',cursor:'pointer',fontSize:14,fontWeight:600,color:'#64748b'}}>Cancelar</button>
          <button onClick={handleSave} style={{flex:2,padding:'10px',border:'none',borderRadius:8,background:valid?'#16a34a':'#94a3b8',color:'white',cursor:valid?'pointer':'not-allowed',fontSize:14,fontWeight:700}}>Guardar</button>
        </div>
      </div>
    </div>
  )
}
function ExpenseModal({exp,onSave,onClose,onConvert}){
  const initLines=()=>exp.ivaLines?.length?exp.ivaLines.map(l=>({base:String(l.base),rate:l.rate})):exp.amount?[{base:String(exp.amount),rate:exp.ivaRate??23}]:[{base:'',rate:23}]
  const[f,setF]=useState({ivaDeductible:true,...exp,attachments:[...(exp.attachments||[])]})
  const[lines,setLines]=useState(initLines)
  const set=(k,v)=>setF(p=>({...p,[k]:v}))
  const totalBase=lines.reduce((s,l)=>s+Number(l.base||0),0)
  const valid=f.name&&totalBase>0
  const expPrompt='Extract from this expense receipt/invoice. Respond ONLY with valid JSON (no markdown): {"name":"description","supplier":"vendor name","date":"YYYY-MM-DD","ivaLines":[{"base":0,"rate":23}]}.'
  const onExtracted=p=>{
    if(p.name||p.description)set('name',p.name||p.description)
    if(p.supplier)set('cat',p.supplier)
    if(p.date)set('date',p.date)
    if(p.ivaLines?.length)setLines(p.ivaLines.map(l=>({base:String(l.base||0),rate:Number(l.rate||0)})))
    else if(p.amount)setLines([{base:String(p.amount),rate:p.ivaRate??23}])
  }
  const handleSave=()=>{
    if(!valid)return
    const cleanLines=lines.filter(l=>Number(l.base||0)>0).map(l=>({base:Number(l.base),rate:Number(l.rate||0)}))
    onSave({...f,amount:totalBase,ivaLines:cleanLines,ivaRate:cleanLines[0]?.rate??0})
  }
  return(
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.45)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:200,padding:16}}>
      <div style={{background:'white',borderRadius:16,padding:24,width:'100%',maxWidth:420,maxHeight:'92vh',overflowY:'auto',boxShadow:'0 20px 60px rgba(0,0,0,0.2)'}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:18}}>
          <div style={{fontWeight:700,fontSize:16}}>💸 {exp.id?'Editar Despesa':'Nova Despesa'}</div>
          <button onClick={onClose} style={{border:'none',background:'#f1f5f9',borderRadius:8,width:32,height:32,cursor:'pointer',fontSize:16}}>✕</button>
        </div>
        <div style={{display:'flex',flexDirection:'column',gap:13}}>
          <AiDocStrip onExtracted={onExtracted} addAttachment={a=>set('attachments',[...f.attachments,a])} prompt={expPrompt}/>
          <Field label="Nome" value={f.name} onChange={v=>set('name',v)} required/>
          <Field label="Categoria / Fornecedor" value={f.cat} onChange={v=>set('cat',v)}/>
          <Field label="Data" value={f.date} onChange={v=>set('date',v)} type="date"/>
          <IvaLinesEditor lines={lines} setLines={setLines} ivaDeductible={f.ivaDeductible} setIvaDeductible={v=>set('ivaDeductible',v)} showDeductible={true}/>
          <label style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',fontSize:13}}><input type="checkbox" checked={!!f.paid} onChange={e=>set('paid',e.target.checked)} style={{width:15,height:15}}/>Já pago</label>
<label style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',fontSize:13,background:f.alertTag?'#fef2f2':'#f8fafc',padding:'8px 10px',borderRadius:8,border:`1px solid ${f.alertTag?'#fecaca':'#e2e8f0'}`}}><input type="checkbox" checked={!!f.alertTag} onChange={e=>set('alertTag',e.target.checked)} style={{width:15,height:15}}/>🚩 Marcar com tag de alerta</label>
<Field label="Notas" value={f.notes||''} onChange={v=>set('notes',v)}/><button onClick={()=>onConvert(f,lines)} style={{border:'1px dashed #93c5fd',background:'#eff6ff',color:'#1d4ed8',borderRadius:8,padding:'9px 10px',fontSize:12,fontWeight:700,cursor:'pointer'}}>🏛️ Transformar em Despesa PRR</button>
          <FileUploadZone attachments={f.attachments} onAdd={a=>set('attachments',[...f.attachments,a])} onRemove={i=>set('attachments',f.attachments.filter((_,j)=>j!==i))}/>
        </div>
        <div style={{display:'flex',gap:8,marginTop:20}}>
          <button onClick={onClose} style={{flex:1,padding:'10px',border:'1px solid #e2e8f0',borderRadius:8,background:'white',cursor:'pointer',fontSize:14,fontWeight:600,color:'#64748b'}}>Cancelar</button>
          <button onClick={handleSave} style={{flex:2,padding:'10px',border:'none',borderRadius:8,background:valid?'#16a34a':'#94a3b8',color:'white',cursor:valid?'pointer':'not-allowed',fontSize:14,fontWeight:700}}>Guardar</button>
        </div>
      </div>
    </div>
  )
}

// ── Future Expense Modal ───────────────────────────────────────────────────
function FutureExpModal({fexp,onSave,onClose}){
  const[f,setF]=useState({...fexp})
  const set=(k,v)=>setF(p=>({...p,[k]:v}))
  const valid=f.name&&f.amount
  const previewDue=f.type==='once'?f.date:computeNextDue(f.frequency,f.dayOfMonth)
  const handleSave=()=>{
    if(!valid)return
    const nextDue=f.type==='once'?f.date:computeNextDue(f.frequency,f.dayOfMonth)
    onSave({...f,amount:Number(f.amount),nextDue})
  }
  return(
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.45)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:200,padding:16}}>
      <div style={{background:'white',borderRadius:16,padding:24,width:'100%',maxWidth:440,maxHeight:'92vh',overflowY:'auto',boxShadow:'0 20px 60px rgba(0,0,0,0.2)'}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:18}}>
          <div style={{fontWeight:700,fontSize:16}}>🗓️ {fexp.id?'Editar Despesa Futura':'Nova Despesa Futura'}</div>
          <button onClick={onClose} style={{border:'none',background:'#f1f5f9',borderRadius:8,width:32,height:32,cursor:'pointer',fontSize:16}}>✕</button>
        </div>
        <div style={{display:'flex',flexDirection:'column',gap:13}}>
          <Field label="Nome" value={f.name} onChange={v=>set('name',v)} required/>
          <Field label="Categoria / Fornecedor" value={f.cat} onChange={v=>set('cat',v)}/>
          <Field label="Valor estimado (€ c/ IVA)" value={f.amount} onChange={v=>set('amount',v)} type="number" required/>
          <div>
            <label style={{fontSize:12,fontWeight:600,color:'#64748b',display:'block',marginBottom:4}}>Taxa IVA</label>
            <select value={f.ivaRate} onChange={e=>set('ivaRate',Number(e.target.value))} style={{width:'100%',padding:'8px 12px',border:'1px solid #e2e8f0',borderRadius:8,fontSize:13,fontFamily:'inherit',background:'white'}}>
              {IVA_RATES.map(r=><option key={r.v} value={r.v}>{r.l}</option>)}
            </select>
          </div>
          <div>
            <label style={{fontSize:12,fontWeight:600,color:'#64748b',display:'block',marginBottom:4}}>Tipo</label>
            <div style={{display:'flex',gap:8}}>
              <button onClick={()=>set('type','recurring')} style={{flex:1,padding:'8px',border:`2px solid ${f.type==='recurring'?'#16a34a':'#e2e8f0'}`,borderRadius:8,background:f.type==='recurring'?'#f0fdf4':'white',cursor:'pointer',fontSize:13,fontWeight:600,color:f.type==='recurring'?'#16a34a':'#64748b'}}>🔄 Recorrente</button>
              <button onClick={()=>set('type','once')} style={{flex:1,padding:'8px',border:`2px solid ${f.type==='once'?'#16a34a':'#e2e8f0'}`,borderRadius:8,background:f.type==='once'?'#f0fdf4':'white',cursor:'pointer',fontSize:13,fontWeight:600,color:f.type==='once'?'#16a34a':'#64748b'}}>🔷 Pontual</button>
            </div>
          </div>
          {f.type==='recurring'?(
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
              <div>
                <label style={{fontSize:12,fontWeight:600,color:'#64748b',display:'block',marginBottom:4}}>Frequência</label>
                <select value={f.frequency} onChange={e=>set('frequency',e.target.value)} style={{width:'100%',padding:'8px 12px',border:'1px solid #e2e8f0',borderRadius:8,fontSize:13,fontFamily:'inherit',background:'white'}}>
                  {FREQUENCIES.map(fr=><option key={fr.v} value={fr.v}>{fr.l}</option>)}
                </select>
              </div>
              <Field label="Dia do mês" value={f.dayOfMonth} onChange={v=>set('dayOfMonth',Math.min(28,Math.max(1,Number(v)||1)))} type="number"/>
            </div>
          ):(
            <Field label="Data" value={f.date} onChange={v=>set('date',v)} type="date" required/>
          )}
          {previewDue&&<div style={{background:'#f0fdf4',borderRadius:8,padding:'8px 12px',fontSize:12,color:'#16a34a'}}>📅 Próximo vencimento: <strong>{fmtD(previewDue)}</strong></div>}
          <label style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',fontSize:13}}><input type="checkbox" checked={!!f.ivaDeductible} onChange={e=>set('ivaDeductible',e.target.checked)} style={{width:15,height:15}}/>IVA dedutível</label>
          <label style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',fontSize:13}}><input type="checkbox" checked={!!f.isPRR} onChange={e=>set('isPRR',e.target.checked)} style={{width:15,height:15}}/>Despesa PRR elegível</label>
          {f.isPRR&&(
            <div>
              <label style={{fontSize:12,fontWeight:600,color:'#1d4ed8',display:'block',marginBottom:4}}>Categoria PRR</label>
              <select value={f.prrCategoryId||''} onChange={e=>set('prrCategoryId',e.target.value)} style={{width:'100%',padding:'8px 12px',border:'1px solid #bfdbfe',borderRadius:8,fontSize:12,fontFamily:'inherit',background:'#eff6ff'}}>
                <option value="">Selecionar…</option>
                {PRR_CATS.map(c=><option key={c.id} value={c.id}>{c.label}</option>)}
              </select>
            </div>
          )}
          <Field label="Notas" value={f.notes} onChange={v=>set('notes',v)}/>
        </div>
        <div style={{display:'flex',gap:8,marginTop:20}}>
          <button onClick={onClose} style={{flex:1,padding:'10px',border:'1px solid #e2e8f0',borderRadius:8,background:'white',cursor:'pointer',fontSize:14,fontWeight:600,color:'#64748b'}}>Cancelar</button>
          <button onClick={handleSave} style={{flex:2,padding:'10px',border:'none',borderRadius:8,background:valid?'#16a34a':'#94a3b8',color:'white',cursor:valid?'pointer':'not-allowed',fontSize:14,fontWeight:700}}>Guardar</button>
        </div>
      </div>
    </div>
  )
}

// ── Pay Future Expense Modal ───────────────────────────────────────────────
function PayFutureExpModal({fexp,onPay,onClose}){
  const initLines=[{base:String(fexp.amount||''),rate:fexp.ivaRate||23}]
  const[lines,setLines]=useState(initLines)
  const[file,setFile]=useState(null)
  const[ivaDeductible,setIvaDeductible]=useState(fexp.ivaDeductible!==false)
  const ref=useRef()
  const totalBase=lines.reduce((s,l)=>s+Number(l.base||0),0)
  const valid=totalBase>0
  const readFile=f=>new Promise((res,rej)=>{if(f.size>5242880){alert('Max 5MB');return rej()}const r=new FileReader();r.onload=e=>res({name:f.name,type:f.type,size:f.size,data:e.target.result.split(',')[1],mime:f.type});r.onerror=rej;r.readAsDataURL(f)})
  const pick=async e=>{const f=e.target.files[0];if(!f)return;try{setFile(await readFile(f))}catch{}}
  return(
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:250,padding:16}}>
      <div style={{background:'white',borderRadius:16,padding:24,width:'100%',maxWidth:420,maxHeight:'92vh',overflowY:'auto',boxShadow:'0 20px 60px rgba(0,0,0,0.25)'}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
          <div style={{fontWeight:700,fontSize:16}}>💳 Registar pagamento</div>
          <button onClick={onClose} style={{border:'none',background:'#f1f5f9',borderRadius:8,width:32,height:32,cursor:'pointer',fontSize:16}}>✕</button>
        </div>
        <div style={{background:'#f8fafc',borderRadius:10,padding:'10px 14px',marginBottom:14}}>
          <div style={{fontWeight:600,fontSize:14}}>{fexp.name}</div>
          <div style={{fontSize:12,color:'#64748b'}}>{fexp.cat}{fexp.type==='recurring'?` · ${FREQUENCIES.find(f=>f.v===fexp.frequency)?.l}`:''}</div>
        </div>
        <div style={{display:'flex',flexDirection:'column',gap:13}}>
          <IvaLinesEditor lines={lines} setLines={setLines} ivaDeductible={ivaDeductible} setIvaDeductible={setIvaDeductible} showDeductible={true}/>
          {/* Prova de pagamento */}
          <div>
            <label style={{fontSize:12,fontWeight:600,color:'#64748b',display:'block',marginBottom:6}}>📎 Prova de pagamento</label>
            {file?(
              <div style={{display:'flex',alignItems:'center',gap:8,padding:'8px 10px',background:'#f0fdf4',borderRadius:8,border:'1px solid #bbf7d0'}}>
                <span>{FICON(file.type)}</span>
                <span style={{flex:1,fontSize:12,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{file.name}</span>
                <button onClick={()=>setFile(null)} style={{border:'none',background:'#fee2e2',color:'#dc2626',borderRadius:5,padding:'3px 6px',fontSize:11,cursor:'pointer'}}>✕</button>
              </div>
            ):(
              <button onClick={()=>ref.current.click()} style={{width:'100%',border:'2px dashed #fde68a',borderRadius:8,padding:'10px',background:'#fefce8',cursor:'pointer',fontSize:12,color:'#92400e',fontWeight:600}}>⚠️ Anexar documento (obrigatório para PRR)</button>
            )}
            <input ref={ref} type="file" accept="application/pdf,image/*" style={{display:'none'}} onChange={pick}/>
          </div>
        </div>
        <div style={{display:'flex',gap:8,marginTop:20}}>
          <button onClick={onClose} style={{flex:1,padding:'10px',border:'1px solid #e2e8f0',borderRadius:8,background:'white',cursor:'pointer',fontSize:14,fontWeight:600,color:'#64748b'}}>Cancelar</button>
          <button onClick={()=>valid&&onPay({lines,file,ivaDeductible})} style={{flex:2,padding:'10px',border:'none',borderRadius:8,background:valid?'#16a34a':'#94a3b8',color:'white',cursor:valid?'pointer':'not-allowed',fontSize:14,fontWeight:700}}>✓ Confirmar pagamento</button>
        </div>
      </div>
    </div>
  )
}

function GrantExpModal({exp,onSave,onClose}){
  const[f,setF]=useState({ivaRate:23,...exp});const set=(k,v)=>setF(p=>({...p,[k]:v}))
  const ref=useRef();const valid=f.categoryId&&f.name&&f.amount&&f.date
  const cat=PRR_CATS.find(c=>c.id===f.categoryId)
  const readInv=file=>new Promise((res,rej)=>{if(file.size>5242880){alert('Max 5MB');return rej()}const r=new FileReader();r.onload=e=>res({name:file.name,type:file.type,size:file.size,data:e.target.result.split(',')[1],mime:file.type});r.onerror=rej;r.readAsDataURL(file)})
  const onPick=async e=>{const file=e.target.files[0];if(!file)return;try{set('invoiceFile',await readInv(file))}catch{}}
  const dlInv=()=>{if(!f.invoiceFile)return;dlAll([f.invoiceFile])}
  return(
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.45)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:200,padding:16}}>
      <div style={{background:'white',borderRadius:16,padding:24,width:'100%',maxWidth:520,maxHeight:'94vh',overflowY:'auto',boxShadow:'0 20px 60px rgba(0,0,0,0.2)'}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:18}}>
          <div style={{display:'flex',alignItems:'center',gap:8}}>
            <div style={{background:'#1d4ed8',color:'white',borderRadius:7,padding:'3px 9px',fontSize:11,fontWeight:800}}>PRR</div>
            <div style={{fontWeight:700,fontSize:16}}>{exp.id?'Editar Despesa PRR':'Nova Despesa Elegível'}</div>
          </div>
          <button onClick={onClose} style={{border:'none',background:'#f1f5f9',borderRadius:8,width:32,height:32,cursor:'pointer',fontSize:16}}>✕</button>
        </div>
        <div style={{display:'flex',flexDirection:'column',gap:13}}>
          <div>
            <label style={{fontSize:12,fontWeight:600,color:'#64748b',display:'block',marginBottom:4}}>Categoria PRR *</label>
            <select value={f.categoryId} onChange={e=>set('categoryId',e.target.value)} style={{width:'100%',padding:'8px 12px',border:'1px solid #e2e8f0',borderRadius:8,fontSize:13,fontFamily:'inherit',background:'white'}}>
              <option value="">Selecionar categoria…</option>
              {PRR_CATS.map(c=><option key={c.id} value={c.id}>{c.label} — {fmt(c.budget)}</option>)}
            </select>
            {cat&&<div style={{fontSize:11,color:'#64748b',marginTop:3}}>📌 {cat.sub}</div>}
          </div>
          <Field label="Descrição" value={f.name} onChange={v=>set('name',v)} required/>
          <Field label="Fornecedor" value={f.supplier} onChange={v=>set('supplier',v)}/>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
            <Field label="Valor total (€ c/ IVA)" value={f.amount} onChange={v=>set('amount',v)} type="number" required/>
            <Field label="Data" value={f.date} onChange={v=>set('date',v)} type="date" required/>
          </div>
          <div>
            <label style={{fontSize:12,fontWeight:600,color:'#64748b',display:'block',marginBottom:4}}>Taxa IVA</label>
            <select value={f.ivaRate} onChange={e=>set('ivaRate',Number(e.target.value))} style={{width:'100%',padding:'8px 12px',border:'1px solid #e2e8f0',borderRadius:8,fontSize:13,fontFamily:'inherit',background:'white'}}>
              {IVA_RATES.map(r=><option key={r.v} value={r.v}>{r.l}</option>)}
            </select>
          </div>
          {f.amount&&Number(f.ivaRate)>0&&<div style={{fontSize:11,color:'#64748b'}}>Base s/ IVA: <strong>{fmt(Number(f.amount)/(1+Number(f.ivaRate)/100))}</strong></div>}
          {f.amount&&<div style={{fontSize:11,color:'#16a34a'}}>💡 Reembolso esperado (75% da base): <strong>{fmt(Number(f.amount)/(1+(Number(f.ivaRate)||0)/100)*0.75)}</strong></div>}
<label style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',fontSize:13}}><input type="checkbox" checked={!!f.paid} onChange={e=>set('paid',e.target.checked)} style={{width:15,height:15}}/>Já pago</label>
<label style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',fontSize:13,background:f.alertTag?'#fef2f2':'#f8fafc',padding:'8px 10px',borderRadius:8,border:`1px solid ${f.alertTag?'#fecaca':'#e2e8f0'}`}}><input type="checkbox" checked={!!f.alertTag} onChange={e=>set('alertTag',e.target.checked)} style={{width:15,height:15}}/>🚩 Marcar com tag de alerta</label>
<div style={{background:f.invoiceFile?'#f0fdf4':'#fff5f5',border:`2px solid ${f.invoiceFile?'#bbf7d0':'#fecaca'}`,borderRadius:10,padding:12}}>
            <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:8}}><span>{f.invoiceFile?'✅':'⚠️'}</span><span style={{fontSize:12,fontWeight:700,color:f.invoiceFile?'#16a34a':'#dc2626'}}>{f.invoiceFile?'Fatura anexada':'Fatura obrigatória'}</span></div>
            {f.invoiceFile?(
              <div style={{display:'flex',alignItems:'center',gap:8,background:'white',borderRadius:8,padding:'8px 10px',border:'1px solid #bbf7d0'}}>
                <span>{FICON(f.invoiceFile.type)}</span>
                <div style={{flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{f.invoiceFile.name}</div></div>
                <button onClick={dlInv} style={{border:'none',background:'#f0fdf4',color:'#16a34a',borderRadius:6,padding:'3px 8px',fontSize:11,cursor:'pointer'}}>↓</button>
                <button onClick={()=>set('invoiceFile',null)} style={{border:'none',background:'#fee2e2',color:'#dc2626',borderRadius:6,padding:'3px 7px',fontSize:11,cursor:'pointer'}}>✕</button>
              </div>
            ):(
              <button onClick={()=>ref.current.click()} style={{width:'100%',border:'2px dashed #fca5a5',borderRadius:8,padding:'10px',background:'white',cursor:'pointer',fontSize:13,color:'#dc2626',fontWeight:600}}>📎 Carregar fatura</button>
            )}
            <input ref={ref} type="file" accept="application/pdf,image/*" style={{display:'none'}} onChange={onPick}/>
          </div>
          <div style={{background:'#eff6ff',borderRadius:10,padding:12,border:'1px solid #bfdbfe'}}>
            <div style={{fontSize:12,fontWeight:700,color:'#1d4ed8',marginBottom:8}}>📤 Balcão dos Fundos</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
              <Field label="Data de submissão" value={f.submittedDate} onChange={v=>set('submittedDate',v)} type="date"/>
              <Field label="Previsão" value={f.expectedSubmission} onChange={v=>set('expectedSubmission',v)} type="date"/>
            </div>
          </div>
          <div style={{background:'#f0fdf4',borderRadius:10,padding:12,border:'1px solid #bbf7d0'}}>
            <div style={{fontSize:12,fontWeight:700,color:'#16a34a',marginBottom:8}}>💶 Reembolso (75%)</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
              <Field label="Valor recebido (€)" value={f.reimbursementAmount} onChange={v=>set('reimbursementAmount',v)} type="number"/>
              <Field label="Data de receção" value={f.reimbursementDate} onChange={v=>set('reimbursementDate',v)} type="date"/>
            </div>
          </div>
          <Field label="Notas" value={f.notes} onChange={v=>set('notes',v)}/>
        </div>
        <div style={{display:'flex',gap:8,marginTop:20}}>
          <button onClick={onClose} style={{flex:1,padding:'10px',border:'1px solid #e2e8f0',borderRadius:8,background:'white',cursor:'pointer',fontSize:14,fontWeight:600,color:'#64748b'}}>Cancelar</button>
          <button onClick={()=>valid&&onSave(f)} style={{flex:2,padding:'10px',border:'none',borderRadius:8,background:valid?'#1d4ed8':'#94a3b8',color:'white',cursor:valid?'pointer':'not-allowed',fontSize:14,fontWeight:700}}>Guardar</button>
        </div>
      </div>
    </div>
  )
}

function BatchUploadModal({onSave,onClose}){
  const[files,setFiles]=useState([]);const[results,setResults]=useState([]);const[analyzing,setAnalyzing]=useState(false);const[analyzed,setAnalyzed]=useState(false)
  const ref=useRef()
  const readFile=f=>new Promise((res,rej)=>{if(f.size>5242880){alert('Max 5MB');return rej()}const r=new FileReader();r.onload=e=>res({name:f.name,type:f.type,size:f.size,data:e.target.result.split(',')[1],mime:f.type});r.onerror=rej;r.readAsDataURL(f)})
  const addFiles=async fl=>{const arr=[];for(const f of Array.from(fl)){try{arr.push(await readFile(f))}catch{}}setFiles(p=>[...p,...arr])}
  const analyze=async()=>{
    setAnalyzing(true);const res=[]
    for(const file of files){
      try{
        const ci=file.type==='application/pdf'?{type:'document',source:{type:'base64',media_type:file.mime,data:file.data}}:{type:'image',source:{type:'base64',media_type:file.mime,data:file.data}}
        const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:600,messages:[{role:"user",content:[ci,{type:"text",text:'Extract. Respond ONLY with valid JSON (no markdown): {"date":"YYYY-MM-DD","description":"","supplier":"","ivaLines":[{"base":0,"rate":23}],"isPRR":false,"suggestedPRRCat":""}. List each VAT rate separately.'}]}]})})
        const data=await r.json();const txt=data.content?.find(b=>b.type==='text')?.text||'{}'
        const parsed=JSON.parse(txt.replace(/```json|```/g,'').trim())
        const totalBase=parsed.ivaLines?.reduce((s,l)=>s+Number(l.base||0),0)||0
        res.push({file,error:null,selected:true,saveAs:parsed.isPRR?'grant':'expense',grantCategoryId:'',totalBase,...parsed})
      }catch{res.push({file,error:true,selected:true,saveAs:'expense',grantCategoryId:'',date:'',description:'',supplier:'',ivaLines:[{base:'',rate:23}],isPRR:false,suggestedPRRCat:'',totalBase:0})}
    }
    setResults(res);setAnalyzing(false);setAnalyzed(true)
  }
  const upd=(i,k,v)=>setResults(p=>p.map((r,j)=>j===i?{...r,[k]:v}:r))
  const saveAll=()=>onSave(results.filter(r=>r.selected))
  return(
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:300,padding:16}}>
      <div style={{background:'white',borderRadius:16,padding:24,width:'100%',maxWidth:640,maxHeight:'94vh',overflowY:'auto',boxShadow:'0 20px 60px rgba(0,0,0,0.25)'}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:18}}>
          <div style={{fontWeight:700,fontSize:16}}>🤖 AI Batch Import</div>
          <button onClick={onClose} style={{border:'none',background:'#f1f5f9',borderRadius:8,width:32,height:32,cursor:'pointer',fontSize:16}}>✕</button>
        </div>
        {!analyzed?(
          <>
            <div onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();addFiles(e.dataTransfer.files)}} onClick={()=>ref.current.click()}
              style={{border:'2px dashed #cbd5e1',borderRadius:10,padding:24,textAlign:'center',cursor:'pointer',background:'#f8fafc',marginBottom:12}}>
              <div style={{fontSize:28,marginBottom:4}}>📂</div>
              <div style={{fontSize:13,color:'#64748b'}}>Arrastar ou <span style={{color:'#16a34a',fontWeight:600}}>escolher</span></div>
              <div style={{fontSize:11,color:'#94a3b8',marginTop:3}}>PDF, PNG, JPG · max 5MB cada</div>
              <input ref={ref} type="file" multiple accept="application/pdf,image/*" style={{display:'none'}} onChange={e=>addFiles(e.target.files)}/>
            </div>
            {files.length>0&&<div style={{marginBottom:14}}>{files.map((f,i)=>(
              <div key={i} style={{display:'flex',alignItems:'center',gap:8,padding:'7px 10px',background:'#f8fafc',borderRadius:8,marginBottom:4,border:'1px solid #e2e8f0'}}>
                <span>{FICON(f.type)}</span><span style={{flex:1,fontSize:12,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{f.name}</span>
                <span style={{fontSize:11,color:'#94a3b8'}}>{fmtBytes(f.size)}</span>
                <button onClick={()=>setFiles(p=>p.filter((_,j)=>j!==i))} style={{border:'none',background:'#fee2e2',color:'#dc2626',borderRadius:5,padding:'2px 6px',fontSize:11,cursor:'pointer'}}>✕</button>
              </div>
            ))}</div>}
            <div style={{display:'flex',gap:8}}>
              <button onClick={onClose} style={{flex:1,padding:'10px',border:'1px solid #e2e8f0',borderRadius:8,background:'white',cursor:'pointer',fontSize:14,fontWeight:600,color:'#64748b'}}>Cancelar</button>
              <button onClick={analyze} disabled={!files.length||analyzing} style={{flex:2,padding:'10px',border:'none',borderRadius:8,background:files.length&&!analyzing?'#16a34a':'#94a3b8',color:'white',cursor:files.length&&!analyzing?'pointer':'not-allowed',fontSize:14,fontWeight:700}}>
                {analyzing?'🤖 A analisar…':`🔍 Analisar ${files.length} ficheiro${files.length!==1?'s':''}`}
              </button>
            </div>
          </>
        ):(
          <>
            <div style={{fontSize:13,color:'#64748b',marginBottom:14}}>Confirmar antes de guardar:</div>
            {results.map((r,i)=>(
              <div key={i} style={{background:'#f8fafc',borderRadius:10,padding:13,marginBottom:10,border:'2px solid #e2e8f0'}}>
                <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:10}}>
                  <input type="checkbox" checked={!!r.selected} onChange={e=>upd(i,'selected',e.target.checked)} style={{width:15,height:15,cursor:'pointer'}}/>
                  <span style={{fontSize:12,fontWeight:600,flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{FICON(r.file.type)} {r.file.name}</span>
                  <span style={{fontSize:10,fontWeight:600,padding:'2px 6px',borderRadius:5,...(r.error?{background:'#fef3c7',color:'#d97706'}:{background:'#dcfce7',color:'#16a34a'})}}>{r.error?'⚠️':'✓ AI'}</span>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:8}}>
                  <Field label="Data" value={r.date||''} onChange={v=>upd(i,'date',v)} type="date"/>
                  <Field label="Fornecedor" value={r.supplier||''} onChange={v=>upd(i,'supplier',v)}/>
                  <div style={{gridColumn:'1/-1'}}><Field label="Descrição" value={r.description||''} onChange={v=>upd(i,'description',v)}/></div>
                </div>
                {r.ivaLines?.map((l,li)=>(
                  <div key={li} style={{display:'flex',gap:6,marginBottom:4,alignItems:'center'}}>
                    <span style={{fontSize:11,color:'#64748b',minWidth:40}}>Base:</span>
                    <input type="number" value={l.base||''} onChange={e=>{const nl=[...r.ivaLines];nl[li]={...nl[li],base:e.target.value};upd(i,'ivaLines',nl)}} style={{flex:1,padding:'5px 8px',border:'1px solid #e2e8f0',borderRadius:6,fontSize:12,fontFamily:'inherit'}}/>
                    <select value={l.rate} onChange={e=>{const nl=[...r.ivaLines];nl[li]={...nl[li],rate:Number(e.target.value)};upd(i,'ivaLines',nl)}} style={{padding:'5px 8px',border:'1px solid #e2e8f0',borderRadius:6,fontSize:12,fontFamily:'inherit',background:'white'}}>
                      {IVA_RATES.map(rate=><option key={rate.v} value={rate.v}>{rate.l}</option>)}
                    </select>
                    <span style={{fontSize:11,color:'#7c3aed',minWidth:50,textAlign:'right'}}>{l.base&&l.rate?fmt(Number(l.base)*l.rate/100):''}</span>
                  </div>
                ))}
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginTop:8}}>
                  <div>
                    <label style={{fontSize:11,fontWeight:600,color:'#64748b',display:'block',marginBottom:4}}>Guardar como</label>
                    <select value={r.saveAs||'expense'} onChange={e=>upd(i,'saveAs',e.target.value)} style={{width:'100%',padding:'7px 10px',border:'1px solid #e2e8f0',borderRadius:7,fontSize:12,fontFamily:'inherit',background:'white'}}>
                      <option value="expense">💸 Despesa</option>
                      <option value="invoice">💰 Venda (receita)</option>
                      <option value="grant">🏛️ Despesa PRR</option>
                    </select>
                  </div>
                  {r.saveAs==='grant'&&(
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:'#1d4ed8',display:'block',marginBottom:4}}>Categoria PRR *</label>
                      <select value={r.grantCategoryId||''} onChange={e=>upd(i,'grantCategoryId',e.target.value)} style={{width:'100%',padding:'7px 10px',border:`1px solid ${r.grantCategoryId?'#bfdbfe':'#fca5a5'}`,borderRadius:7,fontSize:11,fontFamily:'inherit',background:'#eff6ff'}}>
                        <option value="">Selecionar…</option>
                        {PRR_CATS.map(c=><option key={c.id} value={c.id}>{c.label}</option>)}
                      </select>
                    </div>
                  )}
                </div>
              </div>
            ))}
            <div style={{display:'flex',gap:8,marginTop:14}}>
              <button onClick={()=>{setAnalyzed(false);setResults([])}} style={{flex:1,padding:'10px',border:'1px solid #e2e8f0',borderRadius:8,background:'white',cursor:'pointer',fontSize:14,fontWeight:600,color:'#64748b'}}>← Voltar</button>
              <button onClick={saveAll} style={{flex:2,padding:'10px',border:'none',borderRadius:8,background:'#16a34a',color:'white',cursor:'pointer',fontSize:14,fontWeight:700}}>✓ Guardar {results.filter(r=>r.selected).length} item(s)</button>
            </div>
          </>
        )}
      </div>
    </div>
  )
}

const SBtn=({k,active,onClick,children})=>(
  <button onClick={onClick} style={{border:'none',background:active?'#16a34a':'#f1f5f9',color:active?'white':'#64748b',padding:'5px 11px',borderRadius:20,fontSize:12,fontWeight:600,cursor:'pointer',whiteSpace:'nowrap'}}>{children}</button>
)
const ActBtn=({onClick,children,bg='#f1f5f9',color='#475569'})=>(
  <button onClick={onClick} style={{border:'none',background:bg,color,borderRadius:7,padding:'6px 8px',cursor:'pointer',fontSize:12,flexShrink:0}}>{children}</button>
)
const DlBtn=({count,onClick})=>count>0?(
  <button onClick={onClick} style={{border:'none',background:'#f1f5f9',color:'#475569',borderRadius:8,padding:'6px 10px',fontSize:12,fontWeight:600,cursor:'pointer',display:'flex',alignItems:'center',gap:4}}>
    ⬇ Descarregar todos ({count})
  </button>
):null

export default function App(){
  const[tab,setTab]=useState('dashboard')
  const[invoices,setInvoices]=useState([])
  const[expenses,setExpenses]=useState([])
  const[grantExps,setGrantExps]=useState([])
  const[futureExpenses,setFutureExpenses]=useState([])
  const[yr,setYr]=useState(today.getFullYear())
  const[mo,setMo]=useState(today.getMonth())
  const[ivaYear,setIvaYear]=useState(today.getFullYear())
  const[ivaQ,setIvaQ]=useState(Math.floor(today.getMonth()/3)+1)
  const[ivaSubmitted,setIvaSubmitted]=useState({})
  const[commonExps,setCommonExps]=useState([])
  const[invForm,setInvForm]=useState(null)
  const[expForm,setExpForm]=useState(null)
  const[gexpForm,setGexpForm]=useState(null)
  const[futureForm,setFutureForm]=useState(null)
  const[payModal,setPayModal]=useState(null)
  const[batchModal,setBatchModal]=useState(false)
  const[invFilter,setInvFilter]=useState('all')
const[dExpFilter,setDExpFilter]=useState('all')
const[gCatF,setGCatF]=useState('all')
  const[gStF,setGStF]=useState('all')
  const[cfFilter,setCfFilter]=useState('all')
    const[cfExpanded,setCfExpanded]=useState(false)
  const[allTime,setAllTime]=useState(false)
  const[loaded,setLoaded]=useState(false)

  const selYM=`${yr}-${String(mo+1).padStart(2,'0')}`

  useEffect(()=>{
    (async()=>{
      try{
        const i=await storageGet('ft_inv');if(i)setInvoices(JSON.parse(i.value))
        const e=await storageGet('ft_exp');if(e)setExpenses(JSON.parse(e.value))
        const g=await storageGet('ft_gexp');if(g)setGrantExps(JSON.parse(g.value))
        const fe=await storageGet('ft_future');if(fe)setFutureExpenses(JSON.parse(fe.value))
        const s=await storageGet('ft_ivasub');if(s)setIvaSubmitted(JSON.parse(s.value))
        const c=await storageGet('ft_common');if(c)setCommonExps(JSON.parse(c.value))
      }catch(err){
        console.error('❌ Load error:', err)
      }finally{
        setLoaded(true)
      }
    })()
  },[])
  useEffect(()=>{if(loaded){const save=async()=>await storageSet('ft_inv',JSON.stringify(invoices));save()}},[invoices,loaded])
  useEffect(()=>{if(loaded){const save=async()=>await storageSet('ft_exp',JSON.stringify(expenses));save()}},[expenses,loaded])
  useEffect(()=>{if(loaded){const save=async()=>await storageSet('ft_gexp',JSON.stringify(grantExps));save()}},[grantExps,loaded])
  useEffect(()=>{if(loaded){const save=async()=>await storageSet('ft_future',JSON.stringify(futureExpenses));save()}},[futureExpenses,loaded])
  useEffect(()=>{if(loaded){const save=async()=>await storageSet('ft_ivasub',JSON.stringify(ivaSubmitted));save()}},[ivaSubmitted,loaded])
  useEffect(()=>{if(loaded){const save=async()=>await storageSet('ft_common',JSON.stringify(commonExps));save()}},[commonExps,loaded])

  const prevMo=()=>{if(mo===0){setMo(11);setYr(y=>y-1)}else setMo(m=>m-1)}
  const nextMo=()=>{if(mo===11){setMo(0);setYr(y=>y+1)}else setMo(m=>m+1)}

  // Notifications: future expenses due today or overdue
  const notifications=useMemo(()=>
    futureExpenses.filter(fe=>{const d=feDueDate(fe);return d&&d<=todayStr}).sort((a,b)=>feDueDate(a).localeCompare(feDueDate(b)))
  ,[futureExpenses])

  const calc=useMemo(()=>{
    const mInv=allTime?invoices:invoices.filter(i=>ym(i.issued)===selYM)
    const income=mInv.filter(i=>i.status==='paid').reduce((s,i)=>s+Number(i.amount),0)
    const pending=mInv.filter(i=>i.status==='unpaid').reduce((s,i)=>s+Number(i.amount),0)
    const mExp=allTime?expenses:expenses.filter(e=>ym(e.date)===selYM)
    const mGexp=allTime?grantExps:grantExps.filter(e=>ym(e.date)===selYM)
    const totalExp=mExp.reduce((s,e)=>s+Number(e.amount)+linesIva(getLines(e)),0)+mGexp.reduce((s,e)=>s+Number(e.amount),0)
    const paidExp=mExp.filter(e=>e.paid).reduce((s,e)=>s+Number(e.amount)+linesIva(getLines(e)),0)+mGexp.reduce((s,e)=>s+Number(e.amount),0)
    return{mInv,income,pending,mExp,mGexp,totalExp,paidExp}
  },[invoices,expenses,grantExps,selYM,allTime])

  const catStats=useMemo(()=>PRR_CATS.map(c=>{
    const exps=grantExps.filter(e=>e.categoryId===c.id)
    const spent=exps.reduce((s,e)=>s+grantBase(e),0)
    return{...c,exps,spent,pct:pct(spent,c.budget),reimb:spent*0.75}
  }),[grantExps])

  const prrStats=useMemo(()=>{
    const spent=grantExps.reduce((s,e)=>s+grantBase(e),0)
    const reimbursed=grantExps.filter(e=>e.reimbursementDate).reduce((s,e)=>s+Number(e.reimbursementAmount||grantReim(e)),0)
    const totalReceived=PRR_ADVANCE+reimbursed
    const noInvoice=grantExps.filter(e=>!e.invoiceFile).length
    const readyToSubmit=grantExps.filter(e=>e.invoiceFile&&!e.submittedDate).length
    return{spent,reimbursable:spent*0.75,reimbursed,totalReceived,remaining:PRR_REIMB-totalReceived,noInvoice,readyToSubmit}
  },[grantExps])

  const cashflowEntries=useMemo(()=>{
    const entries=[]
    invoices.forEach(inv=>entries.push({key:`inv-${inv.id}`,date:inv.issued,label:inv.client+(inv.desc?` — ${inv.desc}`:''),amount:Number(inv.amount),flow:'in',settled:inv.status==='paid',tag:inv.type==='premio'?'premio':null,alertTag:!!inv.alertTag}))
expenses.filter(e=>e.date).forEach(exp=>entries.push({key:`exp-${exp.id}`,date:exp.date,label:exp.name+(exp.cat?` · ${exp.cat}`:''),amount:Number(exp.amount)+linesIva(getLines(exp)),flow:'out',settled:!!exp.paid,tag:null,alertTag:!!exp.alertTag}))
    grantExps.forEach(ge=>{const cat=PRR_CATS.find(c=>c.id===ge.categoryId);entries.push({key:`ge-${ge.id}`,date:ge.date,label:ge.name+(ge.supplier?` — ${ge.supplier}`:''),amount:Number(ge.amount),flow:'out',settled:!!ge.paid,tag:'PRR',prrCat:cat?.label,alertTag:!!ge.alertTag})})
    // Future expenses as projected (next occurrence)
    futureExpenses.forEach(fe=>{
      const d=feDueDate(fe)
      if(d)entries.push({key:`fe-${fe.id}`,date:d,label:fe.name+(fe.cat?` · ${fe.cat}`:''),amount:feTotal(fe),flow:'out',settled:false,tag:fe.isPRR?'PRR':'future'})
    })
    entries.sort((a,b)=>a.date.localeCompare(b.date)||(a.flow==='in'?-1:1))
    let bal=0;entries.forEach(e=>{if(e.settled)bal+=e.flow==='in'?e.amount:-e.amount;e.runningBal=bal})
    return[...entries].reverse()
  },[invoices,expenses,grantExps,futureExpenses])

  const currentBalance=cashflowEntries.length?cashflowEntries[0].runningBal:0
  const nextMoDate=new Date(today.getFullYear(),today.getMonth()+1,1)
  const nextYM=`${nextMoDate.getFullYear()}-${String(nextMoDate.getMonth()+1).padStart(2,'0')}`
  const nextMoLabel=`${MONTHS[nextMoDate.getMonth()]} ${nextMoDate.getFullYear()}`
  const nextMonthTotal=useMemo(()=>{
    let total=0
    futureExpenses.forEach(fe=>{
      const amt=feTotal(fe)
      if(fe.type==='once'){if(ym(fe.date)===nextYM)total+=amt}
      else{
        if(fe.frequency==='monthly')total+=amt
        else if(ym(fe.nextDue)===nextYM)total+=amt
      }
    })
    return total
  },[futureExpenses,nextYM])

  const filteredCf=useMemo(()=>{
if(cfFilter==='all')return cashflowEntries
if(cfFilter==='in')return cashflowEntries.filter(e=>e.flow==='in')
if(cfFilter==='out')return cashflowEntries.filter(e=>e.flow==='out')
if(cfFilter==='prr')return cashflowEntries.filter(e=>e.tag==='PRR')
if(cfFilter==='alert')return cashflowEntries.filter(e=>e.alertTag)
if(cfFilter==='unsettled')return cashflowEntries.filter(e=>!e.settled)
return cashflowEntries
},[cashflowEntries,cfFilter])

  const ivaStats=useMemo(()=>{
    const startMo=(ivaQ-1)*3;const endMo=startMo+2
    const inQ=d=>{if(!d)return false;const dt=new Date(d+'T12:00:00');return dt.getFullYear()===ivaYear&&dt.getMonth()>=startMo&&dt.getMonth()<=endMo}
    const qInv=invoices.filter(i=>inQ(i.issued))
    const getExpLines=e=>e.ivaLines?.length?e.ivaLines:[{base:Number(e.amount||0),rate:Number(e.ivaRate||0)}]
    const ivaOut=qInv.filter(i=>i.status==='paid').reduce((s,i)=>s+linesIva(getLines(i)),0)
    const ivaOutPending=qInv.filter(i=>i.status!=='paid').reduce((s,i)=>s+linesIva(getLines(i)),0)
    let ivaIn=0
    expenses.filter(e=>inQ(e.date)&&e.ivaDeductible!==false).forEach(e=>{getExpLines(e).filter(l=>l.rate>0).forEach(l=>{ivaIn+=Number(l.base)*l.rate/100})})
    grantExps.filter(ge=>inQ(ge.date)&&ge.ivaRate>0).forEach(ge=>{ivaIn+=ivaAmt(ge.amount,ge.ivaRate)})
    const deadline=qDeadline(ivaQ,ivaYear);const daysLeft=Math.ceil((deadline-today)/(1000*60*60*24))
    const invLines=qInv.flatMap(i=>{const ls=getLines(i).filter(l=>l.rate>0);return ls.map(l=>({label:i.client+(i.desc?` — ${i.desc}`:'')+( ls.length>1?` (${l.rate}%)`:''),date:i.issued,base:Number(l.base),rate:l.rate,iva:Number(l.base)*l.rate/100,settled:i.status==='paid'}))})
    const expLines=[]
    expenses.filter(e=>inQ(e.date)&&e.ivaDeductible!==false).forEach(e=>{const ls=getExpLines(e).filter(l=>l.rate>0);ls.forEach(l=>expLines.push({label:e.name+(ls.length>1?` (${l.rate}%)`:''),date:e.date,base:Number(l.base),rate:l.rate,iva:Number(l.base)*l.rate/100}))})
    grantExps.filter(ge=>inQ(ge.date)&&ge.ivaRate>0).forEach(ge=>expLines.push({label:ge.name,date:ge.date,base:Number(ge.amount),rate:Number(ge.ivaRate||0),iva:ivaAmt(ge.amount,ge.ivaRate)}))
    return{ivaOut,ivaOutPending,ivaIn,ivaNet:ivaOut-ivaIn,deadline,daysLeft,invLines,expLines}
  },[invoices,expenses,grantExps,ivaYear,ivaQ])

  // Docs
  const[showExportMenu,setShowExportMenu]=useState(false)
  const exportExcel=mode=>{
    const getPer=d=>{if(!d)return'';const dt=new Date(d+'T12:00:00');return mode==='month'?`${MONTHS[dt.getMonth()]} ${dt.getFullYear()}`:String(dt.getFullYear())}
    const vendasRows=invoices.map(inv=>({'Período':getPer(inv.issued),'Data':fmtD(inv.issued),'Cliente':inv.client,'Descrição':inv.desc||'','Base (€)':Number(inv.amount),'IVA (%)':Number(inv.ivaRate||0),'IVA (€)':Math.round(linesIva(getLines(inv))*100)/100,'Total c/ IVA (€)':Math.round((Number(inv.amount)+linesIva(getLines(inv)))*100)/100,'Estado':inv.status==='paid'?'Recebida':'Por receber'})).sort((a,b)=>a['Data'].localeCompare(b['Data']))
    const despesasRows=[
      ...expenses.filter(e=>e.date).map(e=>({'Período':getPer(e.date),'Data':fmtD(e.date),'Nome':e.name,'Categoria':e.cat||'','Tipo':'Normal','Base (€)':Number(e.amount),'IVA (%)':Number(e.ivaRate||0),'IVA (€)':Math.round(linesIva(getLines(e))*100)/100,'Total c/ IVA (€)':Math.round((Number(e.amount)+linesIva(getLines(e)))*100)/100,'IVA Dedutível':e.ivaDeductible?'Sim':'Não','Pago':e.paid?'Sim':'Não'})),
      ...grantExps.map(ge=>({'Período':getPer(ge.date),'Data':fmtD(ge.date),'Nome':ge.name,'Categoria':PRR_CATS.find(c=>c.id===ge.categoryId)?.label||'','Tipo':'PRR','Base (€)':Math.round(grantBase(ge)*100)/100,'IVA (%)':Number(ge.ivaRate||0),'IVA (€)':Math.round((Number(ge.amount)-grantBase(ge))*100)/100,'Total c/ IVA (€)':Number(ge.amount),'IVA Dedutível':'Sim','Pago':ge.reimbursementDate?'Reembolsado':ge.submittedDate?'Submetido':'Por submeter'}))
    ].sort((a,b)=>a['Data'].localeCompare(b['Data']))
    const wb=XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(vendasRows),'Vendas')
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(despesasRows),'Despesas')
    XLSX.writeFile(wb,`GreenDash_${mode==='month'?'mensal':'anual'}_${todayStr}.xlsx`)
    setShowExportMenu(false)
  }
  const invDocs=useMemo(()=>invoices.flatMap(inv=>(inv.attachments||[]).map(a=>({dateStr:inv.issued,label:inv.client+(inv.desc?` — ${inv.desc}`:''),file:a}))),[invoices])
  const expDocs=useMemo(()=>expenses.flatMap(exp=>(exp.attachments||[]).map(a=>({dateStr:exp.date,label:exp.name,file:a}))),[expenses])
  const prrDocs=useMemo(()=>grantExps.filter(ge=>ge.invoiceFile).map(ge=>({dateStr:ge.date,label:ge.name+(ge.supplier?` — ${ge.supplier}`:''),category:PRR_CATS.find(c=>c.id===ge.categoryId)?.label||'Sem categoria',file:ge.invoiceFile})),[grantExps])

  const saveInv=inv=>{const n={...inv,id:inv.id||Date.now(),amount:Number(inv.amount)};setInvoices(p=>p.find(i=>i.id===n.id)?p.map(i=>i.id===n.id?n:i):[...p,n]);setInvForm(null)}
  const delInv=id=>setInvoices(p=>p.filter(i=>i.id!==id))
  const toggleInvSt=id=>setInvoices(p=>p.map(i=>i.id===id?{...i,status:i.status==='paid'?'unpaid':'paid'}:i))
  const saveExp=exp=>{const n={...exp,id:exp.id||Date.now(),amount:Number(exp.amount)};setExpenses(p=>p.find(e=>e.id===n.id)?p.map(e=>e.id===n.id?n:e):[...p,n]);setExpForm(null)}
  const delExp=id=>setExpenses(p=>p.filter(e=>e.id!==id))
  const saveCommonExp=t=>setCommonExps(p=>p.find(c=>c.id===t.id)?p.map(c=>c.id===t.id?t:c):[...p,t])
  const delCommonExp=id=>setCommonExps(p=>p.filter(c=>c.id!==id))
  const toggleExpPaid=id=>setExpenses(p=>p.map(e=>e.id===id?{...e,paid:!e.paid}:e));const convertExpToGrant=(exp,lines)=>{const cleanLines=lines.filter(l=>Number(l.base||0)>0).map(l=>({base:Number(l.base),rate:Number(l.rate||0)}));const totalBase=cleanLines.reduce((s,l)=>s+Number(l.base||0),0);const totalWithIva=totalBase+linesIva(cleanLines);if(exp.id)setExpenses(p=>p.filter(e=>e.id!==exp.id));setExpForm(null);setGexpForm({id:0,grantId:'PRR',categoryId:'',name:exp.name||'',supplier:exp.cat||'',amount:totalWithIva,date:exp.date||todayStr,ivaRate:cleanLines[0]?.rate??Number(exp.ivaRate||23),invoiceFile:(exp.attachments&&exp.attachments[0])||null,submittedDate:'',expectedSubmission:'',reimbursementDate:'',reimbursementAmount:'',paid:!!exp.paid,alertTag:!!exp.alertTag,notes:exp.notes||'',createdAt:new Date().toISOString()})}
  const saveGexp=exp=>{const n={...exp,id:exp.id||Date.now(),amount:Number(exp.amount)};setGrantExps(p=>p.find(e=>e.id===n.id)?p.map(e=>e.id===n.id?n:e):[...p,n]);setGexpForm(null)}
  const delGexp=id=>setGrantExps(p=>p.filter(e=>e.id!==id))
const toggleGexpPaid=id=>setGrantExps(p=>p.map(e=>e.id===id?{...e,paid:!e.paid}:e))
const saveFutureExp=fe=>{const n={...fe,id:fe.id||Date.now()};setFutureExpenses(p=>p.find(e=>e.id===n.id)?p.map(e=>e.id===n.id?n:e):[...p,n]);setFutureForm(null)}
  const delFutureExp=id=>setFutureExpenses(p=>p.filter(e=>e.id!==id))

  const payFutureExp=(fexp,{lines,file,ivaDeductible})=>{
    const totalBase=lines.reduce((s,l)=>s+Number(l.base||0),0)
    const cleanLines=lines.filter(l=>Number(l.base||0)>0).map(l=>({base:Number(l.base),rate:Number(l.rate||0)}))
    const id=Date.now()
    const fa=file?[file]:[]
    if(fexp.isPRR){
      setGrantExps(p=>[...p,{id,grantId:'PRR',categoryId:fexp.prrCategoryId||'',name:fexp.name,supplier:fexp.cat,amount:totalBase,ivaLines:cleanLines,ivaRate:cleanLines[0]?.rate??0,date:todayStr,invoiceFile:file||null,submittedDate:'',expectedSubmission:'',reimbursementDate:'',reimbursementAmount:'',paid:true,alertTag:false,notes:'',createdAt:new Date().toISOString()}])
    }else{
      setExpenses(p=>[...p,{id,name:fexp.name,cat:fexp.cat,amount:totalBase,ivaLines:cleanLines,ivaRate:cleanLines[0]?.rate??0,type:'one-off',day:1,date:todayStr,paid:true,ivaDeductible,createdAt:new Date().toISOString(),attachments:fa}])
    }
    if(fexp.type==='recurring')setFutureExpenses(p=>p.map(e=>e.id===fexp.id?advanceDue(e):e))
    else setFutureExpenses(p=>p.filter(e=>e.id!==fexp.id))
    setPayModal(null)
  }

  const handleCfEdit=e=>{
    if(e.key.startsWith('inv-')){const id=Number(e.key.replace('inv-',''));const inv=invoices.find(i=>i.id===id);if(inv)setInvForm({...inv})}
    else if(e.key.startsWith('exp-')){const id=Number(e.key.replace('exp-',''));const exp=expenses.find(i=>i.id===id);if(exp)setExpForm({...exp})}
    else if(e.key.startsWith('ge-')){const id=Number(e.key.replace('ge-',''));const ge=grantExps.find(i=>i.id===id);if(ge)setGexpForm({...ge})}
    else if(e.key.startsWith('fe-')){const id=Number(e.key.replace('fe-',''));const fe=futureExpenses.find(i=>i.id===id);if(fe)setFutureForm({...fe})}
  }
  const saveBatch=items=>{
    items.forEach((item,idx)=>{
      const id=Date.now()+idx;const fa=item.file?[item.file]:[]
      const totalBase=item.ivaLines?.reduce((s,l)=>s+Number(l.base||0),0)||0
      const cleanLines=(item.ivaLines||[]).filter(l=>Number(l.base||0)>0).map(l=>({base:Number(l.base),rate:Number(l.rate||0)}))
      if(item.saveAs==='invoice')setInvoices(p=>[...p,{id,client:item.supplier||'Importado',desc:item.description,amount:totalBase,ivaLines:cleanLines,ivaRate:cleanLines[0]?.rate??23,issued:item.date||todayStr,due:'',status:'unpaid',createdAt:new Date().toISOString(),attachments:fa}])
      else if(item.saveAs==='grant')setGrantExps(p=>[...p,{id,grantId:'PRR',categoryId:item.grantCategoryId||'',name:item.description||'Importado',supplier:item.supplier,amount:totalBase,ivaLines:cleanLines,ivaRate:cleanLines[0]?.rate??23,date:item.date||todayStr,invoiceFile:item.file||null,submittedDate:'',expectedSubmission:'',reimbursementDate:'',reimbursementAmount:'',paid:false,alertTag:false,notes:'',createdAt:new Date().toISOString()}])
      else setExpenses(p=>[...p,{id,name:item.description||'Importado',cat:'',amount:totalBase,ivaLines:cleanLines,ivaRate:cleanLines[0]?.rate??23,type:'one-off',day:1,date:item.date||todayStr,paid:false,ivaDeductible:true,createdAt:new Date().toISOString(),attachments:fa}])
    })
    setBatchModal(false)
  }

  const visibleInv=useMemo(()=>{
let l=allTime?invoices:invoices.filter(i=>ym(i.issued)===selYM)
if(invFilter==='paid')l=l.filter(i=>i.status==='paid')
if(invFilter==='unpaid')l=l.filter(i=>i.status==='unpaid')
if(invFilter==='alert')l=l.filter(i=>i.alertTag);if(invFilter==='premio')l=l.filter(i=>i.type==='premio')
return l.sort((a,b)=>b.issued.localeCompare(a.issued))
},[invoices,selYM,invFilter,allTime])
  const oooFilter=e=>ym(e.date)===(allTime?ym(e.date):selYM)||allTime

  const filteredGexp=useMemo(()=>{
    let l=[...grantExps]
    if(gCatF!=='all')l=l.filter(e=>e.categoryId===gCatF)
    if(gStF==='no_inv')l=l.filter(e=>!e.invoiceFile)
    if(gStF==='ready')l=l.filter(e=>e.invoiceFile&&!e.submittedDate)
    if(gStF==='submitted')l=l.filter(e=>e.submittedDate&&!e.reimbursementDate)
    if(gStF==='reimbursed')l=l.filter(e=>e.reimbursementDate)
    return l.sort((a,b)=>b.date.localeCompare(a.date))
  },[grantExps,gCatF,gStF])

  const monthLabel=`${MONTHS[mo]} ${yr}`;const isNow=selYM===nowYM
  const tabSt=(k,col='#16a34a')=>({border:'none',background:'none',padding:'12px 10px',fontSize:13,fontWeight:tab===k?700:400,color:tab===k?col:'#64748b',borderBottom:tab===k?`2px solid ${col}`:'2px solid transparent',cursor:'pointer',whiteSpace:'nowrap'})
  const invAttachments=visibleInv.flatMap(i=>i.attachments||[])
  const expAttachments=expenses.flatMap(e=>e.attachments||[])
  const gexpInvoices=filteredGexp.map(ge=>ge.invoiceFile).filter(Boolean)

  return(
    <div style={{fontFamily:'Inter,system-ui,sans-serif',background:'#f8fafc',minHeight:'100vh',color:'#1e293b'}}>
      <div style={{background:'white',borderBottom:'1px solid #e2e8f0',padding:'10px 16px',display:'flex',alignItems:'center',justifyContent:'space-between',flexWrap:'wrap',gap:8}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}>
          <div style={{width:34,height:34,background:'#16a34a',borderRadius:9,display:'flex',alignItems:'center',justifyContent:'center',color:'white',fontWeight:900,fontSize:18}}>€</div>
          <span style={{fontWeight:800,fontSize:17}}>Finance Tracker</span>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:8,flexWrap:'wrap'}}>
          <button onClick={()=>setBatchModal(true)} style={{background:'#f0fdf4',color:'#16a34a',border:'1px solid #bbf7d0',borderRadius:8,padding:'7px 12px',fontSize:12,fontWeight:700,cursor:'pointer'}}>🤖 AI Import</button>
          <div style={{position:'relative'}}>
            <button onClick={()=>setShowExportMenu(p=>!p)} style={{background:'#f0fdf4',color:'#16a34a',border:'1px solid #bbf7d0',borderRadius:8,padding:'7px 12px',fontSize:12,fontWeight:700,cursor:'pointer'}}>📥 Excel</button>
            {showExportMenu&&(
              <div style={{position:'absolute',top:'calc(100% + 4px)',right:0,background:'white',borderRadius:10,boxShadow:'0 4px 20px rgba(0,0,0,0.15)',border:'1px solid #e2e8f0',zIndex:100,minWidth:180,overflow:'hidden'}}>
                <button onClick={()=>exportExcel('month')} style={{width:'100%',padding:'10px 14px',border:'none',background:'white',cursor:'pointer',textAlign:'left',fontFamily:'inherit',display:'flex',gap:10,alignItems:'center'}}>
                  <span style={{fontSize:16}}>📅</span>
                  <div><div style={{fontWeight:600,fontSize:13}}>Por mês</div><div style={{fontSize:10,color:'#94a3b8'}}>Uma linha por transação</div></div>
                </button>
                <div style={{height:1,background:'#f1f5f9'}}/>
                <button onClick={()=>exportExcel('year')} style={{width:'100%',padding:'10px 14px',border:'none',background:'white',cursor:'pointer',textAlign:'left',fontFamily:'inherit',display:'flex',gap:10,alignItems:'center'}}>
                  <span style={{fontSize:16}}>📆</span>
                  <div><div style={{fontWeight:600,fontSize:13}}>Por ano</div><div style={{fontSize:10,color:'#94a3b8'}}>Agrupado por ano</div></div>
                </button>
              </div>
            )}
          </div>
          <div style={{display:'flex',alignItems:'center',gap:6,flexWrap:'wrap'}}>
            <div style={{display:'flex',alignItems:'center',gap:6,background:'#f1f5f9',borderRadius:9,padding:'4px 8px',opacity:allTime?0.4:1}}>
              <button onClick={prevMo} disabled={allTime} style={{border:'none',background:'none',cursor:allTime?'default':'pointer',padding:'3px 8px',color:'#475569',fontSize:16,fontWeight:700}}>‹</button>
              <span style={{fontSize:13,fontWeight:700,minWidth:130,textAlign:'center'}}>{allTime?'Todo o período':monthLabel}{!allTime&&isNow?' · Agora':''}</span>
              <button onClick={nextMo} disabled={allTime} style={{border:'none',background:'none',cursor:allTime?'default':'pointer',padding:'3px 8px',color:'#475569',fontSize:16,fontWeight:700}}>›</button>
            </div>
            <button onClick={()=>setAllTime(p=>!p)} style={{border:'none',background:allTime?'#16a34a':'#f1f5f9',color:allTime?'white':'#64748b',borderRadius:9,padding:'7px 12px',fontSize:12,fontWeight:700,cursor:'pointer',whiteSpace:'nowrap'}}>
              {allTime?'✓ Tudo':'Tudo'}
            </button>
          </div>
        </div>
      </div>

      <div style={{background:'white',borderBottom:'1px solid #e2e8f0',padding:'0 16px',display:'flex',overflowX:'auto'}}>
        <button style={tabSt('dashboard')} onClick={()=>setTab('dashboard')}>📊 Dashboard</button>
        <button style={tabSt('invoices')} onClick={()=>setTab('invoices')}>💰 Vendas</button>
        <button style={tabSt('expenses')} onClick={()=>setTab('expenses')}>
          💸 Despesas{notifications.length>0&&<span style={{background:'#dc2626',color:'white',borderRadius:99,fontSize:10,padding:'1px 5px',marginLeft:4}}>{notifications.length}</span>}
        </button>
        <button style={tabSt('iva','#7c3aed')} onClick={()=>setTab('iva')}>📑 IVA</button>
        <button style={tabSt('grants','#1d4ed8')} onClick={()=>setTab('grants')}>
          🏛️ PRR{prrStats.noInvoice>0&&<span style={{background:'#dc2626',color:'white',borderRadius:99,fontSize:10,padding:'1px 5px',marginLeft:4}}>{prrStats.noInvoice}</span>}
        </button>
      </div>

      <div style={{padding:'16px',maxWidth:940,margin:'0 auto'}}>

        {/* ── DASHBOARD ── */}
        {tab==='dashboard'&&(
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(160px,1fr))',gap:10}}>
              {[
                {t:'💰 Vendas',v:fmt(calc.income),s:`${calc.mInv.filter(i=>i.status==='paid').length} faturas pagas`,c:'#16a34a',a:'#16a34a'},
                {t:'⏳ Por Receber',v:fmt(calc.pending),s:`${calc.mInv.filter(i=>i.status==='unpaid').length} por pagar`,c:'#d97706',a:'#f59e0b'},
                {t:'💸 Total Despesas',v:fmt(calc.totalExp),s:`${fmt(calc.paidExp)} pagas`,c:'#dc2626',a:'#ef4444'},
              ].map(({t,v,s,c,a})=>(
                <div key={t} style={{background:'white',borderRadius:12,padding:'13px 16px',boxShadow:'0 1px 4px rgba(0,0,0,0.07)',borderTop:`3px solid ${a}`}}>
                  <div style={{fontSize:11,color:'#64748b',fontWeight:600,marginBottom:4}}>{t}</div>
                  <div style={{fontSize:19,fontWeight:800,color:c}}>{v}</div>
                  <div style={{fontSize:11,color:'#94a3b8',marginTop:2}}>{s}</div>
                </div>
              ))}
            </div>
            {/* Notifications */}
            {notifications.length>0&&(
              <div style={{display:'flex',flexDirection:'column',gap:6}}>
                {notifications.map(fe=>{
                  const overdue=feIsOverdue(fe);const dueToday=feIsDueToday(fe)
                  return(
                    <div key={fe.id} style={{background:overdue?'#fef2f2':'#fef3c7',border:`1px solid ${overdue?'#fecaca':'#fde68a'}`,borderRadius:10,padding:'10px 14px',display:'flex',alignItems:'center',justifyContent:'space-between',gap:8,flexWrap:'wrap'}}>
                      <div>
                        <div style={{display:'flex',alignItems:'center',gap:6}}>
                          <span style={{fontSize:14}}>{overdue?'🔴':'🟡'}</span>
                          <span style={{fontWeight:700,fontSize:13,color:overdue?'#dc2626':'#92400e'}}>{fe.name}</span>
                          {fe.isPRR&&<span style={{background:'#dbeafe',color:'#1d4ed8',borderRadius:4,padding:'1px 5px',fontSize:9,fontWeight:800}}>PRR</span>}
                        </div>
                        <div style={{fontSize:11,color:'#64748b',marginTop:2}}>
                          {overdue?'Vencida em':'Vence hoje'} {fmtD(feDueDate(fe))} · {fmt(feTotal(fe))}
                          {fe.type==='recurring'&&<span style={{marginLeft:4}}>· {FREQUENCIES.find(f=>f.v===fe.frequency)?.l}</span>}
                        </div>
                      </div>
                      <div style={{display:'flex',gap:6,flexShrink:0}}>
                        <button onClick={()=>setFutureForm({...fe})} style={{border:'none',background:'#f1f5f9',color:'#475569',borderRadius:7,padding:'6px 10px',fontSize:12,cursor:'pointer'}}>✎</button>
                        <button onClick={()=>setPayModal(fe)} style={{border:'none',background:overdue?'#dc2626':'#d97706',color:'white',borderRadius:7,padding:'6px 12px',fontSize:12,fontWeight:700,cursor:'pointer'}}>💳 Pagar</button>
                      </div>
                    </div>
                  )
                })}
              </div>
            )}
            <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
              <button onClick={()=>setInvForm(blankInv())} style={{flex:1,minWidth:120,background:'white',border:'2px solid #16a34a',color:'#16a34a',borderRadius:10,padding:'10px 8px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ 💰 Venda</button>
              <button onClick={()=>setExpForm(blankExp())} style={{flex:1,minWidth:120,background:'white',border:'2px solid #dc2626',color:'#dc2626',borderRadius:10,padding:'10px 8px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ 💸 Despesa</button>
              <button onClick={()=>setGexpForm(blankGExp())} style={{flex:1,minWidth:120,background:'white',border:'2px solid #1d4ed8',color:'#1d4ed8',borderRadius:10,padding:'10px 8px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ 🏛️ PRR</button>
            </div>
            <div style={{background:'linear-gradient(135deg,#0f172a,#1e3a5f)',borderRadius:16,padding:'18px 22px',color:'white'}}>
              <div style={{fontSize:11,fontWeight:600,opacity:.6,textTransform:'uppercase',letterSpacing:1,marginBottom:4}}>Gastos previstos — {nextMoLabel}</div>
              <div style={{fontSize:36,fontWeight:900,letterSpacing:-1}}>{fmt(nextMonthTotal)}</div>
              <div style={{display:'flex',gap:18,marginTop:12,flexWrap:'wrap',alignItems:'flex-start'}}>
                <div><div style={{fontSize:10,opacity:.55,marginBottom:2}}>Em caixa agora</div><div style={{fontSize:16,fontWeight:700,color:currentBalance>=0?'#6ee7b7':'#fca5a5'}}>{fmt(currentBalance)}</div></div>
                <div style={{width:1,background:'rgba(255,255,255,0.15)',alignSelf:'stretch',minHeight:30}}/>
                {nextMonthTotal===0
                  ?<div><div style={{fontSize:10,opacity:.55,marginBottom:2}}>Situação</div><div style={{fontSize:14,fontWeight:600,opacity:.7}}>Sem gastos previstos</div></div>
                  :currentBalance>=nextMonthTotal
                    ?<div><div style={{fontSize:10,opacity:.55,marginBottom:2}}>Situação</div><div style={{fontSize:16,fontWeight:700,color:'#6ee7b7'}}>✓ Coberto</div><div style={{fontSize:11,opacity:.55,marginTop:2}}>Sobra {fmt(currentBalance-nextMonthTotal)}</div></div>
                    :<div><div style={{fontSize:10,opacity:.55,marginBottom:2}}>Em falta</div><div style={{fontSize:16,fontWeight:700,color:'#fca5a5'}}>- {fmt(nextMonthTotal-currentBalance)}</div><div style={{fontSize:11,opacity:.55,marginTop:2}}>para cobrir {nextMoLabel}</div></div>
                }
              </div>
            </div>
            {/* Cash Flow */}
            <div style={{background:'white',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,0.07)',overflow:'hidden'}}>
              <div style={{padding:'12px 16px',borderBottom:'1px solid #f1f5f9',display:'flex',justifyContent:'space-between',alignItems:'center',flexWrap:'wrap',gap:8}}>
                <div style={{fontWeight:700,fontSize:14}}>Fluxo de Caixa</div>
                <div style={{display:'flex',alignItems:'center',gap:6}}><span style={{fontSize:12,color:'#64748b'}}>Saldo:</span><span style={{fontSize:15,fontWeight:800,color:currentBalance>=0?'#16a34a':'#dc2626'}}>{fmt(currentBalance)}</span></div>
                <div style={{display:'flex',gap:5,flexWrap:'wrap'}}>
                  {[['all','Tudo'],['in','💰'],['out','💸'],['prr','PRR'],['alert','🚩'],['unsettled','⏳']].map(([k,l])=><SBtn key={k} k={k} active={cfFilter===k} onClick={()=>setCfFilter(k)}>{l}</SBtn>)}
                </div>
              </div>
              {filteredCf.length===0?<div style={{padding:'20px',textAlign:'center',color:'#94a3b8',fontSize:13}}>Sem transações</div>:(
                <>
                  {(cfExpanded?filteredCf:filteredCf.slice(0,20)).map((e,i,arr)=>(
                    <div key={e.key} style={{padding:'9px 16px',borderBottom:i<arr.length-1?'1px solid #f1f5f9':'none',opacity:e.settled?1:0.65}}>
                      <div style={{display:'flex',alignItems:'center',gap:8}}>
                        <div style={{width:28,height:28,borderRadius:7,background:e.flow==='in'?'#dcfce7':'#fee2e2',display:'flex',alignItems:'center',justifyContent:'center',fontSize:13,flexShrink:0}}>{e.flow==='in'?'↓':'↑'}</div>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{display:'flex',alignItems:'center',gap:4,flexWrap:'wrap'}}>
                            <span style={{fontSize:12,fontWeight:600}}>{e.label}</span>
                            {e.tag==='PRR'&&<span style={{background:'#f1f5f9',color:'#475569',borderRadius:4,padding:'0px 5px',fontSize:9,fontWeight:700}}>PRR</span>}
{e.tag==='future'&&<span style={{background:'#fef3c7',color:'#92400e',borderRadius:4,padding:'0px 5px',fontSize:9,fontWeight:700}}>🗓️ prev.</span>}{e.tag==='premio'&&<span style={{background:'#fef3c7',color:'#92400e',borderRadius:4,padding:'0px 5px',fontSize:9,fontWeight:700}}>🏆 Prémio</span>}
{e.alertTag&&<span style={{background:'#fee2e2',color:'#dc2626',borderRadius:4,padding:'0px 5px',fontSize:9,fontWeight:700}}>🚩 Alerta</span>}
{!e.settled&&e.tag!=='future'&&<span style={{background:'#fef3c7',color:'#d97706',borderRadius:4,padding:'0px 5px',fontSize:9,fontWeight:600}}>⏳</span>}
                          </div>
                          <div style={{fontSize:10,color:'#94a3b8'}}>{fmtD(e.date)}</div>
                        </div>
                        <div style={{textAlign:'right',flexShrink:0}}>
                          <div style={{fontSize:13,fontWeight:800,color:e.flow==='in'?'#16a34a':'#dc2626'}}>{e.flow==='in'?'+':'-'}{fmt(e.amount)}</div>
                          {e.settled&&<div style={{fontSize:10,color:'#94a3b8'}}>{fmt(e.runningBal)}</div>}
                        </div>
                        <ActBtn onClick={()=>handleCfEdit(e)}>✎</ActBtn>
                      </div>
                    </div>
                  ))}
                  {filteredCf.length>20&&<div onClick={()=>setCfExpanded(p=>!p)} style={{padding:'8px 16px',background:'#f8fafc',textAlign:'center',fontSize:12,color:'#16a34a',cursor:'pointer',fontWeight:600}}>{cfExpanded?'▲ Mostrar menos':`▼ Ver todas (${filteredCf.length})`}</div>}
                </>
              )}
            </div>
          </div>
        )}

        {/* ── VENDAS ── */}
        {tab==='invoices'&&(
          <div style={{display:'flex',flexDirection:'column',gap:16}}>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',flexWrap:'wrap',gap:8}}>
              <div style={{display:'flex',gap:5,flexWrap:'wrap'}}>
                {[['all','Todas'],['paid','Recebidas'],['unpaid','Por receber'],['premio','🏆 Prémios'],['alert','🚩 Alerta']].map(([k,l])=><SBtn key={k} k={k} active={invFilter===k} onClick={()=>setInvFilter(k)}>{l}</SBtn>)}
              </div>
              <div style={{display:'flex',gap:6,alignItems:'center'}}>
                <DlBtn count={invAttachments.length} onClick={()=>dlAll(invAttachments)}/>
                <button onClick={()=>setInvForm(blankInv())} style={{background:'#16a34a',color:'white',border:'none',borderRadius:8,padding:'8px 14px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ 💰 Venda</button>
              </div>
            </div>
            {visibleInv.length===0?(
              <div style={{textAlign:'center',padding:'40px',color:'#94a3b8',background:'white',borderRadius:12}}>
                <div style={{fontSize:32,marginBottom:6}}>💰</div><div>Sem vendas{!allTime?` em ${monthLabel}`:''}</div>
                <button onClick={()=>setInvForm(blankInv())} style={{marginTop:10,background:'#16a34a',color:'white',border:'none',borderRadius:8,padding:'8px 14px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ Adicionar</button>
              </div>
            ):(
              <div style={{background:'white',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,0.07)',overflow:'hidden'}}>
                {visibleInv.map((inv,i)=>{
                  const ls=getLines(inv).filter(l=>l.rate>0);const totalIva=linesIva(getLines(inv))
                  return(
                    <div key={inv.id} style={{padding:'12px 16px',borderBottom:i<visibleInv.length-1?'1px solid #f1f5f9':'none'}}>
                      <div style={{display:'flex',alignItems:'flex-start',gap:8}}>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{display:'flex',alignItems:'center',gap:6,flexWrap:'wrap',marginBottom:2}}>
                            <span style={{fontSize:14,fontWeight:700}}>{inv.client}</span>
                            <span style={{background:inv.status==='paid'?'#dcfce7':'#fff7ed',color:inv.status==='paid'?'#16a34a':'#d97706',padding:'2px 7px',borderRadius:99,fontSize:11,fontWeight:700}}>{inv.status==='paid'?'✓ Recebida':'⏳ Por receber'}</span>{inv.type==='premio'&&<span style={{background:'#fef3c7',color:'#92400e',padding:'2px 7px',borderRadius:99,fontSize:11,fontWeight:700}}>🏆 Prémio</span>}
                            {ls.map((l,li)=><span key={li} style={{background:'#f3e8ff',color:'#7c3aed',borderRadius:5,padding:'1px 6px',fontSize:10,fontWeight:700}}>{l.rate}%</span>)}
{inv.alertTag&&<span style={{background:'#fee2e2',color:'#dc2626',borderRadius:5,padding:'1px 6px',fontSize:10,fontWeight:700}}>🚩 Alerta</span>}
</div>
                          <div style={{fontSize:12,color:'#64748b'}}>{inv.desc}</div>
                          <div style={{fontSize:11,color:'#94a3b8',marginTop:2}}>Emissão {fmtD(inv.issued)} · Venc. {fmtD(inv.due)}</div>
                          <AttachmentChips attachments={inv.attachments}/>
                        </div>
                        <div style={{display:'flex',alignItems:'center',gap:5,flexShrink:0}}>
                          <div style={{textAlign:'right'}}>
                            <div style={{fontWeight:800,fontSize:15}}>{fmt(inv.amount)}</div>
                            {totalIva>0&&<div style={{fontSize:10,color:'#7c3aed'}}>+{fmt(totalIva)} IVA</div>}
                          </div>
                          <ActBtn onClick={()=>toggleInvSt(inv.id)}>{inv.status==='paid'?'↩':'✓'}</ActBtn>
                          <ActBtn onClick={()=>setInvForm({...inv})}>✎</ActBtn>
                          <ActBtn onClick={()=>delInv(inv.id)} bg="#fef2f2" color="#dc2626">✕</ActBtn>
                        </div>
                      </div>
                    </div>
                  )
                })}
              </div>
            )}
            <div>
              <div style={{fontWeight:700,fontSize:15,marginBottom:10}}>📁 Documentos <span style={{fontSize:12,fontWeight:400,color:'#94a3b8'}}>{invDocs?.length} ficheiro{invDocs?.length!==1?'s':''}</span></div>
              <DocsFolder items={invDocs||[]}/>
            </div>
          </div>
        )}

        {/* ── DESPESAS ── */}
        {tab==='expenses'&&(
          <div style={{display:'flex',flexDirection:'column',gap:16}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',flexWrap:'wrap',gap:8}}>
              <DlBtn count={expAttachments.length} onClick={()=>dlAll(expAttachments)}/>
              <div style={{display:'flex',gap:8}}>
                <button onClick={()=>setFutureForm(blankFutureExp())} style={{background:'#fef3c7',color:'#92400e',border:'1px solid #fde68a',borderRadius:8,padding:'8px 14px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ 🗓️ Futura</button>
                <button onClick={()=>setGexpForm(blankGExp())} style={{background:'#eff6ff',color:'#1d4ed8',border:'1px solid #bfdbfe',borderRadius:8,padding:'8px 14px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ 🏛️ PRR</button>
                <button onClick={()=>setExpForm(blankExp())} style={{background:'#16a34a',color:'white',border:'none',borderRadius:8,padding:'8px 14px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ 💸</button>
              </div>
            </div>

            {/* Despesas Futuras */}
            {(()=>{
              const sorted=[...futureExpenses].sort((a,b)=>feDueDate(a).localeCompare(feDueDate(b)))
              const total=sorted.reduce((s,fe)=>s+feTotal(fe),0)
              return(
                <div>
                  <div style={{display:'flex',justifyContent:'space-between',marginBottom:8}}>
                    <div style={{fontWeight:700,fontSize:15}}>🗓️ Despesas Futuras <span style={{fontSize:12,fontWeight:400,color:'#94a3b8'}}>{sorted.length} agendada{sorted.length!==1?'s':''}</span></div>
                    {total>0&&<span style={{fontSize:13,fontWeight:700,color:'#92400e'}}>{fmt(total)}</span>}
                  </div>
                  {sorted.length===0?(
                    <div style={{padding:'24px',textAlign:'center',color:'#94a3b8',background:'white',borderRadius:12,border:'2px dashed #e2e8f0'}}>
                      <div style={{fontSize:28,marginBottom:6}}>🗓️</div>
                      <div style={{fontSize:13,marginBottom:8}}>Sem despesas agendadas</div>
                      <button onClick={()=>setFutureForm(blankFutureExp())} style={{background:'#f59e0b',color:'white',border:'none',borderRadius:8,padding:'7px 14px',fontSize:12,fontWeight:700,cursor:'pointer'}}>+ Agendar despesa</button>
                    </div>
                  ):(
                    <div style={{background:'white',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,0.07)',overflow:'hidden'}}>
                      {sorted.map((fe,i,arr)=>{
                        const overdue=feIsOverdue(fe);const dueToday=feIsDueToday(fe);const due=overdue||dueToday
                        const freq=FREQUENCIES.find(f=>f.v===fe.frequency)
                        return(
                          <div key={fe.id} style={{padding:'11px 16px',borderBottom:i<arr.length-1?'1px solid #f1f5f9':'none',background:overdue?'#fff5f5':dueToday?'#fffbeb':'white'}}>
                            <div style={{display:'flex',alignItems:'center',gap:8}}>
                              <div style={{flex:1,minWidth:0}}>
                                <div style={{display:'flex',alignItems:'center',gap:5,flexWrap:'wrap'}}>
                                  <span style={{fontSize:13,fontWeight:600}}>{fe.name}</span>
                                  {fe.isPRR&&<span style={{background:'#dbeafe',color:'#1d4ed8',borderRadius:4,padding:'1px 5px',fontSize:9,fontWeight:800}}>PRR</span>}
                                  {overdue&&<span style={{background:'#fee2e2',color:'#dc2626',borderRadius:4,padding:'1px 5px',fontSize:9,fontWeight:700}}>VENCIDA</span>}
                                  {dueToday&&<span style={{background:'#fef3c7',color:'#92400e',borderRadius:4,padding:'1px 5px',fontSize:9,fontWeight:700}}>HOJE</span>}
                                </div>
                                <div style={{fontSize:11,color:'#64748b',marginTop:2}}>
                                  {fe.cat&&<span>{fe.cat} · </span>}
                                  {fe.type==='recurring'?<span>🔄 {freq?.l} · Dia {fe.dayOfMonth}</span>:<span>🔷 Pontual</span>}
                                  <span style={{marginLeft:4,color:overdue?'#dc2626':dueToday?'#d97706':'#94a3b8'}}>· {overdue?'Venceu':'Vence'} {fmtD(feDueDate(fe))}</span>
                                </div>
                              </div>
                              <span style={{fontWeight:700,fontSize:14,color:due?'#dc2626':'#1e293b'}}>{fmt(feTotal(fe))}</span>
                              {due&&<button onClick={()=>setPayModal(fe)} style={{border:'none',background:overdue?'#dc2626':'#d97706',color:'white',borderRadius:7,padding:'6px 10px',fontSize:11,fontWeight:700,cursor:'pointer',whiteSpace:'nowrap',flexShrink:0}}>💳 Pagar</button>}
                              <ActBtn onClick={()=>setFutureForm({...fe})}>✎</ActBtn>
                              <ActBtn onClick={()=>delFutureExp(fe.id)} bg="#fef2f2" color="#dc2626">✕</ActBtn>
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  )}
                </div>
              )
            })()}

            {/* Despesas Registadas */}
            {(()=>{
              const regList=allTime?expenses:expenses.filter(e=>ym(e.date)===selYM)
              const prrList=allTime?grantExps:grantExps.filter(e=>ym(e.date)===selYM)
              const combined=[
...regList.map(e=>({key:`exp-${e.id}`,name:e.name,sub:e.cat,date:e.date,total:Number(e.amount)+linesIva(getLines(e)),hasIva:linesIva(getLines(e))>0,isPRR:false,paid:e.paid,notes:e.notes||'',alertTag:!!e.alertTag,attachments:e.attachments||[],onEdit:()=>setExpForm({...e}),onDel:()=>delExp(e.id),onToggle:()=>toggleExpPaid(e.id)})),
...prrList.map(ge=>({key:`ge-${ge.id}`,name:ge.name,sub:PRR_CATS.find(c=>c.id===ge.categoryId)?.label||'',date:ge.date,total:Number(ge.amount),hasIva:Number(ge.ivaRate||0)>0,isPRR:true,paid:!!ge.paid,notSubmitted:!ge.submittedDate&&!ge.reimbursementDate,notes:ge.notes||'',alertTag:!!ge.alertTag,attachments:[],onEdit:()=>setGexpForm({...ge}),onDel:()=>delGexp(ge.id),onToggle:()=>toggleGexpPaid(ge.id)}))
].sort((a,b)=>b.date.localeCompare(a.date))
const filteredCombined=dExpFilter==='alert'?combined.filter(i=>i.alertTag):combined
const totalReg=filteredCombined.reduce((s,i)=>s+i.total,0)
return(
<div>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8,flexWrap:'wrap',gap:8}}>
<div style={{fontWeight:700,fontSize:15}}>💸 Despesas Registadas {!allTime&&<span style={{fontSize:12,fontWeight:400,color:'#94a3b8'}}>{monthLabel}</span>}</div>
<div style={{display:'flex',gap:5,alignItems:'center'}}>
{[['all','Todas'],['alert','🚩 Alerta']].map(([k,l])=><SBtn key={k} k={k} active={dExpFilter===k} onClick={()=>setDExpFilter(k)}>{l}</SBtn>)}
</div>
<span style={{fontSize:13,fontWeight:700,color:'#dc2626'}}>{fmt(totalReg)}</span>
</div>
{filteredCombined.length===0?(
<div style={{padding:'20px',textAlign:'center',color:'#94a3b8',background:'white',borderRadius:12,fontSize:13}}>Sem despesas registadas{!allTime?` em ${monthLabel}`:''}</div>
):(
<div style={{background:'white',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,0.07)',overflow:'hidden'}}>
{filteredCombined.map((item,i,arr)=>(
                        <div key={item.key} style={{padding:'10px 16px',borderBottom:i<arr.length-1?'1px solid #f1f5f9':'none'}}>
                          <div style={{display:'flex',alignItems:'center',gap:8}}>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontSize:13,fontWeight:600,display:'flex',alignItems:'center',gap:5}}>
{item.name}
{item.isPRR&&<span style={{background:'#dbeafe',color:'#1d4ed8',borderRadius:4,padding:'1px 6px',fontSize:9,fontWeight:800}}>PRR</span>}
{item.notSubmitted&&<span style={{background:'#fef3c7',color:'#92400e',borderRadius:4,padding:'1px 6px',fontSize:9,fontWeight:800}}>📤 Não submetido no PRR</span>}
{item.alertTag&&<span style={{background:'#fee2e2',color:'#dc2626',borderRadius:4,padding:'1px 6px',fontSize:9,fontWeight:800}}>🚩 Alerta</span>}
</div>
                              <div style={{fontSize:11,color:'#94a3b8'}}>{item.sub}{item.sub?' · ':''}{fmtD(item.date)}{item.hasIva&&<span style={{color:'#7c3aed'}}> · IVA incl.</span>}</div>{item.notes&&<div style={{fontSize:11,color:'#64748b',fontStyle:'italic',marginTop:2}}>📝 {item.notes}</div>}
                              <AttachmentChips attachments={item.attachments}/>
                            </div>
                            <span style={{fontWeight:700,fontSize:14}}>{fmt(item.total)}</span>
                            {item.onToggle
                              ?<button onClick={item.onToggle} style={{border:'none',background:item.paid?'#dcfce7':'#fff7ed',color:item.paid?'#16a34a':'#d97706',borderRadius:7,padding:'4px 9px',fontSize:11,fontWeight:700,cursor:'pointer',flexShrink:0}}>{item.paid?'✓ Pago':'⏳ Por pagar'}</button>
                              :<span style={{fontSize:11,color:item.paid?'#16a34a':'#94a3b8',fontWeight:600,flexShrink:0}}>{item.paid?'✓':'⏳'}</span>
                            }
                            <ActBtn onClick={item.onEdit}>✎</ActBtn>
                            <ActBtn onClick={item.onDel} bg="#fef2f2" color="#dc2626">✕</ActBtn>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )
            })()}

            <div>
              <div style={{fontWeight:700,fontSize:15,marginBottom:10}}>📁 Documentos <span style={{fontSize:12,fontWeight:400,color:'#94a3b8'}}>{expDocs.length} ficheiro{expDocs.length!==1?'s':''}</span></div>
              <DocsFolder items={expDocs}/>
            </div>
          </div>
        )}

        {/* ── IVA ── */}
        {tab==='iva'&&(
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <div style={{background:'white',borderRadius:12,padding:'14px 18px',boxShadow:'0 1px 4px rgba(0,0,0,0.07)',display:'flex',alignItems:'center',justifyContent:'space-between',flexWrap:'wrap',gap:10}}>
              <div>
                <div style={{fontSize:11,color:'#94a3b8',fontWeight:600,marginBottom:6}}>DECLARAÇÃO PERIÓDICA</div>
                <div style={{display:'flex',gap:6}}>
                  {[['Q1','Jan–Mar'],['Q2','Abr–Jun'],['Q3','Jul–Set'],['Q4','Out–Dez']].map(([label,months],idx)=>(
                  <button key={idx} onClick={()=>setIvaQ(idx+1)} style={{border:'none',borderRadius:8,padding:'7px 12px',fontWeight:700,fontSize:12,cursor:'pointer',background:ivaQ===idx+1?'#7c3aed':'#f3e8ff',color:ivaQ===idx+1?'white':'#7c3aed',textAlign:'center',lineHeight:1.3}}>
                    <div>{label}</div><div style={{fontSize:9,fontWeight:500,opacity:.8,marginTop:2}}>{months}</div>
                  </button>
                ))}
                </div>
              </div>
              <div style={{display:'flex',alignItems:'center',gap:8,background:'#f1f5f9',borderRadius:9,padding:'4px 10px'}}>
                <button onClick={()=>setIvaYear(y=>y-1)} style={{border:'none',background:'none',cursor:'pointer',fontSize:16,fontWeight:700,color:'#475569'}}>‹</button>
                <span style={{fontWeight:700,fontSize:14,minWidth:50,textAlign:'center'}}>{ivaYear}</span>
                <button onClick={()=>setIvaYear(y=>y+1)} style={{border:'none',background:'none',cursor:'pointer',fontSize:16,fontWeight:700,color:'#475569'}}>›</button>
              </div>
              <div style={{background:ivaStats.daysLeft<=30?'#fef2f2':ivaStats.daysLeft<=60?'#fef3c7':'#f0fdf4',borderRadius:9,padding:'8px 14px',textAlign:'center'}}>
                <div style={{fontSize:10,fontWeight:600,color:'#64748b',marginBottom:2}}>Prazo de entrega</div>
                <div style={{fontSize:13,fontWeight:800,color:ivaStats.daysLeft<=30?'#dc2626':ivaStats.daysLeft<=60?'#d97706':'#16a34a'}}>{fmtD(ivaStats.deadline.toISOString().slice(0,10))}</div>
                <div style={{fontSize:10,color:'#94a3b8'}}>{ivaStats.daysLeft>0?`${ivaStats.daysLeft} dias`:'Prazo expirado'}</div>
              </div>
            </div>
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))',gap:10}}>
              <div style={{background:'white',borderRadius:12,padding:'16px',boxShadow:'0 1px 4px rgba(0,0,0,0.07)',borderTop:'3px solid #dc2626'}}>
                <div style={{fontSize:11,color:'#64748b',fontWeight:600,marginBottom:4}}>IVA LIQUIDADO (saída)</div>
                <div style={{fontSize:22,fontWeight:800,color:'#dc2626'}}>{fmt(ivaStats.ivaOut)}</div>
                <div style={{fontSize:11,color:'#94a3b8',marginTop:2}}>Cobrado nas vendas</div>
                {ivaStats.ivaOutPending>0&&<div style={{fontSize:11,color:'#d97706',marginTop:3}}>+{fmt(ivaStats.ivaOutPending)} pendente</div>}
              </div>
              <div style={{background:'white',borderRadius:12,padding:'16px',boxShadow:'0 1px 4px rgba(0,0,0,0.07)',borderTop:'3px solid #16a34a'}}>
                <div style={{fontSize:11,color:'#64748b',fontWeight:600,marginBottom:4}}>IVA DEDUTÍVEL (entrada)</div>
                <div style={{fontSize:22,fontWeight:800,color:'#16a34a'}}>{fmt(ivaStats.ivaIn)}</div>
                <div style={{fontSize:11,color:'#94a3b8',marginTop:2}}>Dedutível nas despesas</div>
              </div>
              <div style={{background:ivaStats.ivaIn-ivaStats.ivaOut>=0?'linear-gradient(135deg,#4c1d95,#7c3aed)':'linear-gradient(135deg,#7f1d1d,#dc2626)',borderRadius:12,padding:'16px',color:'white'}}>
                <div style={{fontSize:11,fontWeight:600,opacity:.7,marginBottom:4}}>RESULTADO IVA</div>
                <div style={{fontSize:28,fontWeight:900,letterSpacing:-1}}>{fmt(ivaStats.ivaIn-ivaStats.ivaOut)}</div>
                <div style={{fontSize:11,opacity:.7,marginTop:2}}>{ivaStats.ivaIn-ivaStats.ivaOut>=0?'Crédito a seu favor':'A pagar ao Estado'}</div>
              </div>
            </div>
            <div style={{background:'white',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,0.07)',overflow:'hidden'}}>
              <div style={{padding:'12px 16px',borderBottom:'1px solid #f1f5f9',fontWeight:700,fontSize:14,display:'flex',justifyContent:'space-between'}}><span>💰 Vendas — IVA liquidado</span><span style={{color:'#dc2626'}}>{fmt(ivaStats.ivaOut+ivaStats.ivaOutPending)}</span></div>
              {ivaStats.invLines.length===0?<div style={{padding:'20px',textAlign:'center',color:'#94a3b8',fontSize:13}}>Sem vendas no Q{ivaQ} {ivaYear}</div>:
                ivaStats.invLines.map((l,i)=>(
                  <div key={i} style={{padding:'10px 16px',borderBottom:i<ivaStats.invLines.length-1?'1px solid #f1f5f9':'none',opacity:l.settled?1:0.6}}>
                    <div style={{display:'flex',alignItems:'center',gap:8}}>
                      <div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600}}>{l.label}</div><div style={{fontSize:11,color:'#94a3b8'}}>{fmtD(l.date)} · base {fmt(l.base)} · {l.rate}%</div></div>
                      <div style={{textAlign:'right'}}><div style={{fontSize:14,fontWeight:800,color:'#dc2626'}}>{fmt(l.iva)}</div>{!l.settled&&<div style={{fontSize:10,color:'#d97706',fontWeight:600}}>⏳</div>}</div>
                    </div>
                  </div>
                ))
              }
            </div>
            <div style={{background:'white',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,0.07)',overflow:'hidden'}}>
              <div style={{padding:'12px 16px',borderBottom:'1px solid #f1f5f9',fontWeight:700,fontSize:14,display:'flex',justifyContent:'space-between'}}><span>💸 Despesas — IVA dedutível</span><span style={{color:'#16a34a'}}>{fmt(ivaStats.ivaIn)}</span></div>
              {ivaStats.expLines.length===0?<div style={{padding:'20px',textAlign:'center',color:'#94a3b8',fontSize:13}}>Sem despesas com IVA dedutível no Q{ivaQ} {ivaYear}</div>:
                ivaStats.expLines.map((l,i)=>(
                  <div key={i} style={{padding:'10px 16px',borderBottom:i<ivaStats.expLines.length-1?'1px solid #f1f5f9':'none'}}>
                    <div style={{display:'flex',alignItems:'center',gap:8}}>
                      <div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600}}>{l.label}</div><div style={{fontSize:11,color:'#94a3b8'}}>{fmtD(l.date)} · base {fmt(l.base)} · {l.rate}%</div></div>
                      <div style={{fontSize:14,fontWeight:800,color:'#16a34a'}}>{fmt(l.iva)}</div>
                    </div>
                  </div>
                ))
              }
            </div>
            {(()=>{const qKey=`${ivaYear}-Q${ivaQ}`;const sentDate=ivaSubmitted[qKey];const toggle=()=>setIvaSubmitted(p=>{if(sentDate){const n={...p};delete n[qKey];return n}return{...p,[qKey]:todayStr}});return(
              <button onClick={toggle} style={{width:'100%',padding:'14px',border:`2px solid ${sentDate?'#16a34a':'#e2e8f0'}`,borderRadius:12,background:sentDate?'#f0fdf4':'white',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center',gap:10,fontFamily:'inherit'}}>
                <span style={{fontSize:20}}>{sentDate?'✅':'📤'}</span>
                <div style={{textAlign:'left'}}>
                  <div style={{fontWeight:700,fontSize:14,color:sentDate?'#16a34a':'#64748b'}}>{sentDate?`IVA Q${ivaQ} ${ivaYear} enviado ao AT`:`Marcar IVA Q${ivaQ} ${ivaYear} como enviado`}</div>
                  {sentDate&&<div style={{fontSize:11,color:'#16a34a',marginTop:1}}>Enviado em {fmtD(sentDate)} · Clique para desmarcar</div>}
                </div>
              </button>
            )})()}
            <div style={{background:'#faf5ff',border:'1px solid #e9d5ff',borderRadius:10,padding:'10px 14px',fontSize:12,color:'#6b21a8'}}>💡 SaaS estrangeiro (Figma, Vercel, etc.) pode estar sujeito a IVA por reverse charge. Confirme com o seu contabilista.</div>
          </div>
        )}

        {/* ── PRR ── */}
        {tab==='grants'&&(
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <div style={{background:'linear-gradient(135deg,#1e3a8a,#1d4ed8)',borderRadius:16,padding:'20px 22px',color:'white'}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',flexWrap:'wrap',gap:10}}>
                <div>
                  <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:4}}>
                    <div style={{background:'rgba(255,255,255,0.2)',borderRadius:7,padding:'3px 10px',fontSize:13,fontWeight:900,letterSpacing:1}}>PRR</div>
                    <span style={{fontSize:13,opacity:.7}}>Plano de Recuperação e Resiliência</span>
                  </div>
                  <div style={{fontSize:11,opacity:.55,marginBottom:2,textTransform:'uppercase',letterSpacing:.8}}>Reembolso máximo (75% da base s/ IVA)</div>
                  <div style={{fontSize:34,fontWeight:900,letterSpacing:-1}}>{fmt(PRR_REIMB)}</div>
                </div>
                <button onClick={()=>setGexpForm(blankGExp())} style={{background:'rgba(255,255,255,0.15)',color:'white',border:'1px solid rgba(255,255,255,0.3)',borderRadius:9,padding:'9px 16px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ Despesa</button>
              </div>
              <div style={{marginTop:14}}>
                <div style={{display:'flex',justifyContent:'space-between',fontSize:10,opacity:.6,marginBottom:4}}><span>Total recebido</span><span>{((prrStats.totalReceived/PRR_REIMB)*100).toFixed(1)}%</span></div>
                <div style={{background:'rgba(255,255,255,0.15)',borderRadius:99,height:8,overflow:'hidden'}}><div style={{background:'#6ee7b7',height:'100%',width:`${pct(prrStats.totalReceived,PRR_REIMB)}%`,borderRadius:99}}/></div>
              </div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(120px,1fr))',gap:10,marginTop:14}}>
                {[{l:'Orçamento total',v:fmt(PRR_BUDGET),c:'white'},{l:'Adiantamento ✅',v:fmt(PRR_ADVANCE),c:'#6ee7b7'},{l:'Base registada',v:fmt(prrStats.spent),c:'#fde68a'},{l:'Reembolsável',v:fmt(prrStats.reimbursable),c:'#bfdbfe'},{l:'Reembolsado',v:fmt(prrStats.reimbursed),c:'#6ee7b7'},{l:'Por receber',v:fmt(Math.max(0,prrStats.remaining)),c:'#fca5a5'}].map(({l,v,c})=>(
                  <div key={l} style={{background:'rgba(255,255,255,0.08)',borderRadius:9,padding:'9px 11px'}}>
                    <div style={{fontSize:9,opacity:.6,marginBottom:3,textTransform:'uppercase'}}>{l}</div>
                    <div style={{fontSize:14,fontWeight:800,color:c}}>{v}</div>
                  </div>
                ))}
              </div>
            </div>
            {prrStats.noInvoice>0&&<div style={{background:'#fef2f2',border:'1px solid #fecaca',borderRadius:10,padding:'10px 14px',display:'flex',justifyContent:'space-between',alignItems:'center'}}><span style={{fontSize:13,color:'#dc2626',fontWeight:600}}>⚠️ {prrStats.noInvoice} despesa(s) sem fatura</span><button onClick={()=>setGStF('no_inv')} style={{border:'none',background:'#dc2626',color:'white',borderRadius:6,padding:'4px 10px',fontSize:12,cursor:'pointer',fontWeight:600}}>Ver</button></div>}
            {prrStats.readyToSubmit>0&&<div style={{background:'#fef3c7',border:'1px solid #fde68a',borderRadius:10,padding:'10px 14px',display:'flex',justifyContent:'space-between',alignItems:'center'}}><span style={{fontSize:13,color:'#d97706',fontWeight:600}}>📋 {prrStats.readyToSubmit} prontas para submeter</span><button onClick={()=>setGStF('ready')} style={{border:'none',background:'#d97706',color:'white',borderRadius:6,padding:'4px 10px',fontSize:12,cursor:'pointer',fontWeight:600}}>Ver</button></div>}
            <div>
              <div style={{fontWeight:700,fontSize:15,marginBottom:10}}>📊 Categorias</div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(260px,1fr))',gap:8}}>
                {catStats.map(c=>(
                  <div key={c.id} style={{background:'white',borderRadius:10,padding:'12px 14px',boxShadow:'0 1px 3px rgba(0,0,0,0.07)',cursor:'pointer',border:gCatF===c.id?'2px solid #1d4ed8':'2px solid transparent'}} onClick={()=>setGCatF(gCatF===c.id?'all':c.id)}>
                    <div style={{display:'flex',justifyContent:'space-between',marginBottom:5}}><div style={{flex:1,marginRight:8}}><div style={{fontSize:12,fontWeight:700,lineHeight:1.3}}>{c.label}</div><div style={{fontSize:10,color:'#94a3b8',marginTop:1}}>{c.sub}</div></div><div style={{textAlign:'right',flexShrink:0}}><div style={{fontSize:13,fontWeight:800}}>{fmt(c.spent)}</div><div style={{fontSize:10,color:'#94a3b8'}}>de {fmt(c.budget)}</div></div></div>
                    <div style={{background:'#f1f5f9',borderRadius:99,height:5,overflow:'hidden',marginBottom:4}}><div style={{background:c.pct>=100?'#dc2626':c.pct>=75?'#d97706':'#1d4ed8',height:'100%',width:`${c.pct}%`,borderRadius:99}}/></div>
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:10,color:'#94a3b8'}}><span>{c.pct.toFixed(0)}% · {c.exps.length} desp.</span><span style={{color:'#16a34a'}}>reimb. {fmt(c.reimb)}</span></div>
                  </div>
                ))}
              </div>
            </div>
            <div>
              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:10,flexWrap:'wrap',gap:8}}>
                <div style={{fontWeight:700,fontSize:15}}>Registo de despesas</div>
                <div style={{display:'flex',gap:5,flexWrap:'wrap',alignItems:'center'}}>
                  <DlBtn count={gexpInvoices.length} onClick={()=>dlAll(gexpInvoices)}/>
                  {[['all','Todas'],['no_inv','⚠️ Sem fatura'],['ready','📋 Prontas'],['submitted','📤 Submetidas'],['reimbursed','✅ Reembolsadas']].map(([k,l])=><SBtn key={k} k={k} active={gStF===k} onClick={()=>setGStF(k)}>{l}</SBtn>)}
                  {gCatF!=='all'&&<button onClick={()=>setGCatF('all')} style={{border:'none',background:'#fee2e2',color:'#dc2626',padding:'5px 10px',borderRadius:20,fontSize:12,fontWeight:600,cursor:'pointer'}}>✕</button>}
                </div>
              </div>
              {filteredGexp.length===0?(
                <div style={{textAlign:'center',padding:'36px',color:'#94a3b8',background:'white',borderRadius:12}}>
                  <div style={{fontSize:30,marginBottom:6}}>🏛️</div><div style={{marginBottom:12}}>Sem despesas PRR</div>
                  <button onClick={()=>setGexpForm(blankGExp())} style={{background:'#1d4ed8',color:'white',border:'none',borderRadius:8,padding:'8px 16px',fontSize:13,fontWeight:700,cursor:'pointer'}}>+ Adicionar primeira despesa</button>
                </div>
              ):(
                <div style={{background:'white',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,0.07)',overflow:'hidden'}}>
                  {filteredGexp.map((e,i)=>{
                    const st=gexpSt(e);const cat=PRR_CATS.find(c=>c.id===e.categoryId)
                    return(
                      <div key={e.id} style={{padding:'13px 16px',borderBottom:i<filteredGexp.length-1?'1px solid #f1f5f9':'none'}}>
                        <div style={{display:'flex',alignItems:'flex-start',gap:8}}>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{display:'flex',alignItems:'center',gap:6,flexWrap:'wrap',marginBottom:2}}>
                              <span style={{fontSize:14,fontWeight:700}}>{e.name}</span>
<span style={{background:st.bg,color:st.color,padding:'2px 7px',borderRadius:99,fontSize:10,fontWeight:700}}>{st.icon} {st.label}</span>
{e.paid&&<span style={{background:'#dcfce7',color:'#16a34a',padding:'2px 7px',borderRadius:99,fontSize:10,fontWeight:700}}>✓ Pago</span>}
{e.alertTag&&<span style={{background:'#fee2e2',color:'#dc2626',padding:'2px 7px',borderRadius:99,fontSize:10,fontWeight:700}}>🚩 Alerta</span>}
</div>
                            {e.supplier&&<div style={{fontSize:12,color:'#64748b'}}>🏢 {e.supplier}</div>}
                            <div style={{fontSize:11,color:'#94a3b8'}}>{cat?.label} · {fmtD(e.date)}</div>
                            <div style={{display:'flex',gap:5,marginTop:4,flexWrap:'wrap'}}>
                              {e.invoiceFile?<span style={{background:'#dcfce7',color:'#16a34a',borderRadius:4,padding:'1px 7px',fontSize:10,fontWeight:700,cursor:'pointer'}} onClick={()=>dlAll([e.invoiceFile])}>📎 {e.invoiceFile.name} ↓</span>:<span style={{background:'#fee2e2',color:'#dc2626',borderRadius:4,padding:'1px 7px',fontSize:10,fontWeight:700,cursor:'pointer'}} onClick={()=>setGexpForm({...e})}>⚠️ Carregar fatura</span>}
                              {e.submittedDate&&<span style={{background:'#dbeafe',color:'#2563eb',borderRadius:4,padding:'1px 6px',fontSize:10,fontWeight:600}}>📤 {fmtD(e.submittedDate)}</span>}
                              {e.reimbursementDate&&<span style={{background:'#dcfce7',color:'#16a34a',borderRadius:4,padding:'1px 6px',fontSize:10,fontWeight:600}}>✅ {fmt(e.reimbursementAmount||grantReim(e))}</span>}
                            </div>
                          </div>
                          <div style={{textAlign:'right',flexShrink:0}}>
                            <div style={{fontSize:13,color:'#64748b'}}>Total: <strong>{fmt(Number(e.amount))}</strong></div>
                            <div style={{fontSize:12,color:'#475569'}}>Base: <strong>{fmt(grantBase(e))}</strong></div>
                            <div style={{fontSize:13,color:'#16a34a',fontWeight:700}}>Reimb.: {fmt(grantReim(e))}</div>
                            <div style={{display:'flex',gap:4,marginTop:6,justifyContent:'flex-end'}}>
                              <ActBtn onClick={()=>setGexpForm({...e})}>✎</ActBtn>
                              <ActBtn onClick={()=>delGexp(e.id)} bg="#fef2f2" color="#dc2626">✕</ActBtn>
                            </div>
                          </div>
                        </div>
                      </div>
                    )
                  })}
                  <div style={{padding:'10px 16px',background:'#f8fafc',display:'flex',justifyContent:'space-between',fontSize:13,fontWeight:600,borderTop:'1px solid #e2e8f0'}}>
                    <span style={{color:'#64748b'}}>{filteredGexp.length} despesa(s)</span>
                    <span>Base: <strong>{fmt(filteredGexp.reduce((s,e)=>s+grantBase(e),0))}</strong> · Reimb.: <strong style={{color:'#16a34a'}}>{fmt(filteredGexp.reduce((s,e)=>s+grantReim(e),0))}</strong></span>
                  </div>
                </div>
              )}
            </div>
            <div>
              <div style={{fontWeight:700,fontSize:15,marginBottom:10}}>📁 Documentos PRR <span style={{fontSize:12,fontWeight:400,color:'#94a3b8'}}>{prrDocs.length} fatura{prrDocs.length!==1?'s':''}</span></div>
              <DocsFolder items={prrDocs} withCategory={true}/>
            </div>
          </div>
        )}
      </div>

      {invForm&&<InvoiceModal inv={invForm} onSave={saveInv} onClose={()=>setInvForm(null)}/>}
      {expForm&&<ExpenseModal exp={expForm} onSave={saveExp} onClose={()=>setExpForm(null)} onConvert={convertExpToGrant} commonExps={commonExps} onSaveCommon={saveCommonExp} onDelCommon={delCommonExp}/>}
      {gexpForm&&<GrantExpModal exp={gexpForm} onSave={saveGexp} onClose={()=>setGexpForm(null)}/>}
      {futureForm&&<FutureExpModal fexp={futureForm} onSave={saveFutureExp} onClose={()=>setFutureForm(null)}/>}
      {payModal&&<PayFutureExpModal fexp={payModal} onPay={p=>payFutureExp(payModal,p)} onClose={()=>setPayModal(null)}/>}
      {batchModal&&<BatchUploadModal onSave={saveBatch} onClose={()=>setBatchModal(false)}/>}
    </div>
  )
}
