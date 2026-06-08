'use client'
import { useState, useEffect, useMemo, useRef } from "react"
import * as XLSX from 'xlsx'
import { createClient } from '@supabase/supabase-js'

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || ''
const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || ''
const supabase = createClient(supabaseUrl, supabaseKey)

async function storageGetAsync(key: string) {
  try {
    const { data, error } = await supabase.from('finance_data').select('value').eq('key', key).single()
    return error ? null : { key, value: data?.value }
  } catch (e) { return null }
}

async function storageSetAsync(key: string, value: string) {
  try {
    await supabase.from('finance_data').delete().eq('key', key)
    const { error } = await supabase.from('finance_data').insert([{ key, value }])
    return !error
  } catch (e) { return false }
}

export default function FinanceTracker(){
  const[loading,setLoading]=useState(true)
  const[inv,setInv]=useState([])
  const[exp,setExp]=useState([])
  const[gexp,setGexp]=useState([])

  useEffect(()=>{
    const load=async()=>{
      try {
        const [a,b,c] = await Promise.all([storageGetAsync('ft_inv'),storageGetAsync('ft_exp'),storageGetAsync('ft_gexp')])
        if(a?.value) setInv(JSON.parse(a.value))
        if(b?.value) setExp(JSON.parse(b.value))
        if(c?.value) setGexp(JSON.parse(c.value))
      } catch (e) {}
      setLoading(false)
    }
    load()
  }, [])

  useEffect(()=>{
    if(loading) return
    const save=async()=>{
      try {
        await Promise.all([storageSetAsync('ft_inv',JSON.stringify(inv)),storageSetAsync('ft_exp',JSON.stringify(exp)),storageSetAsync('ft_gexp',JSON.stringify(gexp))])
      } catch (e) {}
    }
    save()
  }, [inv,exp,gexp,loading])

  if(loading) return <div style={{padding:20}}>⏳ Carregando de Supabase...</div>
  
  return <div style={{padding:20}}><h1>✅ Finance Tracker com Supabase</h1><p>Data: {inv.length} vendas, {exp.length} despesas, {gexp.length} PRR</p><p style={{color:'#16a34a',fontWeight:700}}>✓ Tudo salvo em Supabase!</p></div>
}
