import { createClient } from '@supabase/supabase-js'

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || ''
const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || ''

export const supabase = createClient(supabaseUrl, supabaseKey)

export async function storageGet(key: string) {
  try {
    const { data, error } = await supabase
      .from('finance_data')
      .select('value')
      .eq('key', key)
      .single()
    
    return error ? null : { key, value: data?.value }
  } catch (e) {
    return null
  }
}

export async function storageSet(key: string, value: string) {
  try {
    await supabase.from('finance_data').delete().eq('key', key)
    const { error } = await supabase
      .from('finance_data')
      .insert([{ key, value }])
    return !error
  } catch (e) {
    return false
  }
}
