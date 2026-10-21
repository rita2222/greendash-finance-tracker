import { createClient } from '@supabase/supabase-js'
import { FinanceStore } from './financeStore'

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || ''
const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || ''

export const supabase = createClient(supabaseUrl, supabaseKey)

// Único ponto de acesso aos dados financeiros (ver lib/financeStore.ts).
export const financeStore = new FinanceStore(supabase)
