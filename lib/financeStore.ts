// Camada de persistência do Finance Tracker (tabela `finance_data`: uma linha por chave).
//
// Garantias (existem para que nunca mais se apaguem dados sem querer):
//  1. Ler distingue explicitamente "sem dados" de "erro". Um erro NUNCA é tratado como lista vazia.
//  2. Nunca se grava sem antes ter lido a chave com sucesso (é preciso conhecer a versão).
//  3. Gravar é um "compare-and-swap" atómico: só atualiza se a versão (updated_at) ainda for a
//     que lemos. Se outro separador/dispositivo gravou entretanto, devolve 'conflict' em vez de
//     sobrescrever. Não há "apagar e inserir": não existe janela em que a linha fica vazia.
//  4. Se existirem linhas duplicadas para a mesma chave, lê-se a mais recente (não falha).

export type LoadResult =
  | { status: 'ok'; value: string }
  | { status: 'empty' }
  | { status: 'error'; message: string }

export type SaveResult =
  | { status: 'ok' }
  | { status: 'conflict' }
  | { status: 'error'; message: string }

const TABLE = 'finance_data'

type Known = { exists: boolean; ts: string | null }

const errMsg = (e: any): string => (e && (e.message || e.details || e.hint)) || String(e || 'erro desconhecido')

export class FinanceStore {
  private known = new Map<string, Known>()

  // `db` é um cliente @supabase/supabase-js (ou um dublê com a mesma API, nos testes).
  constructor(private db: any) {}

  async load(key: string): Promise<LoadResult> {
    try {
      const { data, error } = await this.db
        .from(TABLE)
        .select('value, updated_at')
        .eq('key', key)
        .order('updated_at', { ascending: false, nullsFirst: false })
        .limit(1)
      if (error) return { status: 'error', message: errMsg(error) }
      if (!Array.isArray(data)) return { status: 'error', message: 'Resposta inesperada do servidor' }
      if (data.length === 0) {
        this.known.set(key, { exists: false, ts: null })
        return { status: 'empty' }
      }
      const row = data[0]
      if (typeof row.value !== 'string') {
        return { status: 'error', message: `Valor inválido na chave ${key}` }
      }
      this.known.set(key, { exists: true, ts: row.updated_at ?? null })
      return { status: 'ok', value: row.value }
    } catch (e) {
      return { status: 'error', message: errMsg(e) }
    }
  }

  async save(key: string, value: string): Promise<SaveResult> {
    const k = this.known.get(key)
    if (!k) return { status: 'error', message: `A chave ${key} nunca foi lida; gravação recusada` }
    const ts = nextTimestamp(k.ts)
    try {
      if (!k.exists) {
        const { data, error } = await this.db
          .from(TABLE)
          .insert({ key, value, updated_at: ts })
          .select('updated_at')
        if (error) {
          // 23505 = violação de unicidade: alguém criou a linha entretanto
          if ((error as any).code === '23505') return { status: 'conflict' }
          return { status: 'error', message: errMsg(error) }
        }
        this.known.set(key, { exists: true, ts: data?.[0]?.updated_at ?? ts })
        return { status: 'ok' }
      }
      let q = this.db.from(TABLE).update({ value, updated_at: ts }).eq('key', key)
      q = k.ts === null ? q.is('updated_at', null) : q.eq('updated_at', k.ts)
      const { data, error } = await q.select('updated_at')
      if (error) return { status: 'error', message: errMsg(error) }
      if (!Array.isArray(data) || data.length === 0) return { status: 'conflict' }
      this.known.set(key, { exists: true, ts: data[0]?.updated_at ?? ts })
      return { status: 'ok' }
    } catch (e) {
      return { status: 'error', message: errMsg(e) }
    }
  }
}

// Novo valor de updated_at, sempre posterior ao anterior (o relógio do browser pode estar atrasado).
function nextTimestamp(prev: string | null): string {
  const p = prev ? Date.parse(prev) : 0
  let ms = Date.now()
  if (!Number.isNaN(p) && ms <= p) ms = p + 1
  return new Date(ms).toISOString()
}
