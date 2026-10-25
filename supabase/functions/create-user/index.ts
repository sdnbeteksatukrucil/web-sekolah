// Edge Function: admin membuat akun login guru/siswa (username + password)
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
const cors = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type' }
const out = (o: unknown, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { ...cors, 'Content-Type': 'application/json' } })
Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: cors })
  const db = createClient(Deno.env.get('SUPABASE_URL')!, Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!)
  const { data: { user } } = await db.auth.getUser((req.headers.get('Authorization') || '').replace('Bearer ', ''))
  if (!user) return out({ error: 'Belum login' }, 401)
  const { data: me } = await db.from('profiles').select('role').eq('id', user.id).single()
  if (me?.role !== 'admin') return out({ error: 'Hanya admin' }, 403)
  const b = await req.json()
  if (!b.username || !b.password || b.password.length < 6) return out({ error: 'Username wajib, password minimal 6 karakter' }, 400)
  if (!['guru', 'siswa'].includes(b.role)) return out({ error: 'Role tidak valid' }, 400)
  const { data, error } = await db.auth.admin.createUser({ email: `${String(b.username).trim().toLowerCase()}@betek1.local`, password: b.password, email_confirm: true })
  if (error) return out({ error: error.message }, 400)
  const { error: e2 } = await db.from('profiles').insert({ id: data.user.id, role: b.role, nama: b.nama, kelas: b.kelas ?? null, siswa_id: b.siswa_id ?? null, guru_id: b.guru_id ?? null })
  if (e2) { await db.auth.admin.deleteUser(data.user.id); return out({ error: e2.message }, 400) }
  return out({ ok: true })
})
