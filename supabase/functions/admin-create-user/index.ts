import { serve } from 'https://deno.land/std@0.224.0/http/server.ts'
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
}

const SUPABASE_URL = Deno.env.get('SUPABASE_URL')!
const SUPABASE_ANON_KEY = Deno.env.get('SUPABASE_ANON_KEY')!
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get('SERVICE_ROLE_KEY') || Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!

type Body = {
  nama: string
  nrp?: string
  email: string
  password: string
  app_id: string
  role: 'Platform Admin' | 'App Admin' | 'GL' | 'Driver' | 'Atasan Site' | 'Viewer'
  site_id?: string | null
}

const ADMIN_ROLES = ['Platform Admin', 'App Admin']
const SITE_DRIVER_CREATOR_ROLES = ['GL', 'Atasan Site', 'Site Admin', 'Admin Site']

function validEmail(email: string) {
  return /^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(email)
}

function isDuplicateError(error: any) {
  const msg = String(error?.message || error?.details || '')
  return error?.code === '23505' || msg.includes('duplicate key value') || msg.includes('uq_user_app_access_active')
}

async function findAuthUserByEmail(admin: any, email: string) {
  const target = email.toLowerCase().trim()
  let page = 1
  while (page <= 20) {
    const { data, error } = await admin.auth.admin.listUsers({ page, perPage: 1000 })
    if (error) throw error
    const found = data?.users?.find((u: any) => String(u.email || '').toLowerCase() === target)
    if (found) return found
    if (!data?.users || data.users.length < 1000) return null
    page += 1
  }
  return null
}

async function findExistingAccess(admin: any, profileId: string, body: Body) {
  let q = admin
    .from('user_app_access')
    .select('id, role, status, app_id, site_id')
    .eq('user_id', profileId)
    .eq('app_id', body.app_id)
    .eq('role', body.role)
    .eq('status', 'Aktif')

  q = body.site_id ? q.eq('site_id', body.site_id) : q.is('site_id', null)

  const { data, error } = await q.maybeSingle()
  if (error) throw error
  return data
}

serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    if (req.method !== 'POST') throw new Error('Method not allowed')
    if (!SUPABASE_URL || !SUPABASE_ANON_KEY || !SUPABASE_SERVICE_ROLE_KEY) {
      throw new Error('Edge Function env belum lengkap. Set SERVICE_ROLE_KEY, SUPABASE_URL, dan SUPABASE_ANON_KEY.')
    }

    const authHeader = req.headers.get('Authorization') || ''
    if (!authHeader) throw new Error('Unauthorized: missing Authorization header')

    const userClient = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
      global: { headers: { Authorization: authHeader } },
    })
    const admin = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)

    const { data: authData, error: authError } = await userClient.auth.getUser()
    if (authError || !authData?.user) throw new Error('Unauthorized: sesi login tidak valid')

    const callerEmail = authData.user.email
    const { data: callerProfile, error: callerProfileError } = await admin
      .from('users_profile')
      .select('id, email, nama, auth_user_id')
      .or(`auth_user_id.eq.${authData.user.id},email.eq.${callerEmail}`)
      .maybeSingle()
    if (callerProfileError) throw callerProfileError
    if (!callerProfile) throw new Error('Unauthorized: profile admin tidak ditemukan')

    if (!callerProfile.auth_user_id) {
      await admin.from('users_profile').update({ auth_user_id: authData.user.id }).eq('id', callerProfile.id)
    }

    const body = await req.json() as Body
    const email = String(body.email || '').trim().toLowerCase()
    const password = String(body.password || '')
    const nama = String(body.nama || email).trim()
    const nrp = String(body.nrp || '').trim()

    const { data: callerAccess, error: accessError } = await admin
      .from('user_app_access')
      .select('role, status, app_id, site_id')
      .eq('user_id', callerProfile.id)
      .eq('status', 'Aktif')
    if (accessError) throw accessError

    const accessRows = callerAccess || []
    const isPlatformOrAppAdmin = accessRows.some((a: any) => ADMIN_ROLES.includes(String(a.role || '').trim()))
    const isSiteDriverCreator =
      body.role === 'Driver' &&
      !!body.app_id &&
      !!body.site_id &&
      accessRows.some((a: any) =>
        SITE_DRIVER_CREATOR_ROLES.includes(String(a.role || '').trim()) &&
        String(a.app_id || '') === String(body.app_id || '') &&
        String(a.site_id || '') === String(body.site_id || '')
      )

    if (!isPlatformOrAppAdmin && !isSiteDriverCreator) {
      throw new Error('Forbidden: hanya Platform Admin / App Admin, atau admin site di site yang sama yang boleh membuat auth Driver')
    }
    if (!email) throw new Error('Email wajib diisi')
    if (!validEmail(email)) throw new Error('Format email tidak valid. Gunakan format nama@domain.com.')
    if (!password || password.length < 6) throw new Error('Password minimal 6 karakter')
    if (!body.app_id) throw new Error('Aplikasi wajib dipilih')
    if (!body.role) throw new Error('Role wajib dipilih')

    const { data: targetApp, error: targetAppError } = await admin
      .from('applications')
      .select('id, app_code, app_name, status')
      .eq('id', body.app_id)
      .maybeSingle()
    if (targetAppError) throw targetAppError
    if (!targetApp) throw new Error('Aplikasi tidak ditemukan. Cek mapping applications/app_id.')
    if (String(targetApp.status || '').toLowerCase() === 'nonaktif') throw new Error('Aplikasi sedang nonaktif.')
    if (email === String(callerEmail || '').toLowerCase()) throw new Error('Email user baru tidak boleh sama dengan email admin yang sedang login. Gunakan tambah mapping user existing untuk akun admin.')

    let authAction = 'existing_user'
    let authUser = await findAuthUserByEmail(admin, email)
    if (!authUser) {
      const { data: created, error: createError } = await admin.auth.admin.createUser({
        email,
        password,
        email_confirm: true,
        user_metadata: { nama, nrp },
      })
      if (createError) throw createError
      authUser = created.user
      authAction = 'created_user'
    } else {
      const { data: updated, error: updateError } = await admin.auth.admin.updateUserById(authUser.id, {
        password,
        email_confirm: true,
        user_metadata: { ...(authUser.user_metadata || {}), nama, nrp },
      })
      if (updateError) throw updateError
      authUser = updated.user
      authAction = 'updated_existing_user'
    }

    const { data: profile, error: profileError } = await admin
      .from('users_profile')
      .upsert({ auth_user_id: authUser.id, email, nama, nrp, status: 'Aktif' }, { onConflict: 'email' })
      .select()
      .single()
    if (profileError) throw profileError

    let mapping = await findExistingAccess(admin, profile.id, body)
    let mappingAction = mapping ? 'access_already_exists' : 'created_access'

    if (!mapping) {
      const { data: insertedMapping, error: mappingError } = await admin
        .from('user_app_access')
        .insert({
          user_id: profile.id,
          app_id: body.app_id,
          role: body.role,
          site_id: body.site_id || null,
          status: 'Aktif',
        })
        .select('id, role, status, app_id, site_id')
        .single()

      if (mappingError) {
        if (!isDuplicateError(mappingError)) throw mappingError
        mapping = await findExistingAccess(admin, profile.id, body)
        if (!mapping) throw mappingError
        mappingAction = 'access_already_exists'
      } else {
        mapping = insertedMapping
      }
    }

    if (!authUser?.id) throw new Error('Supabase Auth tidak mengembalikan auth user id.')

    return new Response(JSON.stringify({ ok: true, auth_user_id: authUser.id, profile, mapping, app: targetApp, auth_action: authAction, mapping_action: mappingAction }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 200,
    })
  } catch (err) {
    return new Response(JSON.stringify({ ok: false, error: err?.message || String(err) }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 400,
    })
  }
})
