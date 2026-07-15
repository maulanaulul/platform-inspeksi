import React, { useEffect, useMemo, useRef, useState } from 'react'
import { supabase } from './lib/supabase'
import { compressImage } from './utils/imageCompress'
import {
  AlertTriangle, BarChart3, Building2, CalendarCheck, Check, CheckCircle2, ClipboardCheck,
  Database, Download, Eye, FileSpreadsheet, FileText, KeyRound, LayoutDashboard, LogOut,
  MapPin, Menu, ParkingCircle, Search, ShieldCheck, Truck, Upload, Users, X, XCircle
} from 'lucide-react'
import * as XLSX from 'xlsx'
import logoSrgs from './assets/logo-icon.png'
import SharedAdminPanel from './SharedAdminPanel.jsx'
import './styles.css'

const APP_CODE = 'inspeksi_unit'
const ADMIN_ROLES = ['Platform Admin', 'App Admin']
const APPROVAL_ROLES = ['Platform Admin', 'App Admin', 'Atasan Site']
const CATEGORIES = ['Inspeksi Unit', 'Inspeksi Kelayakan Parkiran', 'PM Check']
const MONTHS = [
  ['1','Januari'], ['2','Februari'], ['3','Maret'], ['4','April'], ['5','Mei'], ['6','Juni'],
  ['7','Juli'], ['8','Agustus'], ['9','September'], ['10','Oktober'], ['11','November'], ['12','Desember']
]

const AUTO_CODE_LABEL = 'Otomatis oleh sistem'
function isUnitBasedCategory(category){
  return category === 'Inspeksi Unit' || category === 'PM Check'
}
function targetTypeByCategory(category){
  return isUnitBasedCategory(category) ? 'unit' : 'parkiran'
}
function parameterPrefixByCategory(category){
  if (category === 'Inspeksi Kelayakan Parkiran') return 'PPARK'
  if (category === 'PM Check') return 'PPM'
  return 'PUNIT'
}
function pmChecklistLabel(parameter, plan){
  const raw = parameter?.parameter_name || ''
  if (parameter?.category !== 'PM Check') return raw
  const due = plan?.due_date ? new Date(plan.due_date).toLocaleDateString('id-ID') : 'sesuai tanggal planning'
  return raw.replace('tanggal xxx', `tanggal ${due}`)
}
function generatedPreviewCode(prefix, rowIndex){
  return `${prefix}-AUTO-${String(rowIndex).padStart(3, '0')}`
}
function generatedRecordCode(prefix){
  const fallback = `${Date.now()}${Math.random().toString(36).slice(2, 10)}`
  const raw = globalThis.crypto?.randomUUID?.() || fallback
  return `${prefix}-${String(raw).replace(/-/g, '').slice(0, 8).toUpperCase()}`
}

function clean(v){ return String(v ?? '').trim() }
function normEmail(v){ return clean(v).toLowerCase() }
function today(){ return new Date().toISOString().slice(0, 10) }
function nowISO(){ return new Date().toISOString() }
function pad2(v){ return String(v).padStart(2, '0') }
function isValidYmd(y, m, d){
  const dt = new Date(Date.UTC(y, m - 1, d))
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d
}
function formatYmd(y, m, d){ return `${y}-${pad2(m)}-${pad2(d)}` }
function dateToYmd(date){
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return ''
  return formatYmd(date.getFullYear(), date.getMonth() + 1, date.getDate())
}
function normalizeTwoDigitYear(y){ return y < 100 ? (y >= 70 ? 1900 + y : 2000 + y) : y }
function excelSerialDateToYmd(serial){
  const n = Number(serial)
  if (!Number.isFinite(n) || n <= 0) return ''
  const wholeDays = Math.floor(n)
  const timeMs = Math.round((n - wholeDays) * 86400000)
  const dt = new Date(Date.UTC(1899, 11, 30) + wholeDays * 86400000 + timeMs)
  return formatYmd(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate())
}
function parseFlexibleDate(value){
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return dateToYmd(value)
  if (typeof value === 'number') return excelSerialDateToYmd(value)

  const raw = clean(value)
  if (!raw || raw === '-') return ''

  if (/^\d+(\.\d+)?$/.test(raw)) {
    const n = Number(raw)
    if (n > 20000 && n < 80000) return excelSerialDateToYmd(n)
  }

  const iso = raw.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/)
  if (iso) {
    const y = Number(iso[1]), m = Number(iso[2]), d = Number(iso[3])
    return isValidYmd(y, m, d) ? formatYmd(y, m, d) : ''
  }

  const short = raw.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/)
  if (short) {
    const a = Number(short[1]), b = Number(short[2]), y = normalizeTwoDigitYear(Number(short[3]))
    const tryDmy = isValidYmd(y, b, a) ? formatYmd(y, b, a) : ''
    const tryMdy = isValidYmd(y, a, b) ? formatYmd(y, a, b) : ''
    if (a > 12) return tryDmy
    if (b > 12) return tryMdy
    return tryDmy || tryMdy
  }

  const native = new Date(raw)
  if (!Number.isNaN(native.getTime())) return dateToYmd(native)
  return ''
}
function normalizeHeaderKey(k){
  return clean(k).toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '')
}
function canAdmin(role){ return ADMIN_ROLES.includes(role) }
function canApprove(role){ return APPROVAL_ROLES.includes(role) }
function isHeadOfficeAdmin(context){ return canAdmin(context?.role) && (!context?.site_id || context?.sites?.site_code === 'JIEP') }
function scopeQuery(query, context){ return isHeadOfficeAdmin(context) ? query : query.eq('site_id', context.site_id) }
function siteName(context){ return context?.sites?.site_name || (isHeadOfficeAdmin(context) ? 'All Site' : '-') }
function targetParts(plan){
  if (!plan) return { title:'-', code:'', type:'' }
  const isUnit = plan.target_type === 'unit' || plan.unit_id
  const master = isUnit ? plan.inspection_units : plan.inspection_parkings
  const name = clean(isUnit ? master?.unit_name : master?.parking_name)
  const code = clean(isUnit ? master?.unit_code : master?.parking_code)
  const title = name || code || '-'
  const shouldShowCode = code && normalizeTargetKey(code) !== normalizeTargetKey(title)
  return { title, code: shouldShowCode ? code : '', type: isUnit ? 'Unit' : 'Parkiran' }
}
function targetLabel(plan){
  const t = targetParts(plan)
  return t.code ? `${t.title} (${t.code})` : t.title
}
function targetTitle(plan){ return targetParts(plan).title }
function targetCodeLabel(plan){ return targetParts(plan).code }
function monthLabel(v){ return MONTHS.find(m => String(m[0]) === String(v))?.[1] || v }
function planYear(p){ return p?.created_at ? new Date(p.created_at).getFullYear() : p?.tahun }
function normalizeRow(row){
  const out = {}
  Object.keys(row || {}).forEach(k => {
    const originalKey = String(k).trim()
    const normalizedKey = normalizeHeaderKey(originalKey)
    out[originalKey] = row[k]
    if (normalizedKey && out[normalizedKey] === undefined) out[normalizedKey] = row[k]
  })
  return out
}
function downloadXlsx(filename, rows, sheetName = 'Data'){
  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet(rows?.length ? rows : [{}])
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  XLSX.writeFile(wb, filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`)
}
function downloadTemplate(type){
  const samples = {
    unit: [{ site_code:'BAYA', unit_name:'Dump Truck 001', unit_type:'Dump Truck', location:'Pit Area', status:'Aktif' }],
    parking: [{ site_code:'BAYA', parking_name:'Parkiran Barat', location:'Area Barat', capacity:30, status:'Aktif' }],
    parameter: [{ category:'PM Check', parameter_name:'Apakah PM check telah dilakukan pada tanggal xxx (sesuai yang di planning)', description:'Pilih Aman jika PM check sudah dilakukan sesuai tanggal planning. Pilih Tidak Aman jika belum dilakukan / tidak sesuai jadwal.', severity:'High', status:'Aktif' }],
    plan: [{ site_code:'BAYA', category:'PM Check', target_name:'Dump Truck 001', target_code:'UNIT-0001', bulan:5, due_date:'2026-05-31' }],
    access: [{ nama:'Nama User', nrp:'NRP-001', email:'user@company.co.id', password:'password123', app_code:'inspeksi_unit', role:'GL', site_code:'BAYA' }]
  }
  downloadXlsx(`template-${type}.xlsx`, samples[type] || [{}], 'Template')
}
async function parseExcel(file){
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true }).map(normalizeRow)
}

async function fetchAllRows(table, select = '*', buildQuery = q => q, pageSize = 1000){
  let page = 0
  let all = []
  while (true) {
    const from = page * pageSize
    const to = from + pageSize - 1
    let query = supabase.from(table).select(select).range(from, to)
    query = buildQuery(query)
    const { data, error } = await query
    if (error) throw error
    const rows = data || []
    all = all.concat(rows)
    if (rows.length < pageSize) break
    page += 1
  }
  return all
}
function normalizeTargetKey(v){
  return clean(v)
    .toUpperCase()
    .replace(/\s+/g, '')
    .replace(/[^A-Z0-9]/g, '')
}
function isActiveMasterStatus(status){
  const key = normalizeTargetKey(status)
  return !key || key === 'AKTIF' || key === 'ACTIVE'
}
function sortBySiteAndName(a, b, nameField, codeField){
  const siteA = a.sites?.site_code || a.site_code || ''
  const siteB = b.sites?.site_code || b.site_code || ''
  return siteA.localeCompare(siteB) || clean(a?.[nameField]).localeCompare(clean(b?.[nameField])) || clean(a?.[codeField]).localeCompare(clean(b?.[codeField]))
}
function targetMatches(master, searchName, searchCode, codeField, nameField){
  const keys = [
    normalizeTargetKey(master?.[codeField]),
    normalizeTargetKey(master?.[nameField])
  ].filter(Boolean)
  const searches = [normalizeTargetKey(searchName), normalizeTargetKey(searchCode)].filter(Boolean)
  if (!searches.length) return false
  return searches.some(s => keys.includes(s) || keys.some(k => k && (k.includes(s) || s.includes(k))))
}
async function uploadCompressedImage(bucket, file, folder = 'uploads'){
  if (!file) return null
  const compressed = await compressImage(file)
  const safeName = clean(file.name).replace(/[^a-zA-Z0-9_.-]/g, '-') || 'photo.jpg'
  const path = `${folder}/${Date.now()}-${Math.random().toString(36).slice(2)}-${safeName.replace(/\.[^.]+$/, '.jpg')}`
  const { error } = await supabase.storage.from(bucket).upload(path, compressed, { upsert: false, contentType: 'image/jpeg' })
  if (error) throw error
  const { data } = supabase.storage.from(bucket).getPublicUrl(path)
  return data.publicUrl
}

function withTimeout(promise, ms, label){
  return Promise.race([
    promise,
    new Promise((_, reject) => window.setTimeout(() => reject(new Error(label)), ms))
  ])
}


async function hardLogout(){
  try { await supabase.auth.signOut({ scope: 'local' }) } catch(e) { console.warn('[LOGOUT WARNING]', e) }
  try { window.localStorage.clear(); window.sessionStorage.clear() } catch(e) {}
  window.location.replace('/')
}

function App({ embeddedProfile = null, embeddedContext = null, onChangeApp = null } = {}){
  const [session, setSession] = useState(null)
  const [profile, setProfile] = useState(null)
  const [access, setAccess] = useState([])
  const [context, setContext] = useState(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [dateFrom, setDateFrom] = useState('')
  const [dateTo, setDateTo] = useState('')
  const mountedRef = useRef(false)
  const bootIdRef = useRef(0)

  useEffect(() => {
    if (embeddedProfile && embeddedContext) { setLoading(false); return }
    mountedRef.current = true
    init('boot')

    const { data: sub } = supabase.auth.onAuthStateChange((event) => {
      // Hindari query berat langsung di callback. Cukup re-init saat status auth berubah.
      if (event === 'SIGNED_OUT') {
        setSession(null); setProfile(null); setAccess([]); setContext(null); setLoading(false); setError('')
        return
      }
      if (event === 'SIGNED_IN' || event === 'TOKEN_REFRESHED' || event === 'USER_UPDATED') {
        window.setTimeout(() => init(event), 0)
      }
    })

    return () => {
      mountedRef.current = false
      sub?.subscription?.unsubscribe?.()
    }
  }, [])

  async function init(source = 'boot'){
    const bootId = ++bootIdRef.current
    const stillCurrent = () => mountedRef.current && bootId === bootIdRef.current
    const failHint = 'Booting aplikasi gagal/terlalu lama. Cek .env, koneksi Supabase, SQL Phase 4, lalu clear cache browser.'
    try{
      if (stillCurrent()) { setLoading(true); setError('') }

      const sessionResult = await withTimeout(
        supabase.auth.getSession(),
        8000,
        `${failHint} Detail: Supabase auth getSession timeout.`
      )
      if (!stillCurrent()) return

      const currentSession = sessionResult?.data?.session || null
      setSession(currentSession)

      if (!currentSession?.user){
        setProfile(null); setAccess([]); setContext(null)
        return
      }

      let profileResult = await withTimeout(
        supabase.from('users_profile').select('*').eq('auth_user_id', currentSession.user.id).maybeSingle(),
        8000,
        `${failHint} Detail: query users_profile timeout.`
      )
      if (!stillCurrent()) return
      if (profileResult.error) throw profileResult.error
      let prof = profileResult.data

      if (!prof && currentSession.user.email){
        const emailResult = await withTimeout(
          supabase.from('users_profile').select('*').eq('email', normEmail(currentSession.user.email)).maybeSingle(),
          8000,
          `${failHint} Detail: query users_profile by email timeout.`
        )
        if (!stillCurrent()) return
        if (emailResult.error) throw emailResult.error
        if (emailResult.data){
          const updateResult = await withTimeout(
            supabase.from('users_profile').update({ auth_user_id: currentSession.user.id }).eq('id', emailResult.data.id).select('*').single(),
            8000,
            `${failHint} Detail: update auth_user_id timeout.`
          )
          if (!stillCurrent()) return
          if (updateResult.error) throw updateResult.error
          prof = updateResult.data
        }
      }

      if (!prof){
        setProfile(null); setAccess([]); setContext(null)
        setError('Akun Auth sudah ada, tetapi email ini belum dibuat/mapping di Admin Panel.')
        return
      }
      setProfile(prof)

      const accessResult = await withTimeout(
        supabase
          .from('user_app_access')
          .select('*, applications(app_code, app_name, description), sites(site_code, site_name, region)')
          .eq('user_id', prof.id)
          .eq('status', 'Aktif'),
        8000,
        `${failHint} Detail: query user_app_access timeout.`
      )
      if (!stillCurrent()) return
      if (accessResult.error) throw accessResult.error

      const rows = accessResult.data || []
      setAccess(rows)
      const inspeksiAccess = rows.filter(r => r.applications?.app_code === APP_CODE)
      if (inspeksiAccess.length === 1) setContext(inspeksiAccess[0])
      else setContext(null)
    }catch(e){
      console.error('[BOOT ERROR]', source, e)
      if (stillCurrent()) setError(e.message || String(e))
    }finally{
      if (stillCurrent()) setLoading(false)
    }
  }

  async function resetLocalSession(){
    await hardLogout()
  }

  if (embeddedProfile && embeddedContext) return <Shell profile={embeddedProfile} context={embeddedContext} setContext={() => onChangeApp ? onChangeApp() : null} />
  if (loading) return <FullCenter text="Memuat aplikasi..." />
  if (!session) return <Login />
  if (error) return <FullCenter text={error} action={<div className="actions-inline"><button onClick={() => init('manual-retry')}>Coba Lagi</button><button className="secondary" onClick={resetLocalSession}>Reset Login</button></div>} />
  if (!profile) return <FullCenter text="Profile belum ditemukan." action={<button onClick={resetLocalSession}>Reset Login</button>} />
  if (!context) return <ContextPicker profile={profile} access={access} onSelect={setContext} />
  return <Shell profile={profile} context={context} setContext={setContext} />
}

function Login(){
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  async function submit(e){
    e.preventDefault()
    setBusy(true); setMessage('')
    const { error } = await supabase.auth.signInWithPassword({ email: normEmail(email), password })
    if (error) setMessage(error.message)
    setBusy(false)
  }
  return <section className="login-page">
    <div className="login-hero">
      <div className="brand"><div className="brand-logo">✓</div><div><b>Platform Inspeksi SRGS</b><span>Apps Inspeksi Phase 4</span></div></div>
      <h1>Inspeksi unit, kelayakan parkiran, PM check, approval, outstanding, dan dashboard all site.</h1>
      <p>User dibuat dari Admin Panel. Login email/password, lalu pilih aplikasi, role, dan site sesuai mapping akses.</p>
      <div className="chips"><span>React + Supabase</span><span>Corporate Premium</span><span>Excel Import</span><span>Photo Compress</span></div>
    </div>
    <form className="login-card" onSubmit={submit}>
      <KeyRound size={40} color="#2563eb" />
      <h2>Masuk</h2>
      <p>Gunakan akun yang sudah dibuat Administrator.</p>
      <label>Email<input required type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="nama@company.co.id" /></label>
      <label>Password<input required type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Password" /></label>
      {message && <div className="error">{message}</div>}
      <button disabled={busy}>{busy ? 'Memproses...' : 'Masuk'}</button>
    </form>
  </section>
}

function ContextPicker({ profile, access, onSelect }){
  const validAccess = access.filter(a => a.applications?.app_code === APP_CODE)
  const [appId, setAppId] = useState('')
  const [role, setRole] = useState('')
  const [siteId, setSiteId] = useState('')

  useEffect(() => {
    if (validAccess.length){
      setAppId(validAccess[0].app_id || '')
      setRole(validAccess[0].role || '')
      setSiteId(validAccess[0].site_id || '')
    }
  }, [access.length])

  const appOptions = Array.from(new Map(validAccess.map(a => [a.app_id, a])).values())
  const roleOptions = Array.from(new Set(validAccess.filter(a => !appId || a.app_id === appId).map(a => a.role)))
  const siteOptions = Array.from(new Map(validAccess.filter(a => (!appId || a.app_id === appId) && (!role || a.role === role)).map(a => [a.site_id || 'all', a])).values())

  useEffect(() => {
    const first = validAccess.find(a => a.app_id === appId)
    if (first){ setRole(first.role || ''); setSiteId(first.site_id || '') }
  }, [appId])
  useEffect(() => {
    const first = validAccess.find(a => a.app_id === appId && a.role === role)
    if (first) setSiteId(first.site_id || '')
  }, [role])

  function start(){
    const selected = validAccess.find(a => String(a.app_id) === String(appId) && String(a.role) === String(role) && String(a.site_id || '') === String(siteId || ''))
    if (!selected) return alert('Akses tidak ditemukan. Cek mapping aplikasi, role, dan site di Admin Panel.')
    onSelect(selected)
  }

  return <section className="context-page">
    <div className="context-card">
      <div className="brand dark"><div className="brand-logo">✓</div><div><b>Platform Inspeksi SRGS</b><span>Pilih sesi kerja</span></div></div>
      <h1>Pilih aplikasi, role, dan site</h1>
      <p>User: <b>{profile?.nama}</b> · {profile?.email}</p>
      {!validAccess.length ? <div className="empty"><AlertTriangle/> Email ini belum punya akses ke Apps Inspeksi.</div> : <>
        <div className="form-grid three">
          <label>Aplikasi<select value={appId} onChange={e=>setAppId(e.target.value)}>{appOptions.map(a=><option key={a.app_id} value={a.app_id}>{a.applications?.app_name}</option>)}</select></label>
          <label>Role<select value={role} onChange={e=>setRole(e.target.value)}>{roleOptions.map(r=><option key={r}>{r}</option>)}</select></label>
          <label>Site<select value={siteId} onChange={e=>setSiteId(e.target.value)}>{siteOptions.map(a=><option key={a.id} value={a.site_id || ''}>{a.sites?.site_name || 'All Site'}</option>)}</select></label>
        </div>
        <button onClick={start}>Masuk Aplikasi</button>
      </>}
      <button className="secondary" onClick={()=>supabase.auth.signOut()}>Logout</button>
    </div>
  </section>
}

function Shell({ profile, context, setContext }){
  const [page, setPage] = useState('dashboard')
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const role = context.role
  const nav = useMemo(() => {
    const items = [['dashboard', 'Dashboard', LayoutDashboard]]
    if (canAdmin(role)) items.push(['admin', 'Admin Panel', Users])
    if (canAdmin(role)) items.push(['parameters', 'Parameter Checklist', ClipboardCheck])
    if (canAdmin(role) || role === 'GL' || role === 'Atasan Site') items.push(['units', 'Master Unit', Truck], ['parkings', 'Master Parkiran', ParkingCircle], ['plans', 'Plan Bulanan', CalendarCheck], ['inspections', 'Inspeksi', FileText])
    if (role === 'Viewer') items.push(['outstanding','Outstanding', AlertTriangle])
    if (canAdmin(role) || role === 'GL' || role === 'Atasan Site') items.push(['outstanding','Outstanding', AlertTriangle])
    if (canApprove(role)) items.push(['approval','Approval', CheckCircle2])
    return items
  }, [role])
  useEffect(() => { if (!nav.some(n => n[0] === page)) setPage('dashboard') }, [nav.map(n=>n[0]).join('|')])

  return <div className="app real-ui-shell">
    {sidebarOpen && <button className="mobile-nav-backdrop" aria-label="Tutup navigasi" onClick={()=>setSidebarOpen(false)} />}
    <aside className={sidebarOpen ? 'sidebar open' : 'sidebar'}>
      <button className="sidebar-close" onClick={()=>setSidebarOpen(false)}>×</button>
      <div className="brand side sidebar-brand"><img src={logoSrgs} alt="Platform Inspeksi SRGS" className="brand-logo-img"/><div><b>Inspeksi</b><span>{context.role} · {siteName(context)}</span></div></div>
      <div className="nav">
        {nav.map(([key, label, Icon]) => <button key={key} className={page === key ? 'active' : ''} onClick={()=>{setPage(key);setSidebarOpen(false)}}><Icon size={18}/><span>{label}</span></button>)}
      </div>
      <div className="sidebar-info-cards"><div className="sidebar-info-card"><b>{siteName(context)}</b><span>Lokasi kerja aktif</span></div><div className="sidebar-info-card"><b>Inspeksi</b><span>Unit, parkiran, PM check, approval, dan outstanding</span></div></div>
      <div className="sidebar-card"><b>Pusat Inspeksi</b><p>Kelola plan, pemeriksaan, temuan, dan approval sesuai role serta site kerja.</p></div>
      <div className="side-bottom">
        <button className="secondary" onClick={()=>setContext(null)}>Ganti Aplikasi</button>
      </div>
    </aside>
    <main className="main">
      <header className="topbar app-header">
        <button className="icon-btn mobile-menu-trigger" onClick={()=>setSidebarOpen(true)}><Menu size={20}/></button>
        <div><h1>{nav.find(n=>n[0]===page)?.[1] || 'Dashboard'}</h1><p>{profile.nama} · {profile.email}</p></div>
        <button className="secondary" onClick={hardLogout}><LogOut size={16}/> Logout</button>
      </header>
      <InspeksiPage page={page} profile={profile} context={context} />
    </main>
  </div>
}

function InspeksiPage({ page, profile, context }){
  if (page === 'dashboard') return <Dashboard context={context} />
  if (page === 'admin') return <SharedAdminPanel profile={profile} context={context} />
  if (page === 'units') return <MasterUnits profile={profile} context={context} />
  if (page === 'parkings') return <MasterParkings profile={profile} context={context} />
  if (page === 'parameters') return <ParameterChecklist context={context} />
  if (page === 'plans') return <PlanBulanan profile={profile} context={context} />
  if (page === 'inspections') return <InspectionExecution profile={profile} context={context} />
  if (page === 'outstanding') return <Outstanding profile={profile} context={context} />
  if (page === 'approval') return <Approval profile={profile} context={context} />
  return <Dashboard context={context} />
}



const DASHBOARD_V52_INLINE_CSS = `
/* V52 dashboard emergency layout guard. This keeps the dashboard neat even when styles.css was not copied during deploy. */
.dashboard-redesign{display:grid!important;gap:22px!important;width:100%!important;align-items:start!important}.dashboard-redesign *{box-sizing:border-box!important}.dashboard-hero-panel{position:relative!important;overflow:hidden!important;display:grid!important;grid-template-columns:1fr!important;gap:18px!important;align-items:stretch!important;background:linear-gradient(135deg,#071a3f,#123c98 55%,#2563eb)!important;border:1px solid rgba(255,255,255,.22)!important;border-radius:32px!important;padding:30px!important;box-shadow:0 28px 80px rgba(15,23,42,.18)!important;color:#fff!important}.dashboard-hero-panel:before{content:''!important;position:absolute!important;inset:-80px -80px auto auto!important;width:260px!important;height:260px!important;border-radius:999px!important;background:rgba(255,255,255,.14)!important}.dashboard-hero-panel:after{content:''!important;position:absolute!important;left:-120px!important;bottom:-140px!important;width:300px!important;height:300px!important;border-radius:999px!important;background:rgba(96,165,250,.18)!important}.hero-copy,.hero-score-card{position:relative!important;z-index:1!important}.hero-eyebrow{display:inline-flex!important;align-items:center!important;border:1px solid rgba(255,255,255,.22)!important;background:rgba(255,255,255,.12)!important;color:#dbeafe!important;border-radius:999px!important;padding:8px 12px!important;font-size:12px!important;font-weight:950!important;text-transform:uppercase!important;letter-spacing:.1em!important}.hero-copy h2{margin:14px 0 12px!important;font-size:36px!important;line-height:1.08!important;letter-spacing:-.045em!important;color:#fff!important}.hero-copy p{margin:0!important;max-width:820px!important;color:#dbeafe!important;font-size:16px!important;line-height:1.6!important}.hero-score-card{background:rgba(255,255,255,.14)!important;border:1px solid rgba(255,255,255,.18)!important;border-radius:26px!important;padding:22px!important;box-shadow:inset 0 1px 0 rgba(255,255,255,.12)!important;backdrop-filter:blur(8px)!important}.hero-score-card small{color:#dbeafe!important;font-weight:900!important}.hero-score-card strong{display:block!important;font-size:56px!important;line-height:1!important;margin:10px 0!important;color:#fff!important;letter-spacing:-.05em!important}.hero-score-card p{color:#dbeafe!important;margin:12px 0 0!important}.hero-bar{height:15px!important;background:rgba(255,255,255,.24)!important}.hero-bar span{background:#fff!important}.executive-kpi-grid{display:grid!important;grid-template-columns:repeat(4,minmax(0,1fr))!important;gap:18px!important}.executive-kpi-grid .kpi{background:linear-gradient(180deg,#fff,#f8fbff)!important;border-color:#e4edfa!important}.dashboard-category-grid{display:grid!important;grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:20px!important}.category-dashboard-card{position:relative!important;overflow:hidden!important;background:#fff!important;border:1px solid #e4edfa!important;border-radius:30px!important;padding:24px!important;box-shadow:0 18px 48px rgba(15,23,42,.07)!important;min-width:0!important}.category-dashboard-card:before{content:''!important;position:absolute!important;right:-42px!important;top:-42px!important;width:150px!important;height:150px!important;border-radius:999px!important;background:#eff6ff!important}.category-dashboard-card.pm:before{background:#dbeafe!important}.category-dashboard-card.unit:before{background:#dcfce7!important}.category-dashboard-card.parking:before{background:#ede9fe!important}.category-card-head{position:relative!important;z-index:1!important;display:flex!important;align-items:center!important;gap:14px!important;margin-bottom:12px!important}.category-card-head small{display:block!important;text-transform:uppercase!important;letter-spacing:.08em!important;font-size:11px!important;font-weight:950!important;color:#64748b!important}.category-card-head h3{margin:2px 0 0!important;font-size:24px!important;letter-spacing:-.035em!important}.category-card-head .pill,.category-card-head .badge{margin-left:auto!important;white-space:nowrap!important}.category-icon{width:54px!important;height:54px!important;min-width:54px!important;border-radius:18px!important;display:grid!important;place-items:center!important;background:#eff6ff!important;color:#2563eb!important;box-shadow:0 12px 32px rgba(37,99,235,.14)!important}.category-icon svg{width:24px!important;height:24px!important}.category-dashboard-card.unit .category-icon,.category-icon.unit{background:#ecfdf5!important;color:#16a34a!important}.category-dashboard-card.parking .category-icon,.category-icon.parking{background:#f5f3ff!important;color:#7c3aed!important}.category-dashboard-card.pm .category-icon,.category-icon.pm{background:#eff6ff!important;color:#2563eb!important}.category-dashboard-card p{position:relative!important;z-index:1!important;color:#64748b!important;margin:0 0 18px!important;min-height:50px!important;line-height:1.55!important}.category-main-metric{position:relative!important;z-index:1!important;background:#f8fbff!important;border:1px solid #e5eefb!important;border-radius:22px!important;padding:16px!important;margin-bottom:16px!important}.category-main-metric small,.mini-metric small{display:block!important;color:#64748b!important;font-weight:900!important}.category-main-metric strong{display:block!important;font-size:36px!important;letter-spacing:-.04em!important;margin:4px 0 10px!important;color:#061027!important}.category-metric-grid{position:relative!important;z-index:1!important;display:grid!important;grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:10px!important}.mini-metric{background:#fff!important;border:1px solid #e8eff9!important;border-radius:17px!important;padding:12px!important;min-width:0!important}.mini-metric b{display:block!important;font-size:22px!important;letter-spacing:-.03em!important;margin-top:4px!important;color:#061027!important}.category-comparison-grid{display:grid!important;grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:16px!important;margin-bottom:18px!important}.category-progress-card{background:linear-gradient(180deg,#fff,#f8fbff)!important;border:1px solid #e4edfa!important;border-radius:24px!important;padding:16px!important;box-shadow:0 12px 30px rgba(15,23,42,.05)!important}.category-progress-top{display:grid!important;grid-template-columns:54px 1fr auto!important;gap:12px!important;align-items:center!important;margin-bottom:14px!important}.category-progress-top b{display:block!important;color:#061027!important}.category-progress-top small{display:block!important;margin-top:3px!important;color:#64748b!important}.category-progress-top strong{font-size:26px!important;letter-spacing:-.04em!important;color:#061027!important}.progress-foot{display:flex!important;gap:8px!important;flex-wrap:wrap!important;margin-top:12px!important}.progress-foot span{background:#eef4ff!important;border:1px solid #dbeafe!important;border-radius:999px!important;padding:7px 10px!important;font-size:12px!important;font-weight:900!important;color:#334155!important}.site-category-board{display:grid!important;grid-template-columns:repeat(auto-fit,minmax(320px,1fr))!important;gap:14px!important;margin-bottom:18px!important}.site-category-card{background:#fff!important;border:1px solid #e4edfa!important;border-radius:24px!important;padding:16px!important;box-shadow:0 12px 30px rgba(15,23,42,.04)!important}.site-category-card>div:first-child{display:flex!important;justify-content:space-between!important;gap:10px!important;align-items:flex-start!important;margin-bottom:12px!important}.site-category-card b{font-size:16px!important;color:#061027!important}.site-category-card small{max-width:190px!important;text-align:right!important;color:#64748b!important}.mini-progress-list{display:grid!important;gap:10px!important}.mini-progress-row{display:grid!important;grid-template-columns:62px minmax(0,1fr) 100px!important;gap:10px!important;align-items:center!important}.mini-progress-row span{font-size:12px!important;font-weight:950!important;color:#475569!important}.mini-progress-row .bar{height:10px!important}.mini-progress-row b{font-size:12px!important;text-align:right!important;color:#0f172a!important}.dashboard-redesign .grid-2{align-items:start!important}.dashboard-redesign .grid-2 .panel{min-width:0!important}.dashboard-redesign .bar{height:13px!important;background:#e2e8f0!important;border-radius:999px!important;overflow:hidden!important}.dashboard-redesign .bar span{height:100%!important;display:block!important;background:linear-gradient(90deg,#2563eb,#7c3aed)!important;border-radius:999px!important}.dashboard-redesign .summary-strip{display:flex!important;gap:12px!important;flex-wrap:wrap!important;margin:10px 0 18px!important}.dashboard-redesign .summary-strip span{background:#eef4ff!important;border:1px solid #dbeafe!important;border-radius:16px!important;padding:12px 16px!important;color:#0f172a!important;font-weight:800!important}.dashboard-redesign .site-chart{display:grid!important;gap:14px!important;margin-bottom:18px!important}.dashboard-redesign .site-bar{display:grid!important;grid-template-columns:200px 1fr!important;gap:14px!important;align-items:center!important;background:#f8fbff!important;border:1px solid #dbeafe!important;border-radius:18px!important;padding:14px!important}.dashboard-redesign .site-meta b{display:block!important;color:#061027!important}.dashboard-redesign .site-meta span{display:block!important;color:#64748b!important;font-size:12px!important;margin-top:4px!important}@media(max-width:1180px){.dashboard-hero-panel,.dashboard-category-grid,.category-comparison-grid{grid-template-columns:1fr!important}.executive-kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr))!important}.hero-copy h2{font-size:30px!important}.hero-score-card strong{font-size:44px!important}.site-category-board{grid-template-columns:1fr!important}}@media(max-width:640px){.dashboard-hero-panel{padding:20px!important;border-radius:24px!important}.executive-kpi-grid{grid-template-columns:1fr!important}.category-dashboard-card{padding:18px!important;border-radius:24px!important}.category-metric-grid,.mini-progress-row{grid-template-columns:1fr!important}.mini-progress-row b{text-align:left!important}.site-category-card small{text-align:left!important}.category-progress-top{grid-template-columns:44px 1fr!important}.category-progress-top strong{grid-column:1/-1!important}.category-icon{width:44px!important;height:44px!important;min-width:44px!important;border-radius:15px!important}.dashboard-redesign .site-bar{grid-template-columns:1fr!important}}
`

function DashboardV52Styles(){
  return <style data-dashboard-v52>{DASHBOARD_V52_INLINE_CSS}</style>
}

function DashboardDateFilter({ dateFrom, dateTo, setDateFrom, setDateTo, siteFilter, setSiteFilter, sites, canChooseSite, onClear }) {
  return <Panel title="Filter Dashboard" desc="Pilih tanggal dan site sesuai kebutuhan.">
    <div className="form-grid">
      <label>Dari Tanggal<input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} /></label>
      <label>Sampai Tanggal<input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} /></label>
      <label>Site
        <select value={siteFilter} onChange={e=>setSiteFilter(e.target.value)} disabled={!canChooseSite}>
          {canChooseSite && <option value="">Semua Site</option>}
          {sites.map(s => <option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}
        </select>
      </label>
      <button type="button" className="secondary" onClick={onClear}>Reset Filter</button>
    </div>
  </Panel>
}

function Dashboard({ context }){
  const todayStr = new Date().toISOString().slice(0, 10)
  const [dateFrom, setDateFrom] = useState(todayStr)
  const [dateTo, setDateTo] = useState(todayStr)
  const [loading, setLoading] = useState(true)
  const [plans, setPlans] = useState([])
  const [records, setRecords] = useState([])
  const [findings, setFindings] = useState([])
  const [sites, setSites] = useState([])
  const [units, setUnits] = useState([])
  const [siteFilter, setSiteFilter] = useState(isHeadOfficeAdmin(context) ? '' : (context.site_id || ''))
  const [error, setError] = useState('')

  useEffect(() => { setSiteFilter(isHeadOfficeAdmin(context) ? '' : (context.site_id || '')) }, [context.id, context.site_id])
  useEffect(() => { load() }, [context.id, dateFrom, dateTo])
  async function load(){
    setLoading(true); setError('')
    try{
      const [allPlans, allRecords, allFindings, allSites, allUnits] = await Promise.all([
        fetchAllRows('inspection_plans', '*, sites(site_code,site_name), inspection_units(unit_code,unit_name), inspection_parkings(parking_code,parking_name)', q => {
          q = scopeQuery(q.order('created_at', { ascending:false }), context)
          if (dateFrom) q = q.gte('created_at', dateFrom)
          if (dateTo) q = q.lte('created_at', dateTo + 'T23:59:59')
          return q
        }),
        fetchAllRows('inspection_records', '*, sites(site_code,site_name), inspection_plans(*, inspection_units(unit_code,unit_name), inspection_parkings(parking_code,parking_name))', q => {
          q = scopeQuery(q.order('created_at', { ascending:false }), context)
          if (dateFrom) q = q.gte('inspected_at', dateFrom)
          if (dateTo) q = q.lte('inspected_at', dateTo + 'T23:59:59')
          return q
        }),
        fetchAllRows('inspection_findings', '*, sites(site_code,site_name), inspection_parameters(parameter_code,parameter_name)', q => {
          q = scopeQuery(q.order('created_at', { ascending:false }), context)
          if (dateFrom) q = q.gte('created_at', dateFrom)
          if (dateTo) q = q.lte('created_at', dateTo + 'T23:59:59')
          return q
        }),
        fetchAllRows('sites', '*', q => q.neq('site_code', 'JIEP').order('site_code')),
        fetchAllRows('inspection_units', 'id,site_id,unit_code,unit_name,status', q => scopeQuery(q.order('unit_code'), context))
      ])
      const activeSites = (allSites || []).filter(x => isActiveMasterStatus(x.status))
      setPlans(allPlans || [])
      setRecords(allRecords || [])
      setFindings(allFindings || [])
      setSites(isHeadOfficeAdmin(context) ? activeSites : activeSites.filter(x => x.id === context.site_id))
      setUnits((allUnits || []).filter(x => isActiveMasterStatus(x.status)))
    }catch(e){
      console.error('[DASHBOARD LOAD ERROR]', e)
      setError(e.message || String(e))
      setPlans([]); setRecords([]); setFindings([]); setSites([]); setUnits([])
    }finally{
      setLoading(false)
    }
  }

  const calcPercent = (done, target) => target ? Math.round((done / target) * 100) : 0
  const canChooseSite = isHeadOfficeAdmin(context)
  const effectiveSiteId = canChooseSite ? siteFilter : (context.site_id || '')
  const dashboardSites = [...(effectiveSiteId ? sites.filter(s => s.id === effectiveSiteId) : sites)]
    .sort((a,b)=>String(a.site_code || a.site_name || '').localeCompare(String(b.site_code || b.site_name || ''), 'id', { numeric:true, sensitivity:'base' }))
  const dashboardPlans = effectiveSiteId ? plans.filter(p => p.site_id === effectiveSiteId) : plans
  const dashboardRecords = effectiveSiteId ? records.filter(r => r.site_id === effectiveSiteId) : records
  const dashboardFindings = effectiveSiteId ? findings.filter(f => f.site_id === effectiveSiteId) : findings
  const dashboardUnits = effectiveSiteId ? units.filter(u => u.site_id === effectiveSiteId) : units

  const waitingApproval = dashboardRecords.filter(r => r.status === 'Submitted').length
  const openFindings = dashboardFindings.filter(f => f.status === 'Open').length
  const closeRequested = dashboardFindings.filter(f => f.status === 'Close Requested').length
  const closedFindings = dashboardFindings.filter(f => f.status === 'Closed').length
  const approvedPlans = dashboardPlans.filter(p => p.status === 'Approved').length
  const rejectedPlans = dashboardPlans.filter(p => p.status === 'Rejected').length

  const categoryDefs = [
    {
      key:'PM Check',
      short:'PM Check',
      title:'PM Check',
      tone:'pm',
      icon:<Truck size={24}/> ,
      desc:'Target 2 kali PM Check per unit aktif.',
      targetLabel:'Target PM/Bulan',
      targetValue:dashboardUnits.length * 2
    },
    {
      key:'Inspeksi Unit',
      short:'Unit',
      title:'Inspeksi Unit',
      tone:'unit',
      icon:<Building2 size={24}/> ,
      desc:'Plan, hasil inspeksi, approval, dan temuan unit.',
      targetLabel:'Total Plan Unit',
      targetValue:null
    },
    {
      key:'Inspeksi Kelayakan Parkiran',
      short:'Parkiran',
      title:'Inspeksi Parkiran',
      tone:'parking',
      icon:<ParkingCircle size={24}/> ,
      desc:'Plan, hasil inspeksi, approval, dan temuan parkiran.',
      targetLabel:'Total Plan Parkiran',
      targetValue:null
    }
  ]

  const categoryStats = categoryDefs.map(def => {
    const catPlans = dashboardPlans.filter(p => p.category === def.key)
    const catRecords = dashboardRecords.filter(r => r.category === def.key)
    const catFindings = dashboardFindings.filter(f => f.category === def.key)
    const planApproved = catPlans.filter(p => p.status === 'Approved').length
    const planSubmitted = catPlans.filter(p => p.status === 'Submitted').length
    const recordSubmitted = catRecords.filter(r => r.status === 'Submitted').length
    const target = def.targetValue === null ? catPlans.length : Math.max(def.targetValue, catPlans.length)
    const achievement = calcPercent(planApproved, target)
    return {
      ...def,
      total_plan:catPlans.length,
      approved:planApproved,
      submitted:planSubmitted + recordSubmitted,
      rejected:catPlans.filter(p => p.status === 'Rejected').length,
      target,
      achievement,
      aman:catRecords.filter(r => r.result === 'Aman').length,
      tidak_aman:catRecords.filter(r => r.result === 'Tidak Aman').length,
      temuan_open:catFindings.filter(f => f.status === 'Open').length,
      close_request:catFindings.filter(f => f.status === 'Close Requested').length,
      temuan_closed:catFindings.filter(f => f.status === 'Closed').length
    }
  })

  const pmPlans = dashboardPlans.filter(p => p.category === 'PM Check')
  const pmApproved = pmPlans.filter(p => p.status === 'Approved').length
  const pmUnitTarget = dashboardUnits.length * 2
  const pmAchievement = calcPercent(pmApproved, pmUnitTarget)
  const pmRows = dashboardSites.map(s => {
    const siteUnits = dashboardUnits.filter(u => u.site_id === s.id).length
    const target = siteUnits * 2
    const actual = pmPlans.filter(p => p.site_id === s.id && p.status === 'Approved').length
    return { site:s.site_code, site_name:s.site_name, unit_aktif:siteUnits, target_pm_bulanan:target, pm_approved:actual, belum_pm:Math.max(target-actual,0), achievement:`${calcPercent(actual,target)}%` }
  })

  const categoryRows = categoryStats.map(c => ({
    kategori:c.title,
    target:c.target,
    total_plan:c.total_plan,
    approved:c.approved,
    menunggu_approval:c.submitted,
    rejected:c.rejected,
    aman:c.aman,
    tidak_aman:c.tidak_aman,
    temuan_open:c.temuan_open,
    close_request:c.close_request,
    temuan_closed:c.temuan_closed,
    achievement:`${c.achievement}%`
  }))

  const makeDetailRows = (categoryKey) => dashboardSites.map(s => {
    const catPlans = dashboardPlans.filter(p => p.site_id === s.id && p.category === categoryKey)
    const catRecords = dashboardRecords.filter(r => r.site_id === s.id && r.category === categoryKey)
    const catFindings = dashboardFindings.filter(f => f.site_id === s.id && f.category === categoryKey)
    const approved = catPlans.filter(p => p.status === 'Approved').length
    const waiting = catRecords.filter(r => r.status === 'Submitted').length + catPlans.filter(p => p.status === 'Submitted').length
    const total = catPlans.length
    return {
      site:s.site_code,
      site_name:s.site_name,
      total_plan:total,
      approved,
      menunggu_approval:waiting,
      rejected:catPlans.filter(p => p.status === 'Rejected').length,
      aman:catRecords.filter(r => r.result === 'Aman').length,
      tidak_aman:catRecords.filter(r => r.result === 'Tidak Aman').length,
      temuan_open:catFindings.filter(f => f.status === 'Open').length,
      close_request:catFindings.filter(f => f.status === 'Close Requested').length,
      temuan_closed:catFindings.filter(f => f.status === 'Closed').length,
      achievement:`${calcPercent(approved,total)}%`
    }
  })

  const unitRows = makeDetailRows('Inspeksi Unit')
  const parkingRows = makeDetailRows('Inspeksi Kelayakan Parkiran')
  const unitStats = categoryStats.find(c => c.key === 'Inspeksi Unit') || {}
  const parkingStats = categoryStats.find(c => c.key === 'Inspeksi Kelayakan Parkiran') || {}

  const inspectionRows = dashboardRecords.map(r => ({
    tanggal: r.inspected_at ? new Date(r.inspected_at).toLocaleString('id-ID') : '-',
    site: r.sites?.site_code,
    kategori: r.category,
    target: targetLabel(r.inspection_plans),
    hasil: r.result,
    status: r.status,
    catatan: r.notes || '-'
  }))
  const findingRows = dashboardFindings.map(f => ({
    created: f.created_at ? new Date(f.created_at).toLocaleDateString('id-ID') : '-',
    site: f.sites?.site_code,
    kategori: f.category,
    target: f.target_label,
    parameter: f.inspection_parameters?.parameter_name,
    deskripsi: f.finding_description,
    status: f.status,
    due_date: f.due_date || '-'
  }))

  if (loading) return <Panel title="Dashboard"><p className="muted">Memuat dashboard...</p></Panel>
  if (error) return <Panel title="Dashboard"><div className="error">{error}</div><button onClick={load}>Coba Lagi</button></Panel>
  return <div className="stack dashboard-redesign">
    <DashboardV52Styles />
    <DashboardDateFilter
      dateFrom={dateFrom}
      dateTo={dateTo}
      setDateFrom={setDateFrom}
      setDateTo={setDateTo}
      siteFilter={effectiveSiteId}
      setSiteFilter={setSiteFilter}
      sites={sites}
      canChooseSite={canChooseSite}
      onClear={()=>{ setDateFrom(''); setDateTo(''); if (canChooseSite) setSiteFilter('') }}
    />

    <section className="dashboard-hero-panel">
      <div className="hero-copy">
        <span className="hero-eyebrow">Dashboard Inspeksi</span>
        <h2>Dashboard Inspeksi Unit, Inspeksi Parkiran, dan PM Check</h2>
        <p>Ringkasan plan, approval, achievement, dan temuan.</p>
      </div>
    </section>

    <div className="kpi-grid executive-kpi-grid">
      <Kpi title="Total Plan" value={dashboardPlans.length} icon={<CalendarCheck/>} />
      <Kpi title="Plan Approved" value={approvedPlans} icon={<CheckCircle2/>} />
      <Kpi title="Menunggu Approval" value={waitingApproval} icon={<Eye/>} />
      <Kpi title="Temuan Open" value={openFindings} icon={<AlertTriangle/>} />
      <Kpi title="Close Request" value={closeRequested} icon={<ClipboardCheck/>} />
      <Kpi title="Temuan Closed" value={closedFindings} icon={<Check/>} />
      <Kpi title="Plan Rejected" value={rejectedPlans} icon={<XCircle/>} />
      <Kpi title="Target PM Check" value={pmUnitTarget} icon={<Truck/>} />
    </div>

    <div className="dashboard-category-grid">
      {categoryStats.map(item => <CategoryDashboardCard key={item.key} item={item} />)}
    </div>

    <Panel title="Achievement per Kategori" desc="Ringkasan utama untuk PM Check, Inspeksi Unit, dan Inspeksi Parkiran." action={<button onClick={()=>downloadXlsx('resume-dashboard-inspeksi-per-kategori.xlsx', categoryRows)}><Download size={16}/> Export Excel</button>}>
      <div className="category-comparison-grid">
        {categoryStats.map(item => <CategoryProgress key={item.key} item={item} />)}
      </div>
      <DataTable rows={categoryRows}/>
    </Panel>


    <Panel title="Dashboard PM Check" desc="Pencapaian PM Check dihitung dari target 2 kali PM Check per unit aktif setiap bulan." action={<button onClick={()=>downloadXlsx('dashboard-pm-check.xlsx', pmRows)}><Download size={16}/> Export Excel</button>}>
      <div className="summary-strip"><span><b>{dashboardUnits.length}</b> Unit Aktif</span><span><b>{pmUnitTarget}</b> Target PM/Bulan</span><span><b>{pmApproved}</b> PM Approved</span><span><b>{pmAchievement}%</b> Achievement PM</span></div>
      <div className="site-chart">
        {pmRows.map(r => <div className="site-bar" key={r.site}>
          <div className="site-meta"><b>{r.site}</b><span>{r.pm_approved}/{r.target_pm_bulanan} · {r.achievement}</span></div>
          <div className="bar"><span style={{ width: `${Math.min(parseInt(r.achievement) || 0, 100)}%` }} /></div>
        </div>)}
        {!pmRows.length && <p className="muted">Belum ada data PM Check.</p>}
      </div>
      <DataTable rows={pmRows}/>
    </Panel>

    <div className="grid-2">
      <Panel title="Dashboard Inspeksi Unit" desc="Detail pencapaian inspeksi unit per site." action={<button onClick={()=>downloadXlsx('dashboard-inspeksi-unit.xlsx', unitRows)}><Download size={16}/> Export Excel</button>}>
        <div className="summary-strip"><span><b>{unitStats.total_plan || 0}</b> Total Plan</span><span><b>{unitStats.approved || 0}</b> Approved</span><span><b>{unitStats.achievement || 0}%</b> Achievement</span><span><b>{unitStats.temuan_open || 0}</b> Temuan Open</span></div>
        <div className="site-chart">
          {unitRows.map(r => <div className="site-bar" key={r.site}>
            <div className="site-meta"><b>{r.site}</b><span>{r.approved}/{r.total_plan} · {r.achievement}</span></div>
            <div className="bar"><span style={{ width: `${Math.min(parseInt(r.achievement) || 0, 100)}%` }} /></div>
          </div>)}
          {!unitRows.length && <p className="muted">Belum ada data inspeksi unit.</p>}
        </div>
        <DataTable rows={unitRows}/>
      </Panel>

      <Panel title="Dashboard Inspeksi Parkiran" desc="Detail pencapaian inspeksi parkiran per site." action={<button onClick={()=>downloadXlsx('dashboard-inspeksi-parkiran.xlsx', parkingRows)}><Download size={16}/> Export Excel</button>}>
        <div className="summary-strip"><span><b>{parkingStats.total_plan || 0}</b> Total Plan</span><span><b>{parkingStats.approved || 0}</b> Approved</span><span><b>{parkingStats.achievement || 0}%</b> Achievement</span><span><b>{parkingStats.temuan_open || 0}</b> Temuan Open</span></div>
        <div className="site-chart">
          {parkingRows.map(r => <div className="site-bar" key={r.site}>
            <div className="site-meta"><b>{r.site}</b><span>{r.approved}/{r.total_plan} · {r.achievement}</span></div>
            <div className="bar"><span style={{ width: `${Math.min(parseInt(r.achievement) || 0, 100)}%` }} /></div>
          </div>)}
          {!parkingRows.length && <p className="muted">Belum ada data inspeksi parkiran.</p>}
        </div>
        <DataTable rows={parkingRows}/>
      </Panel>
    </div>

    <div className="grid-2">
      <Panel title="Row Data Inspeksi" desc="Data mentah hasil inspeksi untuk audit dan export Excel." action={<button onClick={()=>downloadXlsx('row-data-inspeksi.xlsx', inspectionRows)}><Download size={16}/> Export Excel</button>}>
        <DataTable rows={inspectionRows}/>
      </Panel>
      <Panel title="Row Data Temuan" desc="Seluruh temuan open, close requested, dan closed untuk monitoring tindak lanjut." action={<button onClick={()=>downloadXlsx('row-data-temuan-inspeksi.xlsx', findingRows)}><Download size={16}/> Export Excel</button>}>
        <div className="summary-strip"><span><b>{openFindings}</b> Open</span><span><b>{closeRequested}</b> Close Requested</span><span><b>{closedFindings}</b> Closed</span></div>
        <DataTable rows={findingRows}/>
      </Panel>
    </div>
  </div>
}

function CategoryDashboardCard({ item }){
  return <article className={`category-dashboard-card ${item.tone}`}>
    <div className="category-card-head">
      <div className="category-icon">{item.icon}</div>
      <div>
        <small>{item.short}</small>
        <h3>{item.title}</h3>
      </div>
      <StatusPill value={`${item.achievement}%`} />
    </div>
    <p>{item.desc}</p>
    <div className="category-main-metric">
      <div><small>Achievement</small><strong>{item.achievement}%</strong></div>
      <div className="bar"><span style={{ width: `${Math.min(item.achievement, 100)}%` }} /></div>
    </div>
    <div className="category-metric-grid">
      <MiniMetric label={item.targetLabel} value={item.target} />
      <MiniMetric label="Total Plan" value={item.total_plan} />
      <MiniMetric label="Approved" value={item.approved} />
      <MiniMetric label="Menunggu" value={item.submitted} />
      <MiniMetric label="Aman" value={item.aman} />
      <MiniMetric label="Tidak Aman" value={item.tidak_aman} />
      <MiniMetric label="Temuan Open" value={item.temuan_open} />
      <MiniMetric label="Closed" value={item.temuan_closed} />
    </div>
  </article>
}

function CategoryProgress({ item }){
  return <div className="category-progress-card">
    <div className="category-progress-top">
      <div className={`category-icon ${item.tone}`}>{item.icon}</div>
      <div><b>{item.title}</b><small>{item.approved}/{item.target} approved</small></div>
      <strong>{item.achievement}%</strong>
    </div>
    <div className="bar"><span style={{ width: `${Math.min(item.achievement, 100)}%` }} /></div>
    <div className="progress-foot"><span>{item.total_plan} plan</span><span>{item.temuan_open} open finding</span><span>{item.close_request} close request</span></div>
  </div>
}

function MiniMetric({ label, value }){
  return <div className="mini-metric"><small>{label}</small><b>{value}</b></div>
}

function MiniProgress({ label, value, meta }){
  return <div className="mini-progress-row">
    <span>{label}</span>
    <div className="bar"><span style={{ width: `${Math.min(value, 100)}%` }} /></div>
    <b>{meta} · {value}%</b>
  </div>
}

function AdminPanel({ profile, context }){
  return <div className="stack">
    <AccessMapping profile={profile} context={context} />
    <SiteMaster />
  </div>
}

function AccessMapping(){
  const [apps, setApps] = useState([])
  const [sites, setSites] = useState([])
  const [profiles, setProfiles] = useState([])
  const [accessRows, setAccessRows] = useState([])
  const [message, setMessage] = useState('')
  const [loading, setLoading] = useState(false)
  const [form, setForm] = useState({ user_id:'', nama:'', nrp:'', email:'', password:'', app_id:'', role:'GL', site_id:'' })

  useEffect(() => { load() }, [])
  async function load(){
    const [a, s, p, ua] = await Promise.all([
      fetchAllRows('applications', '*', q => q.order('app_name')),
      fetchAllRows('sites', '*', q => q.order('site_code')),
      fetchAllRows('users_profile', '*', q => q.order('created_at', { ascending:false })),
      fetchAllRows('user_app_access', '*, users_profile(nama,email,nrp), applications(app_code,app_name), sites(site_code,site_name)', q => q.order('created_at', { ascending:false }))
    ])
    setApps(a || [])
    setSites(s || [])
    setProfiles(p || [])
    setAccessRows(ua || [])
    const inspeksi = (a || []).find(x => x.app_code === APP_CODE) || (a || [])[0]
    const jiep = (s || []).find(x => x.site_code === 'JIEP') || (s || [])[0]
    setForm(f => ({ ...f, app_id: f.app_id || inspeksi?.id || '', site_id: f.site_id || jiep?.id || '' }))
  }
  function selectExisting(id){
    const p = profiles.find(x => x.id === id)
    setForm(f => ({ ...f, user_id:id, nama:p?.nama || '', nrp:p?.nrp || '', email:p?.email || '', password:'' }))
  }
  async function saveAccess(e){
    e.preventDefault()
    setLoading(true); setMessage('')
    try{
      let userId = form.user_id
      const siteId = form.site_id || null
      if (userId){
        const { error } = await supabase.from('user_app_access').insert({ user_id:userId, app_id:form.app_id, role:form.role, site_id:siteId, status:'Aktif' })
        if (error) throw error
      } else {
        const { data, error } = await supabase.functions.invoke('admin-create-user', { body: {
          nama: clean(form.nama) || normEmail(form.email), nrp: clean(form.nrp), email: normEmail(form.email), password: form.password,
          app_id: form.app_id, role: form.role, site_id: siteId
        }})
        if (error) throw error
        if (data?.error) throw new Error(data.error)
      }
      setMessage('User/profile/mapping akses berhasil disimpan.')
      setForm(f => ({ ...f, user_id:'', nama:'', nrp:'', email:'', password:'' }))
      await load()
    }catch(err){ setMessage(err.message || String(err)) }
    finally{ setLoading(false) }
  }
  async function toggle(row){
    const next = row.status === 'Aktif' ? 'Nonaktif' : 'Aktif'
    const { error } = await supabase.from('user_app_access').update({ status: next }).eq('id', row.id)
    setMessage(error ? error.message : `Akses ${next}.`)
    load()
  }
  const rows = accessRows.map(r => ({ user:r.users_profile?.nama, email:r.users_profile?.email, app:r.applications?.app_name, role:r.role, site:r.sites?.site_code || 'All Site', status:r.status, aksi:'' }))
  return <Panel title="Mapping Akses User" desc="Admin bisa membuat user lengkap dengan password dan mapping aplikasi-role-site. Untuk akun existing, pilih user lalu tambahkan mapping baru." action={<button className="secondary" onClick={()=>downloadTemplate('access')}><Download size={16}/> Template</button>}>
    <form className="form-grid" onSubmit={saveAccess}>
      <label>User Existing<select value={form.user_id} onChange={e=>selectExisting(e.target.value)}><option value="">Buat user baru + password</option>{profiles.map(p=><option key={p.id} value={p.id}>{p.nama || p.email} · {p.email}</option>)}</select></label>
      <label>Nama<input disabled={!!form.user_id} value={form.nama} onChange={e=>setForm({...form,nama:e.target.value})}/></label>
      <label>NRP<input disabled={!!form.user_id} value={form.nrp} onChange={e=>setForm({...form,nrp:e.target.value})}/></label>
      <label>Email<input disabled={!!form.user_id} type="email" value={form.email} onChange={e=>setForm({...form,email:e.target.value})}/></label>
      <label>Password<input disabled={!!form.user_id} type="password" placeholder="Minimal 6 karakter" value={form.password} onChange={e=>setForm({...form,password:e.target.value})}/></label>
      <label>Aplikasi<select value={form.app_id} onChange={e=>setForm({...form,app_id:e.target.value})}>{apps.map(a=><option key={a.id} value={a.id}>{a.app_name}</option>)}</select></label>
      <label>Role<select value={form.role} onChange={e=>setForm({...form,role:e.target.value})}>{['Platform Admin','App Admin','GL','Atasan Site','Viewer'].map(r=><option key={r}>{r}</option>)}</select></label>
      <label>Site<select value={form.site_id} onChange={e=>setForm({...form,site_id:e.target.value})}><option value="">All Site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>
      <button disabled={loading}>{loading ? 'Memproses...' : 'Simpan User / Mapping'}</button>
    </form>
    {message && <p className="message">{message}</p>}
    <p className="muted">Catatan: pembuatan password dari aplikasi memakai Supabase Edge Function <code>admin-create-user</code>. Role GL dan Atasan Site sebaiknya dipilih per site. Admin JIEP bisa melihat all site.</p>
    <Panel title="Row Data Mapping Akses" desc="Akses bisa dinonaktifkan tanpa menghapus histori." action={<button onClick={()=>downloadXlsx('mapping-akses-inspeksi.xlsx', rows.map(({aksi,...r})=>r))}><Download size={16}/> Export Excel</button>}>
      <DataTable rows={rows} customActions={(idx)=>{ const r=accessRows[idx]; return <button className="secondary small" onClick={()=>toggle(r)}>{r.status === 'Aktif' ? 'Nonaktifkan' : 'Aktifkan'}</button> }} />
    </Panel>
  </Panel>
}

function SiteMaster(){
  const [sites, setSites] = useState([])
  const [editing, setEditing] = useState(null)
  const [message, setMessage] = useState('')
  const [form, setForm] = useState({ site_code:'', site_name:'', region:'Operation', status:'Aktif' })
  useEffect(() => { load() }, [])
  async function load(){ const data = await fetchAllRows('sites', '*', q => q.order('site_code')); setSites(data || []) }
  function edit(s){ setEditing(s.id); setForm({ site_code:s.site_code, site_name:s.site_name, region:s.region || '', status:s.status || 'Aktif' }) }
  function reset(){ setEditing(null); setForm({ site_code:'', site_name:'', region:'Operation', status:'Aktif' }) }
  async function save(e){
    e.preventDefault(); setMessage('')
    const payload = { ...form, site_code: clean(form.site_code).toUpperCase(), site_name: clean(form.site_name) }
    const res = editing ? await supabase.from('sites').update(payload).eq('id', editing) : await supabase.from('sites').insert(payload)
    setMessage(res.error ? res.error.message : editing ? 'Site berhasil diupdate.' : 'Site berhasil ditambahkan.')
    if (!res.error){ reset(); load() }
  }
  async function remove(s){
    if (!confirm(`Hapus site ${s.site_code}? Jika sudah dipakai transaksi, Supabase akan menolak.`)) return
    const { error } = await supabase.from('sites').delete().eq('id', s.id)
    setMessage(error ? error.message : 'Site berhasil dihapus.')
    load()
  }
  const rows = sites.map(s => ({ code:s.site_code, nama:s.site_name, region:s.region, status:s.status, aksi:'' }))
  return <Panel title="Master Site" desc="JIEP adalah Head Office dan tidak dihitung dalam dashboard achievement all site.">
    <form className="form-grid" onSubmit={save}>
      <label>Site Code<input required value={form.site_code} onChange={e=>setForm({...form,site_code:e.target.value})}/></label>
      <label>Site Name<input required value={form.site_name} onChange={e=>setForm({...form,site_name:e.target.value})}/></label>
      <label>Region<input value={form.region} onChange={e=>setForm({...form,region:e.target.value})}/></label>
      <label>Status<select value={form.status} onChange={e=>setForm({...form,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label>
      <button>{editing ? 'Update Site' : 'Simpan Site'}</button>{editing && <button type="button" className="secondary" onClick={reset}>Batal</button>}
    </form>
    {message && <p className="message">{message}</p>}
    <DataTable rows={rows} customActions={(idx)=>{ const s=sites[idx]; return <div className="row-actions"><button className="secondary small" onClick={()=>edit(s)}>Edit</button><button className="danger small" onClick={()=>remove(s)}>Delete</button></div> }} />
  </Panel>
}

function MasterUnits({ profile, context }){
  const [rows, setRows] = useState([])
  const [sites, setSites] = useState([])
  const [editing, setEditing] = useState(null)
  const [message, setMessage] = useState('')
  const [preview, setPreview] = useState([])
  const adminHO = isHeadOfficeAdmin(context)
  const [form, setForm] = useState({ site_id: context.site_id || '', unit_code:AUTO_CODE_LABEL, unit_name:'', unit_type:'', location:'', status:'Aktif' })
  useEffect(() => { load() }, [context.id])
  async function load(){
    try {
      const [allUnits, allSites] = await Promise.all([
        fetchAllRows('inspection_units', '*, sites(site_code,site_name)', q => q.order('created_at', { ascending:false })),
        fetchAllRows('sites', '*', q => q.neq('site_code','JIEP').eq('status','Aktif').order('site_code'))
      ])
      const scopedUnits = adminHO ? (allUnits || []) : (allUnits || []).filter(x => x.site_id === context.site_id)
      const visibleSites = adminHO ? (allSites || []) : (allSites || []).filter(x => x.id === context.site_id)
      setRows(scopedUnits.sort((a,b)=>sortBySiteAndName(a,b,'unit_name','unit_code')))
      setSites(visibleSites)
      setForm(f => ({ ...f, site_id: f.site_id || context.site_id || visibleSites[0]?.id || '' }))
    } catch (err) {
      setMessage(err.message || 'Gagal memuat master unit.')
    }
  }
  function reset(){ setEditing(null); setForm({ site_id: context.site_id || '', unit_code:AUTO_CODE_LABEL, unit_name:'', unit_type:'', location:'', status:'Aktif' }) }
  function edit(r){ setEditing(r.id); setForm({ site_id:r.site_id, unit_code:r.unit_code, unit_name:r.unit_name, unit_type:r.unit_type || '', location:r.location || '', status:r.status || 'Aktif' }); window.scrollTo({ top:0, behavior:'smooth' }) }
  async function save(e){
    e.preventDefault(); setMessage('')
    const payload = {
      site_id: adminHO ? form.site_id : context.site_id,
      unit_name: clean(form.unit_name),
      unit_type: clean(form.unit_type),
      location: clean(form.location),
      status: form.status || 'Aktif',
      created_by: profile.id
    }
    if (!editing) payload.unit_code = generatedRecordCode('UNIT')
    const res = editing ? await supabase.from('inspection_units').update(payload).eq('id', editing) : await supabase.from('inspection_units').insert(payload)
    setMessage(res.error ? res.error.message : editing ? 'Unit berhasil diupdate.' : 'Unit berhasil ditambahkan.')
    if (!res.error){ reset(); load() }
  }
  async function remove(r){
    if (!confirm(`Hapus unit ${r.unit_code}?`)) return
    const { error } = await supabase.from('inspection_units').delete().eq('id', r.id)
    setMessage(error ? error.message : 'Unit berhasil dihapus.'); load()
  }
  async function previewFile(file){
    if (!file) return
    const raw = await parseExcel(file)
    const allSites = await fetchAllRows('sites', 'id,site_code', q => q.order('site_code'))
    const mapped = raw.map((r, idx) => {
      const siteCode = adminHO ? clean(r.site_code).toUpperCase() : context.sites?.site_code
      const site = (allSites || []).find(s => s.site_code === siteCode)
      let error = ''
      if (!site) error = 'site_code tidak ditemukan'
      else if (!clean(r.unit_name)) error = 'unit_name wajib. unit_code akan dibuat otomatis oleh sistem'
      return { row:idx+2, site_code:siteCode, site_id:site?.id, unit_code:generatedPreviewCode('UNIT', idx+1), unit_name:clean(r.unit_name), unit_type:clean(r.unit_type), location:clean(r.location), status:clean(r.status)||'Aktif', error }
    })
    setPreview(mapped)
  }
  async function submitImport(){
    const valid = preview.filter(r => !r.error)
    if (!valid.length) return setMessage('Tidak ada baris valid.')
    const payload = valid.map(r => ({
      site_id:r.site_id,
      unit_code: generatedRecordCode('UNIT'),
      unit_name:r.unit_name,
      unit_type:r.unit_type,
      location:r.location,
      status:r.status,
      created_by:profile.id
    }))
    const { error } = await supabase.from('inspection_units').insert(payload)
    setMessage(error ? error.message : `Import unit berhasil: ${payload.length} baris.`)
    if (!error){ setPreview([]); load() }
  }
  const table = rows.map(r => ({ site:r.sites?.site_code, unit_code:r.unit_code, unit_name:r.unit_name, type:r.unit_type, lokasi:r.location, status:r.status, aksi:'' }))
  return <div className="stack">
    <Panel title={editing ? 'Edit Master Unit' : 'Tambah Master Unit'} desc={adminHO ? 'Administrator JIEP dapat mengelola unit semua site.' : 'GL hanya bisa mengelola unit pada site sendiri.'}>
      <form className="form-grid" onSubmit={save}>
        {adminHO && <label>Site<select required value={form.site_id} onChange={e=>setForm({...form,site_id:e.target.value})}>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>}
        <label>Kode Unit<input disabled value={editing ? form.unit_code : AUTO_CODE_LABEL} title="Kode dibuat otomatis oleh sistem"/></label>
        <label>Nama Unit<input required value={form.unit_name} onChange={e=>setForm({...form,unit_name:e.target.value})}/></label>
        <label>Tipe Unit<input value={form.unit_type} onChange={e=>setForm({...form,unit_type:e.target.value})}/></label>
        <label>Lokasi<input value={form.location} onChange={e=>setForm({...form,location:e.target.value})}/></label>
        <label>Status<select value={form.status} onChange={e=>setForm({...form,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label>
        <button>{editing ? 'Update Unit' : 'Simpan Unit'}</button>{editing && <button type="button" className="secondary" onClick={reset}>Batal</button>}
      </form>{message && <p className="message">{message}</p>}
    </Panel>
    <Panel title="Import Master Unit Excel" desc="Kode unit dibuat otomatis oleh sistem. User cukup isi site, nama unit, tipe, lokasi, dan status.">
      <div className="import-actions"><button className="secondary" onClick={()=>downloadTemplate('unit')}><Download size={16}/> Download Template</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewFile(e.target.files?.[0])}/></label>{preview.length>0 && <button disabled={!preview.some(r=>!r.error)} onClick={submitImport}>Submit {preview.filter(r=>!r.error).length} Baris Valid</button>}</div>
      {preview.length>0 && <PreviewTable rows={preview}/>} 
    </Panel>
    <Panel title="Row Data Master Unit" action={<button onClick={()=>downloadXlsx('master-unit.xlsx', table.map(({aksi,...r})=>r))}><Download size={16}/> Export Excel</button>}><DataTable rows={table} customActions={(idx)=>{ const r=rows[idx]; return <div className="row-actions"><button className="secondary small" onClick={()=>edit(r)}>Edit</button><button className="danger small" onClick={()=>remove(r)}>Delete</button></div> }} /></Panel>
  </div>
}

function MasterParkings({ profile, context }){
  const [rows, setRows] = useState([])
  const [sites, setSites] = useState([])
  const [editing, setEditing] = useState(null)
  const [message, setMessage] = useState('')
  const [preview, setPreview] = useState([])
  const adminHO = isHeadOfficeAdmin(context)
  const [form, setForm] = useState({ site_id: context.site_id || '', parking_code:AUTO_CODE_LABEL, parking_name:'', location:'', capacity:'', status:'Aktif' })
  useEffect(() => { load() }, [context.id])
  async function load(){
    try {
      const [allParkings, allSites] = await Promise.all([
        fetchAllRows('inspection_parkings', '*, sites(site_code,site_name)', q => q.order('created_at', { ascending:false })),
        fetchAllRows('sites', '*', q => q.neq('site_code','JIEP').eq('status','Aktif').order('site_code'))
      ])
      const scopedParkings = adminHO ? (allParkings || []) : (allParkings || []).filter(x => x.site_id === context.site_id)
      const visibleSites = adminHO ? (allSites || []) : (allSites || []).filter(x => x.id === context.site_id)
      setRows(scopedParkings.sort((a,b)=>sortBySiteAndName(a,b,'parking_name','parking_code')))
      setSites(visibleSites)
      setForm(f => ({ ...f, site_id: f.site_id || context.site_id || visibleSites[0]?.id || '' }))
    } catch (err) {
      setMessage(err.message || 'Gagal memuat master parkiran.')
    }
  }
  function reset(){ setEditing(null); setForm({ site_id: context.site_id || '', parking_code:AUTO_CODE_LABEL, parking_name:'', location:'', capacity:'', status:'Aktif' }) }
  function edit(r){ setEditing(r.id); setForm({ site_id:r.site_id, parking_code:r.parking_code, parking_name:r.parking_name, location:r.location || '', capacity:r.capacity || '', status:r.status || 'Aktif' }); window.scrollTo({ top:0, behavior:'smooth' }) }
  async function save(e){
    e.preventDefault(); setMessage('')
    const payload = {
      site_id: adminHO ? form.site_id : context.site_id,
      parking_name: clean(form.parking_name),
      location: clean(form.location),
      capacity: form.capacity ? Number(form.capacity) : null,
      status: form.status || 'Aktif',
      created_by: profile.id
    }
    if (!editing) payload.parking_code = generatedRecordCode('PARK')
    const res = editing ? await supabase.from('inspection_parkings').update(payload).eq('id', editing) : await supabase.from('inspection_parkings').insert(payload)
    setMessage(res.error ? res.error.message : editing ? 'Parkiran berhasil diupdate.' : 'Parkiran berhasil ditambahkan.')
    if (!res.error){ reset(); load() }
  }
  async function remove(r){ if (!confirm(`Hapus parkiran ${r.parking_code}?`)) return; const { error } = await supabase.from('inspection_parkings').delete().eq('id', r.id); setMessage(error ? error.message : 'Parkiran berhasil dihapus.'); load() }
  async function previewFile(file){
    if (!file) return
    const raw = await parseExcel(file)
    const allSites = await fetchAllRows('sites', 'id,site_code', q => q.order('site_code'))
    const mapped = raw.map((r, idx) => {
      const siteCode = adminHO ? clean(r.site_code).toUpperCase() : context.sites?.site_code
      const site = (allSites || []).find(s => s.site_code === siteCode)
      let error = ''
      if (!site) error = 'site_code tidak ditemukan'
      else if (!clean(r.parking_name)) error = 'parking_name wajib. parking_code akan dibuat otomatis oleh sistem'
      return { row:idx+2, site_code:siteCode, site_id:site?.id, parking_code:generatedPreviewCode('PARK', idx+1), parking_name:clean(r.parking_name), location:clean(r.location), capacity:clean(r.capacity), status:clean(r.status)||'Aktif', error }
    })
    setPreview(mapped)
  }
  async function submitImport(){
    const valid = preview.filter(r => !r.error)
    if (!valid.length) return setMessage('Tidak ada baris valid.')
    const payload = valid.map(r => ({
      site_id:r.site_id,
      parking_code: generatedRecordCode('PARK'),
      parking_name:r.parking_name,
      location:r.location,
      capacity:r.capacity ? Number(r.capacity) : null,
      status:r.status,
      created_by:profile.id
    }))
    const { error } = await supabase.from('inspection_parkings').insert(payload)
    setMessage(error ? error.message : `Import parkiran berhasil: ${payload.length} baris.`)
    if (!error){ setPreview([]); load() }
  }
  const table = rows.map(r => ({ site:r.sites?.site_code, parking_code:r.parking_code, parking_name:r.parking_name, lokasi:r.location, kapasitas:r.capacity ?? '-', status:r.status, aksi:'' }))
  return <div className="stack">
    <Panel title={editing ? 'Edit Master Parkiran' : 'Tambah Master Parkiran'} desc={adminHO ? 'Administrator JIEP dapat mengelola parkiran semua site.' : 'GL hanya bisa mengelola parkiran pada site sendiri.'}>
      <form className="form-grid" onSubmit={save}>
        {adminHO && <label>Site<select required value={form.site_id} onChange={e=>setForm({...form,site_id:e.target.value})}>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>}
        <label>Kode Parkiran<input disabled value={editing ? form.parking_code : AUTO_CODE_LABEL} title="Kode dibuat otomatis oleh sistem"/></label>
        <label>Nama Parkiran<input required value={form.parking_name} onChange={e=>setForm({...form,parking_name:e.target.value})}/></label>
        <label>Lokasi<input value={form.location} onChange={e=>setForm({...form,location:e.target.value})}/></label>
        <label>Kapasitas<input type="number" value={form.capacity} onChange={e=>setForm({...form,capacity:e.target.value})}/></label>
        <label>Status<select value={form.status} onChange={e=>setForm({...form,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label>
        <button>{editing ? 'Update Parkiran' : 'Simpan Parkiran'}</button>{editing && <button type="button" className="secondary" onClick={reset}>Batal</button>}
      </form>{message && <p className="message">{message}</p>}
    </Panel>
    <Panel title="Import Master Parkiran Excel" desc="Kode parkiran dibuat otomatis oleh sistem. User cukup isi site, nama parkiran, lokasi, kapasitas, dan status.">
      <div className="import-actions"><button className="secondary" onClick={()=>downloadTemplate('parking')}><Download size={16}/> Download Template</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewFile(e.target.files?.[0])}/></label>{preview.length>0 && <button disabled={!preview.some(r=>!r.error)} onClick={submitImport}>Submit {preview.filter(r=>!r.error).length} Baris Valid</button>}</div>
      {preview.length>0 && <PreviewTable rows={preview}/>} 
    </Panel>
    <Panel title="Row Data Master Parkiran" action={<button onClick={()=>downloadXlsx('master-parkiran.xlsx', table.map(({aksi,...r})=>r))}><Download size={16}/> Export Excel</button>}><DataTable rows={table} customActions={(idx)=>{ const r=rows[idx]; return <div className="row-actions"><button className="secondary small" onClick={()=>edit(r)}>Edit</button><button className="danger small" onClick={()=>remove(r)}>Delete</button></div> }} /></Panel>
  </div>
}

function ParameterChecklist(){
  const [rows, setRows] = useState([])
  const [editing, setEditing] = useState(null)
  const [message, setMessage] = useState('')
  const [preview, setPreview] = useState([])
  const [form, setForm] = useState({ category:CATEGORIES[0], parameter_code:AUTO_CODE_LABEL, parameter_name:'', description:'', severity:'Medium', status:'Aktif' })
  useEffect(() => { load() }, [])
  async function load(){ const data = await fetchAllRows('inspection_parameters', '*', q => q.order('category').order('parameter_code')); setRows(data || []) }
  function reset(){ setEditing(null); setForm({ category:CATEGORIES[0], parameter_code:AUTO_CODE_LABEL, parameter_name:'', description:'', severity:'Medium', status:'Aktif' }) }
  function edit(r){ setEditing(r.id); setForm({ category:r.category, parameter_code:r.parameter_code || '', parameter_name:r.parameter_name, description:r.description || '', severity:r.severity || 'Medium', status:r.status || 'Aktif' }); window.scrollTo({ top:0, behavior:'smooth' }) }
  async function save(e){
    e.preventDefault(); setMessage('')
    const payload = {
      category:form.category,
      parameter_name: clean(form.parameter_name),
      description: clean(form.description),
      severity: form.severity || 'Medium',
      status: form.status || 'Aktif'
    }
    if (!editing) payload.parameter_code = generatedRecordCode(parameterPrefixByCategory(form.category))
    const res = editing ? await supabase.from('inspection_parameters').update(payload).eq('id', editing) : await supabase.from('inspection_parameters').insert(payload)
    setMessage(res.error ? res.error.message : editing ? 'Parameter berhasil diupdate.' : 'Parameter berhasil ditambahkan.')
    if (!res.error){ reset(); load() }
  }
  async function remove(r){ if (!confirm(`Hapus parameter ${r.parameter_code}?`)) return; const { error } = await supabase.from('inspection_parameters').delete().eq('id', r.id); setMessage(error ? error.message : 'Parameter berhasil dihapus.'); load() }
  async function previewFile(file){
    if (!file) return
    const raw = await parseExcel(file)
    setPreview(raw.map((r, idx) => {
      const category = clean(r.category)
      let error = ''
      if (!CATEGORIES.includes(category)) error = 'category harus Inspeksi Unit / Inspeksi Kelayakan Parkiran / PM Check'
      else if (!clean(r.parameter_name)) error = 'parameter_name wajib. parameter_code akan dibuat otomatis oleh sistem'
      return { row:idx+2, category, parameter_code:generatedPreviewCode(parameterPrefixByCategory(category), idx+1), parameter_name:clean(r.parameter_name), description:clean(r.description), severity:clean(r.severity)||'Medium', status:clean(r.status)||'Aktif', error }
    }))
  }
  async function submitImport(){
    const valid = preview.filter(r => !r.error).map(({row,error,parameter_code,...r}) => ({
      ...r,
      parameter_code: generatedRecordCode(parameterPrefixByCategory(r.category))
    }))
    if (!valid.length) return setMessage('Tidak ada baris valid.')
    const { error } = await supabase.from('inspection_parameters').insert(valid)
    setMessage(error ? error.message : `Import parameter berhasil: ${valid.length} baris.`)
    if (!error){ setPreview([]); load() }
  }
  const table = rows.map(r => ({ kategori:r.category, code:r.parameter_code, parameter:r.parameter_name, severity:r.severity, status:r.status, aksi:'' }))
  return <div className="stack">
    <Panel title={editing ? 'Edit Parameter Checklist' : 'Tambah Parameter Checklist'} desc="Parameter dipakai sebagai checklist pada saat inspeksi. Jika hasil Tidak Aman, parameter ini akan menjadi temuan setelah approval.">
      <form className="form-grid" onSubmit={save}>
        <label>Kategori<select value={form.category} onChange={e=>setForm({...form,category:e.target.value})}>{CATEGORIES.map(c=><option key={c}>{c}</option>)}</select></label>
        <label>Kode Parameter<input disabled value={editing ? form.parameter_code : AUTO_CODE_LABEL} title="Kode dibuat otomatis oleh sistem"/></label>
        <label>Nama Parameter<input required value={form.parameter_name} onChange={e=>setForm({...form,parameter_name:e.target.value})}/></label>
        <label>Severity<select value={form.severity} onChange={e=>setForm({...form,severity:e.target.value})}><option>Low</option><option>Medium</option><option>High</option></select></label>
        <label>Status<select value={form.status} onChange={e=>setForm({...form,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label>
        <label className="span-2">Deskripsi<textarea value={form.description} onChange={e=>setForm({...form,description:e.target.value})}/></label>
        <button>{editing ? 'Update Parameter' : 'Simpan Parameter'}</button>{editing && <button type="button" className="secondary" onClick={reset}>Batal</button>}
      </form>{message && <p className="message">{message}</p>}
    </Panel>
    <Panel title="Import Parameter Excel" desc="parameter_code dibuat otomatis oleh sistem. User cukup isi kategori, nama parameter, deskripsi, severity, dan status."><div className="import-actions"><button className="secondary" onClick={()=>downloadTemplate('parameter')}><Download size={16}/> Download Template</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewFile(e.target.files?.[0])}/></label>{preview.length>0 && <button disabled={!preview.some(r=>!r.error)} onClick={submitImport}>Submit {preview.filter(r=>!r.error).length} Baris Valid</button>}</div>{preview.length>0 && <PreviewTable rows={preview}/>}</Panel>
    <Panel title="Row Data Parameter Checklist" action={<button onClick={()=>downloadXlsx('parameter-checklist.xlsx', table.map(({aksi,...r})=>r))}><Download size={16}/> Export Excel</button>}><DataTable rows={table} customActions={(idx)=>{ const r=rows[idx]; return <div className="row-actions"><button className="secondary small" onClick={()=>edit(r)}>Edit</button><button className="danger small" onClick={()=>remove(r)}>Delete</button></div> }} /></Panel>
  </div>
}

function PlanBulanan({ profile, context }){
  const [plans, setPlans] = useState([])
  const [units, setUnits] = useState([])
  const [parkings, setParkings] = useState([])
  const [sites, setSites] = useState([])
  const [message, setMessage] = useState('')
  const [preview, setPreview] = useState([])
  const adminHO = isHeadOfficeAdmin(context)
  const [form, setForm] = useState({ site_id: context.site_id || '', category:CATEGORIES[0], target_id:'', bulan:String(new Date().getMonth()+1), tahun:String(new Date().getFullYear()), due_date:'' })
  useEffect(() => { load() }, [context.id])
  async function load(){
    try {
      const [allUnits, allParkings, allPlans, allSites] = await Promise.all([
        fetchAllRows('inspection_units', '*, sites(site_code,site_name)', q => q.order('unit_name')),
        fetchAllRows('inspection_parkings', '*, sites(site_code,site_name)', q => q.order('parking_name')),
        fetchAllRows('inspection_plans', '*, sites(site_code,site_name), inspection_units(unit_code,unit_name), inspection_parkings(parking_code,parking_name)', q => q.order('created_at', { ascending:false })),
        fetchAllRows('sites', '*', q => q.neq('site_code','JIEP').eq('status','Aktif').order('site_code'))
      ])
      const scopedUnits = (adminHO ? allUnits : allUnits.filter(x => x.site_id === context.site_id)).filter(x => isActiveMasterStatus(x.status)).sort((a,b)=>sortBySiteAndName(a,b,'unit_name','unit_code'))
      const scopedParkings = (adminHO ? allParkings : allParkings.filter(x => x.site_id === context.site_id)).filter(x => isActiveMasterStatus(x.status)).sort((a,b)=>sortBySiteAndName(a,b,'parking_name','parking_code'))
      const scopedPlans = adminHO ? allPlans : allPlans.filter(x => x.site_id === context.site_id)
      const visibleSites = adminHO ? (allSites || []) : (allSites || []).filter(x => x.id === context.site_id)
      setUnits(scopedUnits || [])
      setParkings(scopedParkings || [])
      setPlans(scopedPlans || [])
      setSites(visibleSites)
      const firstTarget = (isUnitBasedCategory(form.category) ? scopedUnits[0]?.id : scopedParkings[0]?.id) || ''
      setForm(f => ({ ...f, site_id:f.site_id || context.site_id || visibleSites[0]?.id || '', target_id:f.target_id || firstTarget }))
    } catch (err) {
      setMessage(err.message || 'Gagal memuat data plan inspeksi.')
    }
  }
  const filteredTargets = (isUnitBasedCategory(form.category) ? units : parkings).filter(t => adminHO ? (!form.site_id || t.site_id === form.site_id) : t.site_id === context.site_id)
  useEffect(() => { setForm(f => ({ ...f, target_id: filteredTargets[0]?.id || '' })) }, [form.category, form.site_id, units.length, parkings.length])
  async function save(e){
    e.preventDefault(); setMessage('')
    const targetType = targetTypeByCategory(form.category)
    const target = filteredTargets.find(t => t.id === form.target_id)
    const payload = { site_id: adminHO ? form.site_id : context.site_id, category:form.category, target_type:targetType, unit_id: targetType === 'unit' ? form.target_id : null, parking_id: targetType === 'parkiran' ? form.target_id : null, bulan:Number(form.bulan), tahun:new Date().getFullYear(), due_date:form.due_date || null, status:'Planned', planned_by:profile.id }
    if (!target) return setMessage('Target inspeksi belum dipilih / belum ada master data.')
    if (form.category === 'PM Check' && !form.due_date) return setMessage('Due Date wajib diisi untuk kategori PM Check agar tanggal planning jelas.')
    const { error } = await supabase.from('inspection_plans').insert(payload)
    setMessage(error ? error.message : 'Plan inspeksi berhasil dibuat.')
    if (!error) load()
  }
  async function remove(p){ if (!confirm('Hapus plan ini?')) return; const { error } = await supabase.from('inspection_plans').delete().eq('id', p.id); setMessage(error ? error.message : 'Plan berhasil dihapus.'); load() }
  async function previewFile(file){
    if (!file) return
    const raw = await parseExcel(file)
    try {
      const [allSites, allUnits, allParkings] = await Promise.all([
        fetchAllRows('sites', 'id,site_code,site_name,status', q => q.order('site_code')),
        fetchAllRows('inspection_units', 'id,site_id,unit_code,unit_name,status', q => q.order('unit_code')),
        fetchAllRows('inspection_parkings', 'id,site_id,parking_code,parking_name,status', q => q.order('parking_code'))
      ])
      const siteMap = new Map((allSites || []).map(s => [normalizeTargetKey(s.site_code), s]))
      const activeUnits = (allUnits || []).filter(u => isActiveMasterStatus(u.status))
      const activeParkings = (allParkings || []).filter(p => isActiveMasterStatus(p.status))
      setPreview(raw.map((r, idx) => {
        const siteCode = adminHO ? clean(r.site_code || r.site).toUpperCase() : context.sites?.site_code
        const site = siteMap.get(normalizeTargetKey(siteCode))
        const rawCategory = clean(r.category || r.kategori)
        const category = CATEGORIES.find(c => normalizeTargetKey(c) === normalizeTargetKey(rawCategory)) || rawCategory
        const targetName = clean(r.target_name || r.unit_name || r.parking_name || r.target || r.nama_unit || r.nama_parkiran || r.kode_unit || r.kode_parkiran)
        const targetCode = clean(r.target_code || r.unit_code || r.parking_code || r.kode_target || r.kode_unit || r.kode_parkiran)
        const dueRaw = r.due_date ?? r.due ?? r.due_date_pm ?? r.tanggal ?? r.tanggal_planning ?? ''
        const dueDate = parseFlexibleDate(dueRaw)
        const targetType = targetTypeByCategory(category)
        const source = targetType === 'unit' ? activeUnits : activeParkings
        const target = source.find(t => {
          if (t.site_id !== site?.id) return false
          return targetType === 'unit'
            ? targetMatches(t, targetName, targetCode, 'unit_code', 'unit_name')
            : targetMatches(t, targetName, targetCode, 'parking_code', 'parking_name')
        })
        const sameTargetOtherSite = !target ? source.find(t => targetType === 'unit'
          ? targetMatches(t, targetName, targetCode, 'unit_code', 'unit_name')
          : targetMatches(t, targetName, targetCode, 'parking_code', 'parking_name')) : null
        const otherSite = sameTargetOtherSite ? (allSites || []).find(s => s.id === sameTargetOtherSite.site_id) : null
        let error = ''
        if (!site) error = 'site_code tidak ditemukan'
        else if (!CATEGORIES.includes(category)) error = 'category tidak valid'
        else if (!targetName && !targetCode) error = 'target_name atau target_code wajib diisi'
        else if (!target && otherSite) error = `target ditemukan di site ${otherSite.site_code}, bukan ${siteCode}`
        else if (!target) error = 'target tidak ditemukan pada Master Unit/Parkiran aktif. Cek site, status aktif, dan pastikan data sudah termuat semua.'
        else if (!r.bulan) error = 'bulan wajib diisi; tahun otomatis berdasarkan created date'
        else if (clean(dueRaw) && !dueDate) error = 'due_date tidak valid. Gunakan format yyyy-mm-dd, dd/mm/yyyy, atau tanggal Excel.'
        else if (category === 'PM Check' && !dueDate) error = 'due_date wajib untuk PM Check'
        return { row:idx+2, site_code:siteCode, site_id:site?.id, category, target_name:targetName || targetLabel({ target_type:targetType, inspection_units:target, inspection_parkings:target }), target_code:targetCode || (targetType === 'unit' ? target?.unit_code : target?.parking_code), target_id:target?.id, target_type:targetType, bulan:Number(r.bulan), tahun:new Date().getFullYear(), due_date:dueDate, error }
      }))
    } catch (err) {
      setMessage(err.message || 'Gagal membaca master unit/parkiran.')
    }
  }
  async function submitImport(){
    const valid = preview.filter(r => !r.error)
    if (!valid.length) return setMessage('Tidak ada baris valid.')
    const payload = valid.map(r => ({ site_id:r.site_id, category:r.category, target_type:r.target_type, unit_id:r.target_type === 'unit' ? r.target_id : null, parking_id:r.target_type === 'parkiran' ? r.target_id : null, bulan:r.bulan, tahun:new Date().getFullYear(), due_date:r.due_date || null, status:'Planned', planned_by:profile.id }))
    const { error } = await supabase.from('inspection_plans').insert(payload)
    setMessage(error ? error.message : `Import plan berhasil: ${payload.length} baris.`)
    if (!error){ setPreview([]); load() }
  }
  const activePlans = plans.filter(p => ['Planned','Rejected'].includes(p.status))
  const rows = plans.map(p => ({ site:p.sites?.site_code, kategori:p.category, target:targetLabel(p), bulan:monthLabel(p.bulan), tahun:planYear(p), due_date:p.due_date || '-', status:p.status, aksi:'' }))
  return <div className="stack">
    <Panel title="Buat Plan Inspeksi Bulanan" desc="Plan yang sudah submitted/approved otomatis hilang dari menu Inspeksi. Jika rejected, plan kembali muncul untuk inspeksi ulang.">
      <form className="form-grid" onSubmit={save}>
        {adminHO && <label>Site<select required value={form.site_id} onChange={e=>setForm({...form,site_id:e.target.value})}>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>}
        <label>Kategori<select value={form.category} onChange={e=>setForm({...form,category:e.target.value})}>{CATEGORIES.map(c=><option key={c}>{c}</option>)}</select></label>
        <label>Target<select required value={form.target_id} onChange={e=>setForm({...form,target_id:e.target.value})}>{filteredTargets.map(t=><option key={t.id} value={t.id}>{isUnitBasedCategory(form.category) ? `${t.unit_code} - ${t.unit_name}` : `${t.parking_code} - ${t.parking_name}`}</option>)}</select></label>
        <label>Bulan<select value={form.bulan} onChange={e=>setForm({...form,bulan:e.target.value})}>{MONTHS.map(m=><option key={m[0]} value={m[0]}>{m[1]}</option>)}</select></label>
        <label>Tahun<input type="number" disabled value={new Date().getFullYear()} title="Tahun otomatis berdasarkan created date"/></label>
        <label>Due Date<input type="date" value={form.due_date} onChange={e=>setForm({...form,due_date:e.target.value})}/></label>
        <button>Simpan Plan</button>
      </form>{message && <p className="message">{message}</p>}
    </Panel>
    <Panel title="Import Plan Inspeksi Excel" desc="Template bisa memakai target_name atau target_code. Isi kode/nama unit atau parkiran sesuai master data aktif.">
      <div className="import-actions"><button className="secondary" onClick={()=>downloadTemplate('plan')}><Download size={16}/> Download Template</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewFile(e.target.files?.[0])}/></label>{preview.length>0 && <button disabled={!preview.some(r=>!r.error)} onClick={submitImport}>Submit {preview.filter(r=>!r.error).length} Baris Valid</button>}</div>
      {preview.length>0 && <PreviewTable rows={preview}/>} 
    </Panel>
    <Panel title="Plan Aktif untuk Inspeksi" desc="Hanya status Planned dan Rejected yang muncul di sini."><DataTable rows={activePlans.map(p => ({ site:p.sites?.site_code, kategori:p.category, target:targetLabel(p), bulan:monthLabel(p.bulan), tahun:planYear(p), due:p.due_date || '-', status:p.status }))}/></Panel>
    <Panel title="Row Data Plan Inspeksi" action={<button onClick={()=>downloadXlsx('plan-inspeksi-bulanan.xlsx', rows.map(({aksi,...r})=>r))}><Download size={16}/> Export Excel</button>}><DataTable rows={rows} customActions={(idx)=>{ const p=plans[idx]; return ['Planned','Rejected'].includes(p.status) ? <button className="danger small" onClick={()=>remove(p)}>Delete</button> : <span className="muted">Locked</span> }} /></Panel>
  </div>
}

function InspectionExecution({ profile, context }){
  const [plans, setPlans] = useState([])
  const [params, setParams] = useState([])
  const [search, setSearch] = useState('')
  const [active, setActive] = useState(null)
  const [answers, setAnswers] = useState({})
  const [notes, setNotes] = useState('')
  const [photo, setPhoto] = useState(null)
  const [message, setMessage] = useState('')
  const [submitting, setSubmitting] = useState(false)
  useEffect(() => { load() }, [context.id])
  async function load(){
    setMessage('')
    try{
      const [planRows, paramRows, siteRows, unitRows, parkingRows] = await Promise.all([
        fetchAllRows('inspection_plans', '*', q => scopeQuery(q.in('status', ['Planned','Rejected']).order('created_at', { ascending:false }), context)),
        fetchAllRows('inspection_parameters', '*', q => q.order('parameter_code')),
        fetchAllRows('sites', 'id,site_code,site_name,status', q => q.order('site_code')),
        fetchAllRows('inspection_units', 'id,site_id,unit_code,unit_name,status', q => q.order('unit_code')),
        fetchAllRows('inspection_parkings', 'id,site_id,parking_code,parking_name,status', q => q.order('parking_code'))
      ])
      const siteMap = new Map((siteRows || []).map(x => [x.id, x]))
      const unitMap = new Map((unitRows || []).map(x => [x.id, x]))
      const parkingMap = new Map((parkingRows || []).map(x => [x.id, x]))
      const enriched = (planRows || []).map(p => ({
        ...p,
        sites: siteMap.get(p.site_id) || null,
        inspection_units: unitMap.get(p.unit_id) || null,
        inspection_parkings: parkingMap.get(p.parking_id) || null
      }))
      setPlans(enriched)
      setParams((paramRows || []).filter(x => isActiveMasterStatus(x.status)))
    }catch(err){
      console.error('[INSPECTION LOAD ERROR]', err)
      setMessage(err.message || 'Gagal mengambil data inspeksi.')
      setPlans([]); setParams([])
    }
  }
  function openPlan(plan){
    const relevant = params.filter(p => p.category === plan.category)
    const obj = {}
    relevant.forEach(p => { obj[p.id] = { result:'Aman', notes:'' } })
    setActive(plan); setAnswers(obj); setNotes(''); setPhoto(null); setMessage('')
  }
  async function submitInspection(){
    if (!active) return
    const relevant = params.filter(p => p.category === active.category)
    if (!relevant.length) return setMessage('Parameter checklist untuk kategori ini masih kosong.')
    setSubmitting(true); setMessage('')
    try{
      const photoUrl = photo ? await uploadCompressedImage('inspection-unit-photos', photo, 'inspection') : null
      const finalResult = Object.values(answers).some(a => a.result === 'Tidak Aman') ? 'Tidak Aman' : 'Aman'
      const { data:record, error:recordError } = await supabase.from('inspection_records').insert({
        plan_id: active.id, site_id: active.site_id, category: active.category, inspector_id: profile.id,
        inspected_at: nowISO(), result: finalResult, status:'Submitted', notes, photo_url: photoUrl
      }).select().single()
      if (recordError) throw recordError
      const payload = relevant.map(p => ({ record_id:record.id, parameter_id:p.id, result:answers[p.id]?.result || 'Aman', notes:answers[p.id]?.notes || null }))
      const { error:answerError } = await supabase.from('inspection_answers').insert(payload)
      if (answerError) throw answerError
      const { error:planError } = await supabase.from('inspection_plans').update({ status:'Submitted', updated_at:nowISO() }).eq('id', active.id)
      if (planError) throw planError
      setMessage('Inspeksi berhasil disubmit. Menunggu approval atasan/admin.')
      setActive(null); load()
    }catch(e){ setMessage(e.message || String(e)) }
    finally{ setSubmitting(false) }
  }
  const filtered = plans.filter(p => `${p.sites?.site_code} ${p.category} ${targetLabel(p)} ${p.status}`.toLowerCase().includes(search.toLowerCase()))
  return <div className="stack">
    <Panel title="List Inspeksi" desc="List dibuat dalam bentuk card agar mudah discroll. Klik Mulai Inspeksi untuk membuka box fokus inspeksi." action={<div className="searchbox"><Search size={16}/><input placeholder="Search site/nama/kode target" value={search} onChange={e=>setSearch(e.target.value)}/></div>}>
      {message && <p className="message">{message}</p>}
      <div className="card-grid">
        {filtered.map(p => {
          const targetCode = targetCodeLabel(p)
          return <div className="work-card inspection-plan-card" key={p.id}>
            <div className="card-top"><StatusPill value={p.status}/><span>{p.sites?.site_code}</span></div>
            <h3>{targetTitle(p)}</h3>
            {targetCode && <small className="target-code-chip">Kode: {targetCode}</small>}
            <p>{p.category}</p>
            <small>{monthLabel(p.bulan)} {p.tahun} · Due {p.due_date || '-'}</small>
            <button onClick={()=>openPlan(p)}>Mulai Inspeksi</button>
          </div>
        })}
      </div>
      {!filtered.length && <p className="muted">Tidak ada plan yang perlu diinspeksi. Plan submitted/approved tidak ditampilkan.</p>}
    </Panel>
    {active && <Modal title={`Inspeksi: ${targetLabel(active)}`} onClose={()=>setActive(null)}>
      <div className="inspection-meta"><span>{active.sites?.site_code}</span><span>{active.category}</span><span>{monthLabel(active.bulan)} {active.tahun}</span></div>
      <label>Catatan Umum<textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Catatan umum inspeksi..." /></label>
      <label>Foto Inspeksi<input type="file" accept="image/*" onChange={e=>setPhoto(e.target.files?.[0])}/><small>Foto akan dikompres otomatis sebelum upload agar storage hemat.</small></label>
      <div className="checklist">
        {params.filter(p => p.category === active.category).map(p => <div className="check-row" key={p.id}>
          <div><b>{p.parameter_code} · {pmChecklistLabel(p, active)}</b><p>{p.description}</p><small>Severity: {p.severity}</small></div>
          <div className="answer-box">
            <select value={answers[p.id]?.result || 'Aman'} onChange={e=>setAnswers({...answers, [p.id]:{...answers[p.id], result:e.target.value}})}><option>Aman</option><option>Tidak Aman</option></select>
            <input placeholder="Catatan parameter" value={answers[p.id]?.notes || ''} onChange={e=>setAnswers({...answers, [p.id]:{...answers[p.id], notes:e.target.value}})}/>
          </div>
        </div>)}
      </div>
      <div className="modal-actions"><button className="secondary" onClick={()=>setActive(null)}>Batal</button><button disabled={submitting} onClick={submitInspection}>{submitting ? 'Submit...' : 'Submit Inspeksi'}</button></div>
      {message && <p className="message">{message}</p>}
    </Modal>}
  </div>
}

function Outstanding({ profile, context }){
  const [rows, setRows] = useState([])
  const [search, setSearch] = useState('')
  const [active, setActive] = useState(null)
  const [closeNote, setCloseNote] = useState('')
  const [file, setFile] = useState(null)
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  useEffect(() => { load() }, [context.id])
  async function load(){
    const data = await fetchAllRows('inspection_findings', '*, sites(site_code,site_name), inspection_parameters(parameter_code,parameter_name,severity)', q => scopeQuery(q.order('created_at', { ascending:false }), context))
    setRows(data || [])
  }
  async function requestClose(){
    if (!active) return
    setBusy(true); setMessage('')
    try{
      const url = file ? await uploadCompressedImage('inspection-close-evidence', file, 'close-evidence') : null
      const { error } = await supabase.from('inspection_findings').update({ status:'Close Requested', close_requested_by:profile.id, close_note:closeNote, close_photo_url:url, close_requested_at:nowISO(), updated_at:nowISO() }).eq('id', active.id)
      if (error) throw error
      setMessage('Request close outstanding berhasil dikirim. Menunggu approval atasan/admin.')
      setActive(null); setCloseNote(''); setFile(null); load()
    }catch(e){ setMessage(e.message || String(e)) }
    finally{ setBusy(false) }
  }
  const filtered = rows.filter(r => `${r.sites?.site_code} ${r.category} ${r.target_label} ${r.inspection_parameters?.parameter_name} ${r.status}`.toLowerCase().includes(search.toLowerCase()))
  const table = filtered.map(r => ({ site:r.sites?.site_code, kategori:r.category, target:r.target_label, parameter:r.inspection_parameters?.parameter_name, deskripsi:r.finding_description, status:r.status, due:r.due_date || '-', aksi:'' }))
  return <div className="stack">
    <Panel title="Outstanding / Temuan" desc="Temuan hanya masuk ke outstanding setelah inspeksi disetujui. Close outstanding wajib approval." action={<div className="searchbox"><Search size={16}/><input placeholder="Search temuan" value={search} onChange={e=>setSearch(e.target.value)}/></div>}>
      {message && <p className="message">{message}</p>}
      <div className="summary-strip"><span><b>{rows.filter(r=>r.status==='Open').length}</b> Open</span><span><b>{rows.filter(r=>r.status==='Close Requested').length}</b> Close Requested</span><span><b>{rows.filter(r=>r.status==='Closed').length}</b> Closed</span></div>
      <DataTable rows={table} customActions={(idx)=>{ const r=filtered[idx]; return r.status === 'Open' ? <button className="secondary small" onClick={()=>setActive(r)}>Close Request</button> : <StatusPill value={r.status}/> }} />
    </Panel>
    <Panel title="Export Row Data Outstanding" action={<button onClick={()=>downloadXlsx('outstanding-inspeksi.xlsx', table.map(({aksi,...r})=>r))}><Download size={16}/> Export Excel</button>}><p className="muted">Data di atas sudah siap diexport untuk laporan follow up temuan.</p></Panel>
    {active && <Modal title={`Close Outstanding: ${active.target_label}`} onClose={()=>setActive(null)}>
      <p><b>Temuan:</b> {active.finding_description}</p>
      <label>Catatan Perbaikan<textarea value={closeNote} onChange={e=>setCloseNote(e.target.value)} placeholder="Jelaskan tindakan perbaikan..." /></label>
      <label>Foto Evidence<input type="file" accept="image/*" onChange={e=>setFile(e.target.files?.[0])}/><small>Foto evidence dikompres otomatis sebelum upload.</small></label>
      <div className="modal-actions"><button className="secondary" onClick={()=>setActive(null)}>Batal</button><button disabled={busy} onClick={requestClose}>{busy ? 'Mengirim...' : 'Submit Close Request'}</button></div>
    </Modal>}
  </div>
}

function Approval({ profile, context }){
  const [records, setRecords] = useState([])
  const [closings, setClosings] = useState([])
  const [activeRecord, setActiveRecord] = useState(null)
  const [activeClose, setActiveClose] = useState(null)
  const [reason, setReason] = useState('')
  const [message, setMessage] = useState('')
  useEffect(() => { load() }, [context.id])
  async function load(){
    const [r, c] = await Promise.all([
      fetchAllRows('inspection_records', '*, sites(site_code,site_name), inspection_plans(*, inspection_units(unit_code,unit_name), inspection_parkings(parking_code,parking_name)), inspection_answers(*, inspection_parameters(parameter_code,parameter_name,severity,description))', q => scopeQuery(q.eq('status','Submitted').order('created_at', { ascending:false }), context)),
      fetchAllRows('inspection_findings', '*, sites(site_code,site_name), inspection_parameters(parameter_code,parameter_name,severity)', q => scopeQuery(q.eq('status','Close Requested').order('close_requested_at', { ascending:false }), context))
    ])
    setRecords(r || []); setClosings(c || [])
  }
  async function approveInspection(){
    if (!activeRecord) return
    setMessage('')
    const bad = (activeRecord.inspection_answers || []).filter(a => a.result === 'Tidak Aman')
    const plan = activeRecord.inspection_plans
    const target = targetLabel(plan)
    const { error:updateError } = await supabase.from('inspection_records').update({ status:'Approved', approved_by:profile.id, approved_at:nowISO() }).eq('id', activeRecord.id)
    if (updateError) return setMessage(updateError.message)
    const { error:planError } = await supabase.from('inspection_plans').update({ status:'Approved', updated_at:nowISO() }).eq('id', activeRecord.plan_id)
    if (planError) return setMessage(planError.message)
    if (bad.length){
      const payload = bad.map(a => ({
        record_id:activeRecord.id, plan_id:activeRecord.plan_id, site_id:activeRecord.site_id, category:activeRecord.category, target_label:target,
        parameter_id:a.parameter_id, finding_description:a.notes || a.inspection_parameters?.parameter_name || 'Temuan inspeksi', priority:a.inspection_parameters?.severity || 'Medium', status:'Open', created_by:activeRecord.inspector_id
      }))
      const { error:findError } = await supabase.from('inspection_findings').insert(payload)
      if (findError) return setMessage(findError.message)
    }
    setMessage(bad.length ? `Inspeksi approved. ${bad.length} temuan masuk outstanding.` : 'Inspeksi approved. Tidak ada temuan.')
    setActiveRecord(null); load()
  }
  async function rejectInspection(){
    if (!activeRecord) return
    const { error:recError } = await supabase.from('inspection_records').update({ status:'Rejected', rejection_reason:reason || 'Rejected by approver' }).eq('id', activeRecord.id)
    if (recError) return setMessage(recError.message)
    const { error:planError } = await supabase.from('inspection_plans').update({ status:'Rejected', updated_at:nowISO() }).eq('id', activeRecord.plan_id)
    setMessage(planError ? planError.message : 'Inspeksi rejected. Plan kembali open untuk inspeksi ulang.')
    setActiveRecord(null); setReason(''); load()
  }
  async function approveClose(){
    if (!activeClose) return
    const { error } = await supabase.from('inspection_findings').update({ status:'Closed', close_approved_by:profile.id, close_approved_at:nowISO(), updated_at:nowISO() }).eq('id', activeClose.id)
    setMessage(error ? error.message : 'Close outstanding approved. Temuan menjadi Closed.')
    setActiveClose(null); load()
  }
  async function rejectClose(){
    if (!activeClose) return
    const { error } = await supabase.from('inspection_findings').update({ status:'Open', close_rejection_reason:reason || 'Close rejected', updated_at:nowISO() }).eq('id', activeClose.id)
    setMessage(error ? error.message : 'Close rejected. Outstanding kembali Open.')
    setActiveClose(null); setReason(''); load()
  }
  const recRows = records.map(r => ({ site:r.sites?.site_code, kategori:r.category, target:targetLabel(r.inspection_plans), hasil:r.result, tanggal:r.inspected_at ? new Date(r.inspected_at).toLocaleString('id-ID') : '-', aksi:'' }))
  const closeRows = closings.map(f => ({ site:f.sites?.site_code, kategori:f.category, target:f.target_label, parameter:f.inspection_parameters?.parameter_name, close_note:f.close_note, status:f.status, aksi:'' }))
  return <div className="stack">
    {message && <p className="message">{message}</p>}
    <Panel title="Approval Hasil Inspeksi" desc="Klik review untuk melihat detail checklist sebelum approve/reject."><DataTable rows={recRows} customActions={(idx)=><button className="secondary small" onClick={()=>setActiveRecord(records[idx])}><Eye size={14}/> Review</button>} /></Panel>
    <Panel title="Approval Close Outstanding" desc="Close outstanding baru menjadi Closed setelah approval atasan/admin."><DataTable rows={closeRows} customActions={(idx)=><button className="secondary small" onClick={()=>setActiveClose(closings[idx])}><Eye size={14}/> Review</button>} /></Panel>
    {activeRecord && <Modal title={`Review Inspeksi: ${targetLabel(activeRecord.inspection_plans)}`} onClose={()=>setActiveRecord(null)}>
      <div className="inspection-meta"><span>{activeRecord.sites?.site_code}</span><span>{activeRecord.category}</span><span>Hasil: {activeRecord.result}</span></div>
      {activeRecord.photo_url && <a className="photo-link" href={activeRecord.photo_url} target="_blank" rel="noreferrer">Buka Foto Inspeksi</a>}
      <p><b>Catatan:</b> {activeRecord.notes || '-'}</p>
      <div className="review-list">
        {(activeRecord.inspection_answers || []).map(a => <div className="review-item" key={a.id}><div><b>{a.inspection_parameters?.parameter_code} · {pmChecklistLabel(a.inspection_parameters, activeRecord?.inspection_plans)}</b><p>{a.notes || '-'}</p></div><StatusPill value={a.result}/></div>)}
      </div>
      <label>Alasan Reject<textarea value={reason} onChange={e=>setReason(e.target.value)} placeholder="Isi jika reject..." /></label>
      <div className="modal-actions"><button className="danger" onClick={rejectInspection}><XCircle size={16}/> Reject</button><button onClick={approveInspection}><Check size={16}/> Approve</button></div>
    </Modal>}
    {activeClose && <Modal title={`Review Close: ${activeClose.target_label}`} onClose={()=>setActiveClose(null)}>
      <p><b>Temuan:</b> {activeClose.finding_description}</p><p><b>Catatan Close:</b> {activeClose.close_note || '-'}</p>
      {activeClose.close_photo_url && <a className="photo-link" href={activeClose.close_photo_url} target="_blank" rel="noreferrer">Buka Foto Evidence</a>}
      <label>Alasan Reject Close<textarea value={reason} onChange={e=>setReason(e.target.value)} placeholder="Isi jika reject close..." /></label>
      <div className="modal-actions"><button className="danger" onClick={rejectClose}><XCircle size={16}/> Reject Close</button><button onClick={approveClose}><Check size={16}/> Approve Close</button></div>
    </Modal>}
  </div>
}

function Kpi({ title, value, icon }){ return <div className="kpi"><div><small>{title}</small><strong>{value}</strong></div><div className="kpi-icon">{icon}</div></div> }
function Panel({ title, desc, action, children }){ return <section className="panel"><div className="panel-head"><div><h3>{title}</h3>{desc && <p>{desc}</p>}</div>{action}</div>{children}</section> }
function StatusPill({ value }){
  const v = String(value || '-')
  const cls = v.includes('Tidak') || v.includes('Reject') || v === 'Open' ? 'red' : v.includes('Close') || v.includes('Approved') || v === 'Aman' ? 'green' : v.includes('Submitted') || v.includes('Request') ? 'blue' : 'amber'
  return <span className={`pill ${cls}`}>{v}</span>
}
function DataTable({ rows, customActions, actions }){
  const [q, setQ] = useState('')
  if (!rows?.length) return <p className="muted">Belum ada data.</p>
  const cols = Object.keys(rows[0])
  const filtered = rows.filter(r => !q || Object.values(r).some(v => String(v ?? '').toLowerCase().includes(q.toLowerCase())))
  return <div className="data-table-block"><div className="table-search"><Search size={16}/><input placeholder="Search row data..." value={q} onChange={e=>setQ(e.target.value)} /></div><div className="table-wrap"><table><thead><tr>{cols.map(c => <th key={c}>{c}</th>)}</tr></thead><tbody>{filtered.map((r, idx) => { const originalIdx = rows.indexOf(r); return <tr key={idx}>{cols.map(c => <td key={c}>{c === 'aksi' && customActions ? customActions(originalIdx) : String(r[c] ?? '')}</td>)}</tr>})}</tbody></table></div>{filtered.length===0&&<p className="muted">Tidak ada data sesuai pencarian.</p>}</div>
}

function PreviewTable({ rows }){ return <div className="preview"><div className="summary-strip"><span><b>{rows.length}</b> Total Baris</span><span><b>{rows.filter(r=>!r.error).length}</b> Valid</span><span><b>{rows.filter(r=>r.error).length}</b> Error</span></div><DataTable rows={rows}/></div> }
function Modal({ title, children, onClose }){ return <div className="modal-backdrop"><div className="modal"><div className="modal-head"><h2>{title}</h2><button className="icon-btn" onClick={onClose}><X size={20}/></button></div>{children}</div></div> }
function FullCenter({ text, action }){ return <div className="full-center"><ShieldCheck size={48}/><h2>{text}</h2>{action}</div> }

export default App