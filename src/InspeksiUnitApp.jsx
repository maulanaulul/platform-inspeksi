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

function clean(v){ return String(v ?? '').trim() }
function normEmail(v){ return clean(v).toLowerCase() }
function today(){ return new Date().toISOString().slice(0, 10) }
function nowISO(){ return new Date().toISOString() }
function canAdmin(role){ return ADMIN_ROLES.includes(role) }
function canApprove(role){ return APPROVAL_ROLES.includes(role) }
function isHeadOfficeAdmin(context){ return canAdmin(context?.role) && (!context?.site_id || context?.sites?.site_code === 'JIEP') }
function scopeQuery(query, context){ return isHeadOfficeAdmin(context) ? query : query.eq('site_id', context.site_id) }
function siteName(context){ return context?.sites?.site_name || (isHeadOfficeAdmin(context) ? 'All Site' : '-') }
function targetLabel(plan){
  if (!plan) return '-'
  if (plan.target_type === 'unit') return `${plan.inspection_units?.unit_code || ''} ${plan.inspection_units?.unit_name || ''}`.trim()
  return `${plan.inspection_parkings?.parking_code || ''} ${plan.inspection_parkings?.parking_name || ''}`.trim()
}
function monthLabel(v){ return MONTHS.find(m => String(m[0]) === String(v))?.[1] || v }
function planYear(p){ return p?.created_at ? new Date(p.created_at).getFullYear() : p?.tahun }
function normalizeRow(row){
  const out = {}
  Object.keys(row || {}).forEach(k => { out[String(k).trim()] = row[k] })
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
    plan: [{ site_code:'BAYA', category:'PM Check', target_name:'Dump Truck 001', bulan:5, due_date:'2026-05-31' }],
    access: [{ nama:'Nama User', nrp:'NRP-001', email:'user@company.co.id', password:'password123', app_code:'inspeksi_unit', role:'GL', site_code:'BAYA' }]
  }
  downloadXlsx(`template-${type}.xlsx`, samples[type] || [{}], 'Template')
}
async function parseExcel(file){
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  return XLSX.utils.sheet_to_json(ws, { defval: '' }).map(normalizeRow)
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


function DashboardDateFilter({ dateFrom, dateTo, setDateFrom, setDateTo, onClear }) {
  return <Panel title="Filter Tanggal Dashboard" desc="KPI, achievement, chart, dan row data dashboard mengikuti rentang tanggal ini.">
    <div className="form-grid">
      <label>Dari Tanggal<input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} /></label>
      <label>Sampai Tanggal<input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} /></label>
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
  const [error, setError] = useState('')

  useEffect(() => { load() }, [context.id, dateFrom, dateTo])
  async function load(){
    setLoading(true); setError('')
    try{
      let planQ = supabase.from('inspection_plans').select('*, sites(site_code,site_name), inspection_units(unit_code,unit_name), inspection_parkings(parking_code,parking_name)').order('created_at', { ascending:false })
      let recordQ = supabase.from('inspection_records').select('*, sites(site_code,site_name), inspection_plans(*, inspection_units(unit_code,unit_name), inspection_parkings(parking_code,parking_name))').order('created_at', { ascending:false })
      let findingQ = supabase.from('inspection_findings').select('*, sites(site_code,site_name), inspection_parameters(parameter_code,parameter_name)').order('created_at', { ascending:false })
      planQ = scopeQuery(planQ, context)
      recordQ = scopeQuery(recordQ, context)
      findingQ = scopeQuery(findingQ, context)
      if (dateFrom) { planQ = planQ.gte('created_at', dateFrom); recordQ = recordQ.gte('inspected_at', dateFrom); findingQ = findingQ.gte('created_at', dateFrom) }
      if (dateTo) { const end = dateTo + 'T23:59:59'; planQ = planQ.lte('created_at', end); recordQ = recordQ.lte('inspected_at', end); findingQ = findingQ.lte('created_at', end) }
      const [planRes, recordRes, findingRes, siteRes, unitRes] = await Promise.all([
        planQ, recordQ, findingQ,
        supabase.from('sites').select('*').neq('site_code', 'JIEP').eq('status','Aktif').order('site_code'),
        scopeQuery(supabase.from('inspection_units').select('id,site_id,unit_code,unit_name,status').eq('status','Aktif'), context)
      ])
      if (planRes.error) throw planRes.error
      if (recordRes.error) throw recordRes.error
      if (findingRes.error) throw findingRes.error
      if (siteRes.error) throw siteRes.error
      if (unitRes.error) throw unitRes.error
      setPlans(planRes.data || [])
      setRecords(recordRes.data || [])
      setFindings(findingRes.data || [])
      setSites(isHeadOfficeAdmin(context) ? (siteRes.data || []) : (siteRes.data || []).filter(x => x.id === context.site_id))
      setUnits(unitRes.data || [])
    }catch(e){
      console.error('[DASHBOARD LOAD ERROR]', e)
      setError(e.message || String(e))
      setPlans([]); setRecords([]); setFindings([]); setSites([]); setUnits([])
    }finally{
      setLoading(false)
    }
  }

  const waitingApproval = records.filter(r => r.status === 'Submitted').length
  const openFindings = findings.filter(f => f.status === 'Open').length
  const closeRequested = findings.filter(f => f.status === 'Close Requested').length
  const closedFindings = findings.filter(f => f.status === 'Closed').length
  const approvedPlans = plans.filter(p => p.status === 'Approved').length
  const siteRows = sites.map(s => {
    const sitePlans = plans.filter(p => p.site_id === s.id)
    const done = sitePlans.filter(p => p.status === 'Approved').length
    const total = sitePlans.length
    return { site: s.site_code, site_name: s.site_name, total_plan: total, approved: done, belum: Math.max(total - done, 0), achievement: total ? Math.round(done / total * 100) : 0 }
  })
  const pmPlans = plans.filter(p => p.category === 'PM Check')
  const pmApproved = pmPlans.filter(p => p.status === 'Approved').length
  const pmUnitTarget = units.length * 2
  const pmAchievement = pmUnitTarget ? Math.round(pmApproved / pmUnitTarget * 100) : 0
  const pmRows = sites.map(s => {
    const siteUnits = units.filter(u => u.site_id === s.id).length
    const target = siteUnits * 2
    const actual = pmPlans.filter(p => p.site_id === s.id && p.status === 'Approved').length
    return { site:s.site_code, unit_aktif:siteUnits, target_pm_bulanan:target, pm_approved:actual, belum_pm:Math.max(target-actual,0), achievement:target ? Math.round(actual/target*100) : 0 }
  })
  const inspectionRows = records.map(r => ({
    tanggal: r.inspected_at ? new Date(r.inspected_at).toLocaleString('id-ID') : '-',
    site: r.sites?.site_code,
    kategori: r.category,
    target: targetLabel(r.inspection_plans),
    hasil: r.result,
    status: r.status,
    catatan: r.notes || '-'
  }))
  const findingRows = findings.map(f => ({
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
  return <div className="stack">
    <DashboardDateFilter dateFrom={dateFrom} dateTo={dateTo} setDateFrom={setDateFrom} setDateTo={setDateTo} onClear={()=>{setDateFrom('');setDateTo('')}} />
    <div className="kpi-grid">
      <Kpi title="Total Plan" value={plans.length} icon={<CalendarCheck/>} />
      <Kpi title="Plan Approved" value={approvedPlans} icon={<CheckCircle2/>} />
      <Kpi title="Menunggu Approval" value={waitingApproval} icon={<Eye/>} />
      <Kpi title="Temuan Open" value={openFindings} icon={<AlertTriangle/>} />
      <Kpi title="Close Request" value={closeRequested} icon={<ClipboardCheck/>} />
      <Kpi title="Temuan Closed" value={closedFindings} icon={<Check/>} />
      <Kpi title="Target PM Check" value={pmUnitTarget} icon={<Truck/>} />
      <Kpi title="Achievement PM" value={`${pmAchievement}%`} icon={<BarChart3/>} />
    </div>

    <Panel title="Dashboard PM Check" desc="Pencapaian PM Check dihitung dari target 2 kali PM Check per unit aktif setiap bulan." action={<button onClick={()=>downloadXlsx('dashboard-pm-check.xlsx', pmRows)}><Download size={16}/> Export Excel</button>}>
      <div className="summary-strip"><span><b>{units.length}</b> Unit Aktif</span><span><b>{pmUnitTarget}</b> Target PM/Bulan</span><span><b>{pmApproved}</b> PM Approved</span><span><b>{pmAchievement}%</b> Achievement PM</span></div>
      <div className="site-chart">
        {pmRows.map(r => <div className="site-bar" key={r.site}>
          <div className="site-meta"><b>{r.site}</b><span>{r.pm_approved}/{r.target_pm_bulanan} · {r.achievement}%</span></div>
          <div className="bar"><span style={{ width: `${Math.min(r.achievement, 100)}%` }} /></div>
        </div>)}
        {!pmRows.length && <p className="muted">Belum ada data PM Check.</p>}
      </div>
      <DataTable rows={pmRows}/>
    </Panel>
    <Panel title="Achievement All Site" desc="Pencapaian dihitung dari total plan inspeksi yang sudah approved. Site JIEP/Head Office tidak masuk hitungan achievement." action={<button onClick={()=>downloadXlsx('achievement-inspeksi-all-site.xlsx', siteRows)}><Download size={16}/> Export Excel</button>}>
      <div className="site-chart">
        {siteRows.map(r => <div className="site-bar" key={r.site}>
          <div className="site-meta"><b>{r.site}</b><span>{r.approved}/{r.total_plan} · {r.achievement}%</span></div>
          <div className="bar"><span style={{ width: `${Math.min(r.achievement, 100)}%` }} /></div>
        </div>)}
        {!siteRows.length && <p className="muted">Belum ada data site.</p>}
      </div>
      <DataTable rows={siteRows}/>
    </Panel>
    <Panel title="Row Data Inspeksi" desc="Tetap disediakan sebagai data mentah agar bisa diekspor Excel, bukan hanya persentase dashboard." action={<button onClick={()=>downloadXlsx('row-data-inspeksi.xlsx', inspectionRows)}><Download size={16}/> Export Excel</button>}>
      <DataTable rows={inspectionRows}/>
    </Panel>
    <Panel title="Row Data Temuan Open & Closed" desc="Seluruh temuan open, close requested, dan closed untuk kebutuhan monitoring tindak lanjut." action={<button onClick={()=>downloadXlsx('row-data-temuan-inspeksi.xlsx', findingRows)}><Download size={16}/> Export Excel</button>}>
      <div className="summary-strip"><span><b>{openFindings}</b> Open</span><span><b>{closeRequested}</b> Close Requested</span><span><b>{closedFindings}</b> Closed</span></div>
      <DataTable rows={findingRows}/>
    </Panel>
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
    const [{ data:a }, { data:s }, { data:p }, { data:ua }] = await Promise.all([
      supabase.from('applications').select('*').order('app_name'),
      supabase.from('sites').select('*').order('site_code'),
      supabase.from('users_profile').select('*').order('created_at', { ascending:false }),
      supabase.from('user_app_access').select('*, users_profile(nama,email,nrp), applications(app_code,app_name), sites(site_code,site_name)').order('created_at', { ascending:false })
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
  async function load(){ const { data } = await supabase.from('sites').select('*').order('site_code'); setSites(data || []) }
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
    let q = supabase.from('inspection_units').select('*, sites(site_code,site_name)').order('created_at', { ascending:false })
    q = scopeQuery(q, context)
    const [{ data:u }, { data:s }] = await Promise.all([q, supabase.from('sites').select('*').neq('site_code','JIEP').eq('status','Aktif').order('site_code')])
    setRows(u || []); setSites(adminHO ? (s || []) : (s || []).filter(x => x.id === context.site_id))
    setForm(f => ({ ...f, site_id: f.site_id || context.site_id || (s || [])[0]?.id || '' }))
  }
  function reset(){ setEditing(null); setForm({ site_id: context.site_id || '', unit_code:AUTO_CODE_LABEL, unit_name:'', unit_type:'', location:'', status:'Aktif' }) }
  function edit(r){ setEditing(r.id); setForm({ site_id:r.site_id, unit_code:r.unit_code, unit_name:r.unit_name, unit_type:r.unit_type || '', location:r.location || '', status:r.status || 'Aktif' }); window.scrollTo({ top:0, behavior:'smooth' }) }
  async function save(e){
    e.preventDefault(); setMessage('')
    const payload = { site_id: adminHO ? form.site_id : context.site_id, unit_name: clean(form.unit_name), unit_type: clean(form.unit_type), location: clean(form.location), status: form.status || 'Aktif', created_by: profile.id }
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
    const { data:allSites } = await supabase.from('sites').select('id,site_code')
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
    const payload = valid.map(r => ({ site_id:r.site_id, unit_name:r.unit_name, unit_type:r.unit_type, location:r.location, status:r.status, created_by:profile.id }))
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
      <div className="import-actions"><button className="secondary" onClick={()=>downloadTemplate('unit')}><Download size={16}/> Download Template</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewFile(e.target.files?.[0])}/></label>{preview.length>0 && <button disabled={preview.some(r=>r.error)} onClick={submitImport}>Submit Import Valid</button>}</div>
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
    let q = supabase.from('inspection_parkings').select('*, sites(site_code,site_name)').order('created_at', { ascending:false })
    q = scopeQuery(q, context)
    const [{ data:p }, { data:s }] = await Promise.all([q, supabase.from('sites').select('*').neq('site_code','JIEP').eq('status','Aktif').order('site_code')])
    setRows(p || []); setSites(adminHO ? (s || []) : (s || []).filter(x => x.id === context.site_id))
    setForm(f => ({ ...f, site_id: f.site_id || context.site_id || (s || [])[0]?.id || '' }))
  }
  function reset(){ setEditing(null); setForm({ site_id: context.site_id || '', parking_code:AUTO_CODE_LABEL, parking_name:'', location:'', capacity:'', status:'Aktif' }) }
  function edit(r){ setEditing(r.id); setForm({ site_id:r.site_id, parking_code:r.parking_code, parking_name:r.parking_name, location:r.location || '', capacity:r.capacity || '', status:r.status || 'Aktif' }); window.scrollTo({ top:0, behavior:'smooth' }) }
  async function save(e){
    e.preventDefault(); setMessage('')
    const payload = { site_id: adminHO ? form.site_id : context.site_id, parking_name: clean(form.parking_name), location: clean(form.location), capacity: form.capacity ? Number(form.capacity) : null, status: form.status || 'Aktif', created_by: profile.id }
    const res = editing ? await supabase.from('inspection_parkings').update(payload).eq('id', editing) : await supabase.from('inspection_parkings').insert(payload)
    setMessage(res.error ? res.error.message : editing ? 'Parkiran berhasil diupdate.' : 'Parkiran berhasil ditambahkan.')
    if (!res.error){ reset(); load() }
  }
  async function remove(r){ if (!confirm(`Hapus parkiran ${r.parking_code}?`)) return; const { error } = await supabase.from('inspection_parkings').delete().eq('id', r.id); setMessage(error ? error.message : 'Parkiran berhasil dihapus.'); load() }
  async function previewFile(file){
    if (!file) return
    const raw = await parseExcel(file)
    const { data:allSites } = await supabase.from('sites').select('id,site_code')
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
    const payload = valid.map(r => ({ site_id:r.site_id, parking_name:r.parking_name, location:r.location, capacity:r.capacity ? Number(r.capacity) : null, status:r.status, created_by:profile.id }))
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
      <div className="import-actions"><button className="secondary" onClick={()=>downloadTemplate('parking')}><Download size={16}/> Download Template</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewFile(e.target.files?.[0])}/></label>{preview.length>0 && <button disabled={preview.some(r=>r.error)} onClick={submitImport}>Submit Import Valid</button>}</div>
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
  async function load(){ const { data } = await supabase.from('inspection_parameters').select('*').order('category').order('parameter_code'); setRows(data || []) }
  function reset(){ setEditing(null); setForm({ category:CATEGORIES[0], parameter_code:AUTO_CODE_LABEL, parameter_name:'', description:'', severity:'Medium', status:'Aktif' }) }
  function edit(r){ setEditing(r.id); setForm({ category:r.category, parameter_code:r.parameter_code || '', parameter_name:r.parameter_name, description:r.description || '', severity:r.severity || 'Medium', status:r.status || 'Aktif' }); window.scrollTo({ top:0, behavior:'smooth' }) }
  async function save(e){
    e.preventDefault(); setMessage('')
    const payload = { category:form.category, parameter_name: clean(form.parameter_name), description: clean(form.description), severity: form.severity || 'Medium', status: form.status || 'Aktif' }
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
    const valid = preview.filter(r => !r.error).map(({row,error,parameter_code,...r}) => r)
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
    <Panel title="Import Parameter Excel" desc="parameter_code dibuat otomatis oleh sistem. User cukup isi kategori, nama parameter, deskripsi, severity, dan status."><div className="import-actions"><button className="secondary" onClick={()=>downloadTemplate('parameter')}><Download size={16}/> Download Template</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewFile(e.target.files?.[0])}/></label>{preview.length>0 && <button disabled={preview.some(r=>r.error)} onClick={submitImport}>Submit Import Valid</button>}</div>{preview.length>0 && <PreviewTable rows={preview}/>}</Panel>
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
    let unitQ = supabase.from('inspection_units').select('*, sites(site_code,site_name)').eq('status','Aktif').order('unit_code')
    let parkQ = supabase.from('inspection_parkings').select('*, sites(site_code,site_name)').eq('status','Aktif').order('parking_code')
    let planQ = supabase.from('inspection_plans').select('*, sites(site_code,site_name), inspection_units(unit_code,unit_name), inspection_parkings(parking_code,parking_name)').order('created_at', { ascending:false })
    unitQ = scopeQuery(unitQ, context); parkQ = scopeQuery(parkQ, context); planQ = scopeQuery(planQ, context)
    const [{ data:u }, { data:p }, { data:pl }, { data:s }] = await Promise.all([unitQ, parkQ, planQ, supabase.from('sites').select('*').neq('site_code','JIEP').eq('status','Aktif').order('site_code')])
    setUnits(u || []); setParkings(p || []); setPlans(pl || []); setSites(adminHO ? (s || []) : (s || []).filter(x => x.id === context.site_id))
    const firstTarget = (isUnitBasedCategory(form.category) ? (u || [])[0]?.id : (p || [])[0]?.id) || ''
    setForm(f => ({ ...f, site_id:f.site_id || context.site_id || (s || [])[0]?.id || '', target_id:f.target_id || firstTarget }))
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
    const [{ data:allSites }, { data:allUnits }, { data:allParkings }] = await Promise.all([
      supabase.from('sites').select('id,site_code'), supabase.from('inspection_units').select('id,site_id,unit_code,unit_name'), supabase.from('inspection_parkings').select('id,site_id,parking_code,parking_name')
    ])
    setPreview(raw.map((r, idx) => {
      const siteCode = adminHO ? clean(r.site_code).toUpperCase() : context.sites?.site_code
      const site = (allSites || []).find(s => s.site_code === siteCode)
      const category = clean(r.category)
      const targetName = clean(r.target_name || r.unit_name || r.parking_name || r.target)
      const targetCode = clean(r.target_code).toUpperCase()
      const targetType = targetTypeByCategory(category)
      const target = targetType === 'unit'
        ? (allUnits || []).find(u => u.site_id === site?.id && (clean(u.unit_name).toLowerCase() === targetName.toLowerCase() || (targetCode && u.unit_code === targetCode)))
        : (allParkings || []).find(p => p.site_id === site?.id && (clean(p.parking_name).toLowerCase() === targetName.toLowerCase() || (targetCode && p.parking_code === targetCode)))
      let error = ''
      if (!site) error = 'site_code tidak ditemukan'
      else if (!CATEGORIES.includes(category)) error = 'category tidak valid'
      else if (!target) error = 'target_name tidak ditemukan pada site/category tersebut'
      else if (!r.bulan) error = 'bulan wajib diisi; tahun otomatis berdasarkan created date'
      else if (category === 'PM Check' && !clean(r.due_date)) error = 'due_date wajib untuk PM Check'
      return { row:idx+2, site_code:siteCode, site_id:site?.id, category, target_name:targetName || targetLabel({ target_type:targetType, inspection_units:target, inspection_parkings:target }), target_id:target?.id, target_type:targetType, bulan:Number(r.bulan), tahun:new Date().getFullYear(), due_date:clean(r.due_date), error }
    }))
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
    <Panel title="Import Plan Inspeksi Excel" desc="Template memakai target_name. Pilih nama unit/parkiran dari master data; kode target dibuat otomatis oleh sistem dan tidak perlu diisi user.">
      <div className="import-actions"><button className="secondary" onClick={()=>downloadTemplate('plan')}><Download size={16}/> Download Template</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewFile(e.target.files?.[0])}/></label>{preview.length>0 && <button disabled={preview.some(r=>r.error)} onClick={submitImport}>Submit Import Valid</button>}</div>
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
    let q = supabase.from('inspection_plans').select('*').in('status', ['Planned','Rejected']).order('created_at', { ascending:false })
    q = scopeQuery(q, context)

    const [planRes, paramRes, siteRes, unitRes, parkingRes] = await Promise.all([
      q,
      supabase.from('inspection_parameters').select('*').eq('status','Aktif').order('parameter_code'),
      supabase.from('sites').select('id,site_code,site_name'),
      supabase.from('inspection_units').select('id,site_id,unit_code,unit_name'),
      supabase.from('inspection_parkings').select('id,site_id,parking_code,parking_name')
    ])

    if (planRes.error){
      console.error('[INSPECTION PLAN LOAD ERROR]', planRes.error)
      setMessage(planRes.error.message || 'Gagal mengambil data plan inspeksi.')
      setPlans([])
      return
    }
    if (paramRes.error){
      console.error('[INSPECTION PARAM LOAD ERROR]', paramRes.error)
      setMessage(paramRes.error.message || 'Gagal mengambil parameter checklist.')
    }

    const siteMap = new Map((siteRes.data || []).map(x => [x.id, x]))
    const unitMap = new Map((unitRes.data || []).map(x => [x.id, x]))
    const parkingMap = new Map((parkingRes.data || []).map(x => [x.id, x]))
    const enriched = (planRes.data || []).map(p => ({
      ...p,
      sites: siteMap.get(p.site_id) || null,
      inspection_units: unitMap.get(p.unit_id) || null,
      inspection_parkings: parkingMap.get(p.parking_id) || null
    }))

    setPlans(enriched)
    setParams(paramRes.data || [])
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
        {filtered.map(p => <div className="work-card" key={p.id}>
          <div className="card-top"><StatusPill value={p.status}/><span>{p.sites?.site_code}</span></div>
          <h3>{targetLabel(p)}</h3>
          <p>{p.category}</p>
          <small>{monthLabel(p.bulan)} {p.tahun} · Due {p.due_date || '-'}</small>
          <button onClick={()=>openPlan(p)}>Mulai Inspeksi</button>
        </div>)}
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
    let q = supabase.from('inspection_findings').select('*, sites(site_code,site_name), inspection_parameters(parameter_code,parameter_name,severity)').order('created_at', { ascending:false })
    q = scopeQuery(q, context)
    const { data } = await q
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
    let recQ = supabase.from('inspection_records').select('*, sites(site_code,site_name), inspection_plans(*, inspection_units(unit_code,unit_name), inspection_parkings(parking_code,parking_name)), inspection_answers(*, inspection_parameters(parameter_code,parameter_name,severity,description))').eq('status','Submitted').order('created_at', { ascending:false })
    let closeQ = supabase.from('inspection_findings').select('*, sites(site_code,site_name), inspection_parameters(parameter_code,parameter_name,severity)').eq('status','Close Requested').order('close_requested_at', { ascending:false })
    recQ = scopeQuery(recQ, context); closeQ = scopeQuery(closeQ, context)
    const [{ data:r }, { data:c }] = await Promise.all([recQ, closeQ])
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
