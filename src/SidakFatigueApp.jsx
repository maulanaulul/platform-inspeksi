import React, { useEffect, useMemo, useState } from 'react'
import { supabase } from './lib/supabase'
import { exportToCsv } from './utils/exportCsv'
import { compressImage } from './utils/imageCompress'
import {
  AlertTriangle, BarChart3, Building2, CalendarCheck, CheckCircle2, ClipboardCheck,
  Database, Download, FileText, KeyRound, LayoutDashboard, LogOut, Menu,
  ShieldCheck, Truck, Upload, Users, XCircle, Eye, Check, X, FileSpreadsheet, Search
} from 'lucide-react'
import * as XLSX from 'xlsx'
import DrdDriverApp from './DrdDriverApp.jsx'
import InspeksiUnitApp from './InspeksiUnitApp.jsx'
import FoodIndexApp from './FoodIndexApp.jsx'
import logoSrgs from './assets/logo-icon.png'
import loginScene from './assets/login-scene.png'
import roleScene from './assets/role-scene.png'
import './styles.css'

const APP_ICONS = { sidak_fatigue: '🚦', drd_driver: '🧠', inspeksi_unit: '🚚', inspeksi: '🚚', food_index: '🍽️' }
function normalizeAppCode(code = '') {
  const c = String(code || '').toLowerCase()
  if (c.includes('food') || c.includes('catering')) return 'food_index'
  if (c.includes('drd')) return 'drd_driver'
  if (c.includes('inspeksi')) return 'inspeksi_unit'
  if (c.includes('sidak') || c.includes('fatigue')) return 'sidak_fatigue'
  return c
}
const MONTHS = [
  ['1','Januari'], ['2','Februari'], ['3','Maret'], ['4','April'], ['5','Mei'], ['6','Juni'],
  ['7','Juli'], ['8','Agustus'], ['9','September'], ['10','Oktober'], ['11','November'], ['12','Desember']
]

function isHeadOfficeAdmin(context){
  return canAdmin(context.role) && context.sites?.site_code === 'JIEP'
}
function applySiteScope(query, context){
  return isHeadOfficeAdmin(context) ? query : query.eq('site_id', context.site_id)
}
function downloadTemplate(filename, rows){
  downloadXlsx(filename, rows)
}
function downloadXlsx(filename, rows){
  const ws = XLSX.utils.json_to_sheet(rows)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Template')
  XLSX.writeFile(wb, filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`)
}
function normalizeRow(row){
  const out = {}
  Object.keys(row || {}).forEach(k => { out[String(k).trim()] = row[k] })
  return out
}
function cleanText(v){ return String(v ?? '').trim() }
function normEmail(v){ return cleanText(v).toLowerCase() }
function isValidEmail(v){ const e = normEmail(v); return !e || /^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(e) }
function makeVendorCode(idx=0){ return `VEN-${Date.now().toString(36).toUpperCase()}-${String(idx+1).padStart(3,'0')}` }
function excelDateToIso(v){
  if(v === null || v === undefined || v === '') return ''
  if(typeof v === 'number'){
    const d = XLSX.SSF.parse_date_code(v)
    if(d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`
  }
  const raw = cleanText(v)
  if(!raw) return ''
  if(/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw
  const m = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/)
  if(m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`
  return raw
}
function today(){ return new Date().toISOString().slice(0, 10) }
function isExpired(date){ return !!date && String(date) < today() }
function generatePassword(){ return `SRGS${Math.random().toString(36).slice(2,8).toUpperCase()}!` }
async function createAuthUser({ email, password, nama, nrp, app_id, site_id, role='Driver' }){
  if(!email || !password) return { ok:true, skipped:true }
  const cleanEmail = normEmail(email)
  if(!isValidEmail(cleanEmail)) throw new Error('Format email login driver tidak valid. Gunakan format nama@domain.com.')
  if(String(password).length < 6) throw new Error('Password minimal 6 karakter.')
  const { data, error } = await supabase.functions.invoke('admin-create-user', { body:{ email:cleanEmail, password, nama, nrp, app_id, role, site_id } })
  if(error){
    let detail = error.message || 'Gagal membuat akun Auth driver'
    try{
      if(error.context && typeof error.context.json === 'function'){
        const body = await error.context.json()
        detail = body?.error || body?.message || detail
      }
    }catch(_){ }
    throw new Error(detail)
  }
  if(data && data.ok === false) throw new Error(data.error || 'Gagal membuat akun Auth driver')
  return data || { ok:true }
}

function templateRows(type){
  const examples = {
    driver: [{ nama_driver:'Budi', nrp_driver:'D-001', email:'budi@company.co.id', password:'password123', vendor_name:'Vendor A', site_code:'BAYA', status:'Aktif' }],
    vendor: [{ vendor_name:'Vendor A', status:'Aktif' }],
    site: [{ site_code:'BAYA', site_name:'BAYA', region:'Operation', status:'Aktif' }],
    plan: [{ site_code:'BAYA', nrp_driver:'D-001', bulan:5, tahun:2026, status:'Planned' }]
  }
  return examples[type] || []
}

const MENUS = {
  sidak_fatigue: [
    ['dashboard', 'Dashboard', LayoutDashboard],
    ['admin', 'Admin Panel', Users],
    ['drivers', 'Master Driver', Truck],
    ['plans', 'Plan Sidak', CalendarCheck],
    ['inspections', 'Inspeksi', ClipboardCheck],
    ['outstanding', 'Outstanding', AlertTriangle],
    ['approval', 'Approval', CheckCircle2]
  ],
  drd_driver: [
    ['dashboard', 'Dashboard DRD', LayoutDashboard],
    ['placeholder', 'DRD Driver', FileText]
  ],
  inspeksi_unit: [
    ['dashboard', 'Dashboard Unit', LayoutDashboard],
    ['placeholder', 'Inspeksi', FileText]
  ]
}

const ADMIN_ROLES = ['Platform Admin','App Admin']
const APPROVAL_ROLES = ['Platform Admin','App Admin','Atasan Site']
function canAdmin(role){ return ADMIN_ROLES.includes(role) }
function canApprove(role){ return APPROVAL_ROLES.includes(role) }
function menuFor(appCode, role){
  const base = MENUS[appCode] || MENUS.sidak_fatigue
  return base.filter(([key]) => {
    if (key === 'admin') return canAdmin(role)
    if (key === 'approval') return canApprove(role)
    return true
  })
}

function App() {
  const [session, setSession] = useState(null)
  const [profile, setProfile] = useState(null)
  const [access, setAccess] = useState([])
  const [context, setContext] = useState(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session)
      setLoading(false)
    })
    const { data: sub } = supabase.auth.onAuthStateChange((_event, nextSession) => setSession(nextSession))
    return () => sub.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (!session?.user) {
      setProfile(null)
      setAccess([])
      setContext(null)
      return
    }
    loadProfileAndAccess(session.user)
  }, [session])

  async function loadProfileAndAccess(user) {
    setError('')

    // 1) Normal lookup: profile sudah punya auth_user_id.
    let { data: prof, error: profErr } = await supabase
      .from('users_profile')
      .select('*')
      .eq('auth_user_id', user.id)
      .maybeSingle()

    if (profErr) {
      setError(profErr.message)
      return
    }

    // 2) Fallback penting: Admin Panel bisa membuat profile+mapping berdasarkan email
    // sebelum auth_user_id tersambung. Saat user login pertama kali, sistem otomatis
    // mencari profile berdasarkan email lalu mengikat auth_user_id ke akun Auth ini.
    if (!prof && user.email) {
      const { data: emailProfile, error: emailErr } = await supabase
        .from('users_profile')
        .select('*')
        .eq('email', user.email)
        .maybeSingle()

      if (emailErr) {
        setError(emailErr.message)
        return
      }

      if (emailProfile) {
        const { data: updatedProfile, error: updateErr } = await supabase
          .from('users_profile')
          .update({ auth_user_id: user.id })
          .eq('id', emailProfile.id)
          .select('*')
          .single()

        if (updateErr) {
          setError(updateErr.message)
          return
        }

        prof = updatedProfile
      }
    }

    if (!prof) {
      setError('Akun Auth sudah ada, tetapi email ini belum dibuat/mapping di Admin Panel. Hubungi Administrator.')
      return
    }

    setProfile(prof)
    const { data: accessRows, error: accessErr } = await supabase
      .from('user_app_access')
      .select('*, applications(app_code, app_name, description), sites(site_name, site_code)')
      .eq('user_id', prof.id)
      .eq('status', 'Aktif')

    if (accessErr) setError(accessErr.message)
    setAccess(accessRows || [])
    if ((accessRows || []).length === 1) setContext(accessRows[0])
  }

  if (loading) return <FullCenter text="Memuat aplikasi..." />
  if (!session) return <Login />
  if (error) return <FullCenter text={error} action={<button onClick={() => supabase.auth.signOut()}>Logout</button>} />
  if (!context) return <ContextPicker profile={profile} access={access} onSelect={setContext} onLogout={() => supabase.auth.signOut()} />
  return <Shell profile={profile} context={context} setContext={setContext} />
}

function Login() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)

  async function submit(e) {
    e.preventDefault()
    setError('')
    setBusy(true)
    const { error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) setError(error.message)
    setBusy(false)
  }

  return <section className="login-page redesign-login real-ui-login">
    <div className="login-hero redesign-hero real-ui-hero">
      <div className="brand brand-on-dark hero-brand">
        <img src={logoSrgs} alt="Platform Inspeksi SRGS" className="brand-logo-img" />
        <div><b>Platform Inspeksi SRGS</b><span>Sistem terintegrasi untuk sidak fatigue, DRD driver, inspeksi, approval, dan monitoring site.</span></div>
      </div>
      <h1>Satu platform untuk pengawasan operasional yang lebih cepat, rapi, dan terukur.</h1>
      <p>Kelola inspeksi lapangan, validasi DRD driver, tindak lanjut temuan, approval berjenjang, serta dashboard monitoring dalam satu sistem terpadu.</p>
      <div className="chips hero-chips"><span>Multi Aplikasi Terintegrasi</span><span>Approval & Tindak Lanjut</span><span>Dashboard & Laporan</span></div>

      <div className="login-visual-stage"><img src={loginScene} alt="Ilustrasi monitoring Platform Inspeksi SRGS" /></div>
      <div className="hero-ui-board compact-board">
        <div className="hero-stat-card wide">
          <div className="hero-stat-icon"><LayoutDashboard size={20} /></div>
          <div>
            <strong>Monitoring Operasional</strong>
            <p>Pantau capaian inspeksi, DRD, outstanding, dan approval dalam satu tampilan.</p>
          </div>
        </div>
        <div className="hero-grid-mini">
          <div className="hero-mini-card"><ClipboardCheck size={18}/><b>Inspeksi</b><span>Checklist, temuan, dan tindak lanjut</span></div>
          <div className="hero-mini-card"><Truck size={18}/><b>DRD Driver</b><span>Assign test, hasil, dan masa berlaku</span></div>
          <div className="hero-mini-card"><CheckCircle2 size={18}/><b>Approval</b><span>Alur verifikasi berjenjang per site</span></div>
          <div className="hero-mini-card"><BarChart3 size={18}/><b>Dashboard</b><span>Ringkasan KPI dan laporan ekspor</span></div>
        </div>
      </div>
    </div>

    <form className="login-card redesign-card real-ui-card" onSubmit={submit}>
      <div className="card-brand-inline"><img src={logoSrgs} alt="Logo Platform Inspeksi SRGS" className="mini-logo" /><div><strong>Platform Inspeksi SRGS</strong><span>Silakan masuk menggunakan akun yang telah didaftarkan administrator.</span></div></div>
      <h2>Selamat datang kembali</h2>
      <p>Masukkan email dan password, lalu lanjutkan ke pemilihan aplikasi, role, dan site kerja sesuai hak akses.</p>
      <label>Email<input required type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="nama@company.co.id" /></label>
      <label>Password<input required type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Masukkan password" /></label>
      {error && <div className="error">{error}</div>}
      <button disabled={busy}>{busy ? 'Memproses...' : 'Masuk ke Platform'}</button>
      <div className="login-support-grid">
        <div><ShieldCheck size={16}/><span>Akses aman</span></div>
        <div><Users size={16}/><span>Role & site terkontrol</span></div>
        <div><FileText size={16}/><span>Riwayat aktivitas tercatat</span></div>
      </div>
      <div className="security-note"><ShieldCheck size={18}/><span>Keamanan data dijaga dan setiap aktivitas tercatat untuk kebutuhan monitoring.</span></div>
    </form>
  </section>
}

function ContextPicker({ profile, access, onSelect, onLogout }) {
  const [appId, setAppId] = useState('')
  const [role, setRole] = useState('')
  const [siteId, setSiteId] = useState('')

  const appOptions = Array.from(new Map(access.map(a => [normalizeAppCode(a.applications?.app_code || a.applications?.app_name), a])).values())
  const roleOptions = Array.from(new Set(access.filter(a => !appId || normalizeAppCode(a.applications?.app_code || a.applications?.app_name) === appId).map(a => a.role)))
  const siteOptions = Array.from(new Map(access.filter(a => (!appId || normalizeAppCode(a.applications?.app_code || a.applications?.app_name) === appId) && (!role || a.role === role)).map(a => [a.site_id || 'all', a])).values())

  useEffect(() => {
    if (access.length) {
      setAppId(normalizeAppCode(access[0].applications?.app_code || access[0].applications?.app_name) || '')
      setRole(access[0].role || '')
      setSiteId(access[0].site_id || '')
    }
  }, [access])

  useEffect(() => {
    if (!access.length || !appId) return
    const candidates = access.filter(a => normalizeAppCode(a.applications?.app_code || a.applications?.app_name) === appId)
    const first = candidates[0]
    if (first) { setRole(first.role || ''); setSiteId(first.site_id || '') }
  }, [appId])

  useEffect(() => {
    if (!access.length || !appId || !role) return
    const candidates = access.filter(a => normalizeAppCode(a.applications?.app_code || a.applications?.app_name) === appId && a.role === role)
    const first = candidates[0]
    if (first) setSiteId(first.site_id || '')
  }, [role])

  function startSession(){
    const selected = access.find(a => String(normalizeAppCode(a.applications?.app_code || a.applications?.app_name)) === String(appId) && String(a.role) === String(role) && String(a.site_id || '') === String(siteId || ''))
    if (!selected) return alert('Akses tidak ditemukan. Cek mapping aplikasi, role, dan site di Supabase.')
    onSelect(selected)
  }

  return <section className="context-page redesign-context-page real-ui-context-page">
    <div className="context-card redesign-context-card real-ui-context-card">
      <div className="context-visual-panel real-ui-context-visual">
        <div className="brand hero-brand"><img src={logoSrgs} alt="Platform Inspeksi SRGS" className="brand-logo-img" /><div><b>Platform Inspeksi SRGS</b><span>Multi Application Inspection System</span></div></div>

        <div className="context-showcase context-showcase-image">
          <img src={roleScene} alt="Ilustrasi akses aplikasi, role, dan site" />
          <div className="showcase-window compact-showcase-window">
            <div className="showcase-topline"><span></span><span></span><span></span></div>
            <div className="showcase-body">
              <div className="showcase-side">
                <div className="showcase-side-item active"><LayoutDashboard size={16}/> Dashboard</div>
                <div className="showcase-side-item"><Truck size={16}/> DRD Driver</div>
                <div className="showcase-side-item"><ClipboardCheck size={16}/> Inspeksi</div>
                <div className="showcase-side-item"><ShieldCheck size={16}/> Approval</div>
              </div>
              <div className="showcase-main">
                <div className="showcase-kpi-row">
                  <div className="showcase-kpi"><b>12</b><span>Plan Aktif</span></div>
                  <div className="showcase-kpi"><b>8</b><span>Approved</span></div>
                  <div className="showcase-kpi"><b>2</b><span>Outstanding</span></div>
                </div>
                <div className="showcase-form">
                  <div className="showcase-field"><span>Aplikasi</span><b>Sidak Fatigue</b></div>
                  <div className="showcase-field"><span>Role</span><b>Platform Admin</b></div>
                  <div className="showcase-field"><span>Site</span><b>JIEP / Head Office</b></div>
                </div>
                <div className="showcase-action">Masuk ke Aplikasi</div>
              </div>
            </div>
          </div>
        </div>

        <div className="context-benefits">
          <div><b>Akses Terkontrol</b><span>Hak akses mengikuti role dan site kerja Anda.</span></div>
          <div><b>Multi Aplikasi</b><span>Berpindah aplikasi tanpa perlu login ulang.</span></div>
          <div><b>Responsif</b><span>Tampilan tetap nyaman saat diakses melalui laptop maupun HP.</span></div>
        </div>
      </div>
      <div className="context-form-panel real-ui-context-form">
        <div className="picker-head-icon"><Users size={24} /></div>
        <h1>Pilih aplikasi, role, dan site</h1>
        <p>Silakan pilih aplikasi, hak akses, dan site kerja yang sesuai dengan akun Anda.</p>
        <div className="picker-user-box"><div><b>User: {profile?.nama}</b><span>{profile?.email}</span></div><span className="pill green">Akun Terverifikasi</span></div>
        {access.length === 0 ? <div className="empty">
          <AlertTriangle />
          <div><b>Akses belum dimapping.</b><p>Administrator perlu menambahkan user ini ke tabel user_app_access agar dapat digunakan.</p></div>
        </div> : <div className="context-select-grid redesigned-picker-grid">
          <label>Aplikasi<select value={appId} onChange={e=>setAppId(e.target.value)}>{appOptions.map(a => { const code = normalizeAppCode(a.applications?.app_code || a.applications?.app_name); const name = code === 'inspeksi_unit' ? 'Inspeksi' : (a.applications?.app_name || code); return <option key={code} value={code}>{name}</option> })}</select></label>
          <label>Role<select value={role} onChange={e=>setRole(e.target.value)}>{roleOptions.map(r=><option key={r}>{r}</option>)}</select></label>
          <label>Site<select value={siteId} onChange={e=>setSiteId(e.target.value)}>{siteOptions.map(a=><option key={a.site_id || 'all'} value={a.site_id || ''}>{a.sites?.site_name || 'All Site'}</option>)}</select></label>
          <button onClick={startSession}>Masuk ke Aplikasi</button>
        </div>}
        <button className="secondary" onClick={onLogout}>Logout</button>
        <div className="security-note light"><ShieldCheck size={18}/><span>Seluruh aktivitas penggunaan aplikasi tercatat untuk monitoring dan audit internal.</span></div>
      </div>
    </div>
  </section>
}

function Shell({ profile, context, setContext }) {
  const [page, setPage] = useState('dashboard')
  const [sidebar, setSidebar] = useState(false)
  const appCode = normalizeAppCode(context.applications?.app_code || context.applications?.app_name)
  const menu = menuFor(appCode, context.role)
  const appName = appCode === 'inspeksi_unit' ? 'Inspeksi' : (context.applications?.app_name || 'Aplikasi')
  const pageTitle = menu.find(m => m[0] === page)?.[1] || 'Dashboard'
  const pageDescriptions = {
    dashboard: 'Ringkasan aktivitas, pencapaian, approval, dan tindak lanjut seluruh site.',
    admin: 'Kelola user, mapping akses, dan pengaturan aplikasi secara terpusat.',
    drivers: 'Master data driver untuk kebutuhan sidak fatigue dan monitoring.',
    plans: 'Rencanakan objek inspeksi bulanan agar pelaksanaan lebih tertata.',
    inspections: 'Laksanakan inspeksi berdasarkan plan dan catat hasil pemeriksaan.',
    outstanding: 'Pantau temuan aktif dan progres tindak lanjut lapangan.',
    approval: 'Verifikasi hasil inspeksi dan tindak lanjut secara berjenjang.'
  }

  useEffect(() => setPage('dashboard'), [context.id])

  if (appCode === 'drd_driver') return <DrdDriverApp embeddedProfile={profile} embeddedWork={context} onChangeApp={() => setContext(null)} />
  if (appCode === 'inspeksi_unit') return <InspeksiUnitApp embeddedProfile={profile} embeddedContext={context} onChangeApp={() => setContext(null)} />
  if (appCode === 'food_index') return <FoodIndexApp embeddedProfile={profile} embeddedContext={context} onChangeApp={() => setContext(null)} />

  return <div className="app-shell redesign-shell real-ui-shell">
    {sidebar && <button className="mobile-nav-backdrop" aria-label="Tutup navigasi" onClick={()=>setSidebar(false)} />}
    <aside className={sidebar ? 'sidebar open' : 'sidebar'}>
      <button className="sidebar-close" onClick={()=>setSidebar(false)}>×</button>
      <div className="brand dark sidebar-brand"><img src={logoSrgs} alt="Platform Inspeksi SRGS" className="brand-logo-img" /><div><b>Platform Inspeksi SRGS</b><span>{appName} · {context.role}</span></div></div>
      <nav>{menu.map(([key,label,Icon]) => <button key={key} className={page===key?'active':''} onClick={()=>{setPage(key); setSidebar(false)}}><Icon size={18}/>{label}</button>)}</nav>
      <div className="sidebar-info-cards">
        <div className="sidebar-info-card"><b>{context.sites?.site_name || 'All Site'}</b><span>Lokasi kerja aktif</span></div>
        <div className="sidebar-info-card"><b>{appName}</b><span>Aplikasi yang sedang digunakan</span></div>
      </div>
      <div className="sidebar-card"><b>Pusat Monitoring</b><p>Gunakan menu di samping untuk memantau aktivitas, approval, dan tindak lanjut sesuai aplikasi yang dipilih.</p></div>
    </aside>
    <main className="main redesign-main">
      <header className="app-header">
        <button className="icon mobile" onClick={()=>setSidebar(!sidebar)}><Menu /></button>
        <div className="header-copy"><h2>{pageTitle}</h2><p>{pageDescriptions[page] || 'Kelola aktivitas operasional sesuai kebutuhan aplikasi.'}</p></div>
        <div className="header-actions redesign-actions"><div className="user-chip"><Users size={16}/><span>{profile?.nama} · {context.sites?.site_name || 'All Site'}</span></div><button className="secondary" onClick={()=>setContext(null)}>Ganti Aplikasi</button><button className="secondary" onClick={()=>supabase.auth.signOut()}><LogOut size={16}/> Logout</button></div>
      </header>
      <section className="content">
        <SidakPage page={page} context={context} profile={profile} />
      </section>
    </main>
  </div>
}

function SidakPage({ page, context, profile }) {
  if (page === 'dashboard') return <SidakDashboard context={context} />
  if (page === 'admin') return <AdminPanel context={context} profile={profile} />
  if (page === 'drivers') return <DriverMaster context={context} />
  if (page === 'plans') return <FatiguePlans context={context} profile={profile} />
  if (page === 'inspections') return <FatigueInspections context={context} profile={profile} />
  if (page === 'outstanding') return <FatigueOutstanding context={context} profile={profile} />
  if (page === 'approval') return <FatigueApproval context={context} profile={profile} />
  return null
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

function SidakDashboard({ context }) {
  const [plans, setPlans] = useState([])
  const [inspections, setInspections] = useState([])
  const [outs, setOuts] = useState([])
  const [loading, setLoading] = useState(true)
  const [dateFrom, setDateFrom] = useState('')
  const [dateTo, setDateTo] = useState('')
  const globalView = isHeadOfficeAdmin(context)

  useEffect(() => { load() }, [context.id, dateFrom, dateTo])
  async function load() {
    setLoading(true)
    let planQ = supabase.from('fatigue_plans').select('id, bulan, tahun, status, created_at, drivers(nama_driver,nrp_driver), sites(site_name,site_code)').order('created_at', { ascending: false })
    let inspQ = supabase.from('fatigue_inspections').select('id, tanggal_inspeksi, status, drivers(nama_driver,nrp_driver), sites(site_name,site_code)').order('created_at', { ascending: false })
    let outQ = supabase.from('fatigue_outstandings').select('id, description, status, created_at, drivers(nama_driver,nrp_driver), sites(site_name,site_code), fatigue_parameters(parameter_name)').order('created_at', { ascending: false })
    if (!globalView) { planQ = planQ.eq('site_id', context.site_id); inspQ = inspQ.eq('site_id', context.site_id); outQ = outQ.eq('site_id', context.site_id) }
    if (dateFrom) { planQ = planQ.gte('created_at', dateFrom); inspQ = inspQ.gte('tanggal_inspeksi', dateFrom); outQ = outQ.gte('created_at', dateFrom) }
    if (dateTo) { const end = dateTo + 'T23:59:59'; planQ = planQ.lte('created_at', end); inspQ = inspQ.lte('tanggal_inspeksi', dateTo); outQ = outQ.lte('created_at', end) }
    const [{ data: p }, { data: i }, { data: o }] = await Promise.all([planQ, inspQ, outQ])
    // JIEP adalah Head Office: admin JIEP melihat all site operation.
    // Data JIEP sendiri tetap tidak masuk KPI/chart operasional.
    setPlans((p || []).filter(x => x.sites?.site_code !== 'JIEP'))
    setInspections((i || []).filter(x => x.sites?.site_code !== 'JIEP'))
    setOuts((o || []).filter(x => x.sites?.site_code !== 'JIEP'))
    setLoading(false)
  }

  const rows = plans.map(r => ({ id: r.id, bulan: r.bulan, tahun: r.tahun, status: r.status, driver: r.drivers?.nama_driver, nrp: r.drivers?.nrp_driver, site: r.sites?.site_name }))
  const actual = inspections.filter(x => x.status === 'Approved').length
  const open = outs.filter(x => x.status === 'Open').length
  const closed = outs.filter(x => x.status === 'Approved').length
  const achievement = plans.length ? Math.round((actual / plans.length) * 100) : 0

  const siteMap = new Map()
  plans.forEach(p => {
    const code = p.sites?.site_code || '-'
    if (!siteMap.has(code)) siteMap.set(code, { site: p.sites?.site_name || code, plan: 0, actual: 0 })
    siteMap.get(code).plan++
  })
  inspections.filter(i => i.status === 'Approved').forEach(i => {
    const code = i.sites?.site_code || '-'
    if (!siteMap.has(code)) siteMap.set(code, { site: i.sites?.site_name || code, plan: 0, actual: 0 })
    siteMap.get(code).actual++
  })
  const siteRows = Array.from(siteMap.values()).map(x => ({ ...x, achievement: x.plan ? Math.round((x.actual/x.plan)*100) : 0 })).sort((a,b)=>a.site.localeCompare(b.site))
  const temuanRows = outs.map(o => ({
    id: o.id,
    site: o.sites?.site_name || '-',
    driver: o.drivers?.nama_driver || '-',
    nrp: o.drivers?.nrp_driver || '-',
    parameter: o.fatigue_parameters?.parameter_name || 'Temuan',
    deskripsi: o.description || '-',
    status: o.status
  }))
  const openRows = temuanRows.filter(r => r.status === 'Open')
  const closePendingRows = temuanRows.filter(r => r.status === 'Closed')
  const closeApprovedRows = temuanRows.filter(r => r.status === 'Approved')

  return <div className="stack">
    <DashboardDateFilter dateFrom={dateFrom} dateTo={dateTo} setDateFrom={setDateFrom} setDateTo={setDateTo} onClear={()=>{setDateFrom('');setDateTo('')}} />
    <div className="kpi-grid">
      <Kpi title="Plan Bulanan" value={plans.length} icon={<CalendarCheck/>} />
      <Kpi title="Aktual Approved" value={actual} icon={<ClipboardCheck/>} />
      <Kpi title="Achievement" value={`${achievement}%`} icon={<BarChart3/>} />
      <Kpi title="Outstanding Open" value={open} icon={<AlertTriangle/>} />
    </div>
    <div className="grid-2">
      <Panel title="Row Data Plan" desc="Data detail dashboard, siap export CSV/Excel. JIEP tidak masuk hitungan dashboard." action={<button onClick={()=>exportToCsv('sidak-fatigue-plan.csv', rows)}><Download size={16}/> Export CSV</button>}>
        {loading ? <p>Memuat...</p> : <DataTable rows={rows}/>} 
      </Panel>
      <Panel title="Temuan Terbaru" desc={`${closed} approved close`}>
        {outs.slice(0,5).map(o => <div className="mini-card" key={o.id}><b>{o.fatigue_parameters?.parameter_name || 'Temuan'}</b><StatusPill value={o.status}/><p>{o.description}</p><small>{o.drivers?.nama_driver}</small></div>)}
        {outs.length === 0 && <p className="muted">Belum ada temuan.</p>}
      </Panel>
    </div>
    <Panel title="Row Data Temuan Open & Closed" desc="Daftar seluruh temuan open, close pending approval, dan closed approved. Data ini bisa diekspor untuk laporan tindak lanjut." action={<button onClick={()=>downloadXlsx('row-data-temuan-sidak.xlsx', temuanRows)}><Download size={16}/> Export Excel</button>}>
      <div className="summary-strip">
        <span><b>{openRows.length}</b> Open</span>
        <span><b>{closePendingRows.length}</b> Close Pending Approval</span>
        <span><b>{closeApprovedRows.length}</b> Closed Approved</span>
      </div>
      {temuanRows.length ? <DataTable rows={temuanRows}/> : <p className="muted">Belum ada row data temuan open/closed.</p>}
    </Panel>
    <Panel title="Achievement All Site" desc="Pencapaian inspeksi approved per site. Site JIEP/Head Office dikecualikan dari chart dan hitungan." action={<button onClick={()=>downloadXlsx('achievement-all-site.xlsx', siteRows)}><Download size={16}/> Export Excel</button>}>
      <div className="site-chart">
        {siteRows.map(r => <div className="site-bar" key={r.site}>
          <div className="site-meta"><b>{r.site}</b><span>{r.actual}/{r.plan} · {r.achievement}%</span></div>
          <div className="bar"><span style={{width:`${Math.min(r.achievement,100)}%`}} /></div>
        </div>)}
      </div>
      <DataTable rows={siteRows}/>
    </Panel>
  </div>
}
function DriverMaster({ context }) {
  const [drivers, setDrivers] = useState([])
  const [vendors, setVendors] = useState([])
  const [sites, setSites] = useState([])
  const [preview, setPreview] = useState([])
  const [message, setMessage] = useState('')
  const [editingId, setEditingId] = useState(null)
  const [modalOpen, setModalOpen] = useState(false)
  const [driverImporting, setDriverImporting] = useState(false)
  const [driverImportProgress, setDriverImportProgress] = useState({ current: 0, total: 0 })
  const adminHO = isHeadOfficeAdmin(context)
  const emptyForm = { site_id: context.site_id || '', nama_driver: '', nrp_driver: '', email: '', password: '', vendor_id: '', status: 'Aktif', mulai_dinas: '', end_masa_dinas: '' }
  const [form, setForm] = useState(emptyForm)

  useEffect(() => { load() }, [context.id])

  async function load() {
    let driverQ = supabase.from('drivers').select('*, vendors(vendor_code,vendor_name), sites(site_name,site_code)').order('created_at', { ascending: false })
    if (!adminHO) driverQ = driverQ.eq('site_id', context.site_id)
    const [{ data: d }, { data: v }, { data: s }] = await Promise.all([
      driverQ,
      supabase.from('vendors').select('*').order('vendor_name'),
      supabase.from('sites').select('*').neq('site_code','JIEP').order('site_code')
    ])
    setDrivers(d || [])
    setVendors(v || [])
    setSites(s || [])
    setForm(f => ({ ...f, site_id: f.site_id || context.site_id || (s || [])[0]?.id || '' }))
  }

  function resetForm(){
    setEditingId(null)
    setModalOpen(false)
    setForm({ ...emptyForm, site_id: context.site_id || '' })
  }

  function editDriver(d){
    setEditingId(d.id)
    setForm({
      site_id: d.site_id || context.site_id || '',
      nama_driver: d.nama_driver || '',
      nrp_driver: d.nrp_driver || '',
      email: d.email || '',
      password: '',
      vendor_id: d.vendor_id || '',
      status: d.status || 'Aktif',
      mulai_dinas: d.mulai_dinas || '',
      end_masa_dinas: d.end_masa_dinas || ''
    })
    setModalOpen(true)
  }

  function duplicateMessage(dup, nrpKey, emailKey){
    if (dup.nrp_driver && cleanText(dup.nrp_driver).toLowerCase() === nrpKey) return 'NRP driver sudah terdaftar.'
    if (emailKey && cleanText(dup.email).toLowerCase() === emailKey) return 'Email driver sudah terdaftar.'
    return 'Data driver sudah terdaftar.'
  }

  async function save(e) {
    e.preventDefault()
    setMessage('')
    const siteId = adminHO ? form.site_id : context.site_id
    const nrpKey = cleanText(form.nrp_driver).toLowerCase()
    const emailKey = normEmail(form.email)
    const dup = drivers.find(d => String(d.id) !== String(editingId || '') && (
      cleanText(d.nrp_driver).toLowerCase() === nrpKey ||
      (emailKey && normEmail(d.email) === emailKey)
    ))
    if (dup) return setMessage(duplicateMessage(dup, nrpKey, emailKey))
    if (form.email && !isValidEmail(form.email)) return setMessage('Format email driver tidak valid. Gunakan format nama@domain.com.')
    if (form.email && !form.password && !editingId) return setMessage('Password wajib diisi jika membuat akun login driver.')
    if (form.password && String(form.password).length < 6) return setMessage('Password minimal 6 karakter.')

    const payload = {
      site_id: siteId,
      nama_driver: cleanText(form.nama_driver),
      nrp_driver: cleanText(form.nrp_driver),
      email: normEmail(form.email) || null,
      vendor_id: form.vendor_id || null,
      status: form.status || 'Aktif',
      mulai_dinas: form.mulai_dinas || null,
      end_masa_dinas: form.end_masa_dinas || null,
      updated_at: new Date().toISOString()
    }
    try {
      if (form.email && form.password) await createAuthUser({ email: form.email, password: form.password, nama: form.nama_driver, nrp: form.nrp_driver, app_id: context.app_id || context.applications?.id, site_id: siteId, role: 'Driver' })
      if (editingId) {
        const { error } = await supabase.from('drivers').update(payload).eq('id', editingId)
        if (error) throw error
        setMessage(form.password ? 'Driver dan akun login berhasil diupdate.' : 'Driver berhasil diupdate.')
      } else {
        const { error } = await supabase.from('drivers').upsert(payload, { onConflict: 'nrp_driver' })
        if (error) throw error
        setMessage(form.password ? 'Driver dan akun login berhasil ditambahkan.' : 'Driver berhasil ditambahkan.')
      }
      resetForm(); load()
    } catch (err) { setMessage(err.message || 'Gagal menyimpan driver.') }
  }

  async function toggleDriver(d){
    const next = d.status === 'Aktif' ? 'Nonaktif' : 'Aktif'
    const { error } = await supabase.from('drivers').update({ status: next, updated_at: new Date().toISOString() }).eq('id', d.id)
    if (error) setMessage(error.message); else { setMessage(`Driver ${next}.`); load() }
  }
  async function deleteDriver(d){
    if(!confirm('Delete/nonaktifkan driver '+d.nama_driver+'?')) return
    const { error } = await supabase.from('drivers').update({ status:'Nonaktif', updated_at: new Date().toISOString() }).eq('id', d.id)
    setMessage(error ? error.message : 'Driver dinonaktifkan.'); load()
  }

  async function previewDrivers(file){
    if(!file) return
    setMessage('')
    try{
      const rows = (await parseExcelOrCsv(file)).map(normalizeRow)
      const seenNrp=new Set(), seenEmail=new Set()
      const [{data:allSites},{data:allVendors}] = await Promise.all([supabase.from('sites').select('id,site_code'), supabase.from('vendors').select('id,vendor_name')])
      const mapped = rows.map((r,idx)=>{
        const siteCode = adminHO ? cleanText(r.site_code).toUpperCase() : context.sites?.site_code
        const vendorName = cleanText(r.vendor_name || r.vendor)
        const site=(allSites||[]).find(s=>s.site_code===siteCode)
        const vendor= vendorName ? (allVendors||[]).find(v=>cleanText(v.vendor_name).toLowerCase()===vendorName.toLowerCase()) : null
        const nrpKey=cleanText(r.nrp_driver).toLowerCase(), emailKey=normEmail(r.email)
        let error = ''
        if(!site) error = 'site_code tidak ditemukan'
        else if(!cleanText(r.nama_driver) || !cleanText(r.nrp_driver)) error = 'nama_driver dan nrp_driver wajib'
        else if(emailKey && !isValidEmail(emailKey)) error = 'email tidak valid'
        else if(seenNrp.has(nrpKey)) error='nrp_driver double di file import'
        else if(emailKey && seenEmail.has(emailKey)) error='email double di file import'
        else if(drivers.some(d=>cleanText(d.nrp_driver).toLowerCase()===nrpKey)) error='nrp_driver sudah terdaftar'
        else if(emailKey && drivers.some(d=>normEmail(d.email)===emailKey)) error='email sudah terdaftar'
        seenNrp.add(nrpKey); if(emailKey) seenEmail.add(emailKey)
        return { row:idx+2, site_code:siteCode, nama_driver:cleanText(r.nama_driver), nrp_driver:cleanText(r.nrp_driver), email:emailKey, password:cleanText(r.password), vendor_name:vendor?.vendor_name||vendorName, mulai_dinas:excelDateToIso(r.mulai_dinas), end_masa_dinas:excelDateToIso(r.end_masa_dinas), status:cleanText(r.status)||'Aktif', site_id:site?.id, vendor_id:vendor?.id||null, error }
      })
      setPreview(mapped)
    }catch(e){ setMessage(e.message) }
  }

  async function submitDriverImport(){
    setMessage('')
    const valid = preview.filter(r=>!r.error)
    if(!valid.length) return setMessage('Tidak ada baris valid untuk diimport.')
    if(valid.length !== preview.length) return setMessage('Masih ada baris invalid.')

    const errors = []
    let successCount = 0
    setDriverImporting(true)
    setDriverImportProgress({ current: 0, total: valid.length })
    try {
      for (const [idx, r] of valid.entries()) {
        try {
          let vendorId = r.vendor_id || null
          if (!vendorId && r.vendor_name) {
            const { data: existingVendor, error: existingVendorErr } = await supabase.from('vendors').select('id').ilike('vendor_name', r.vendor_name).maybeSingle()
            if (existingVendorErr) throw existingVendorErr
            if (existingVendor?.id) vendorId = existingVendor.id
            else {
              const { data: newVendor, error: vendorErr } = await supabase.from('vendors').insert({ vendor_code: makeVendorCode(idx), vendor_name: r.vendor_name, status: 'Aktif' }).select('id').single()
              if (vendorErr) throw vendorErr
              vendorId = newVendor.id
            }
          }
          const { error: driverErr } = await supabase.from('drivers').upsert({
            nama_driver:r.nama_driver,
            nrp_driver:r.nrp_driver,
            email:r.email||null,
            site_id:r.site_id,
            vendor_id:vendorId,
            status:r.status,
            mulai_dinas:r.mulai_dinas||null,
            end_masa_dinas:r.end_masa_dinas||null,
            updated_at:new Date().toISOString()
          }, { onConflict:'nrp_driver' })
          if (driverErr) throw driverErr
          successCount += 1
          if (r.email && r.password) {
            try {
              await createAuthUser({ email:r.email, password:r.password, nama:r.nama_driver, nrp:r.nrp_driver, app_id:context.app_id || context.applications?.id, site_id:r.site_id, role:'Driver' })
            } catch (authErr) {
              errors.push(`Baris ${r.row}: data driver tersimpan, tapi akun login gagal dibuat (${authErr.message || String(authErr)})`)
            }
          }
        } catch (e) {
          errors.push(`Baris ${r.row}: ${e.message || String(e)}`)
        } finally {
          setDriverImportProgress({ current: idx + 1, total: valid.length })
        }
      }

      if(errors.length){
        setMessage(`Import selesai: ${successCount} driver tersimpan. Catatan akun login: ${errors.slice(0,3).join(' | ')}`)
        setPreview([]); loadAll(); return
      }
      setMessage(`Import driver berhasil: ${successCount} baris. Kode vendor dibuat otomatis jika ada vendor baru.`)
      setPreview([])
      load()
    } finally {
      setDriverImporting(false)
      setDriverImportProgress({ current: 0, total: 0 })
    }
  }

  const rows = drivers.map(d => ({
    nama: d.nama_driver,
    nrp: d.nrp_driver,
    email: d.email || '-',
    site: d.sites?.site_code || d.sites?.site_name || '-',
    vendor: d.vendors?.vendor_name || '-',
    mulai_dinas: d.mulai_dinas || '-',
    end_masa_dinas: d.end_masa_dinas || '-',
    status_masa_dinas: isExpired(d.end_masa_dinas) ? 'Masa Dinas Habis' : 'Aktif',
    status: d.status,
    aksi: ''
  }))

  const renderDriverForm = ({ submitLabel, cancelLabel }) => <form className="form-grid" onSubmit={save} autoComplete="off">
    {adminHO && <label>Site<select required value={form.site_id} onChange={e=>setForm({...form,site_id:e.target.value})}><option value="">Pilih site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>}
    <label>Nama Driver<input required autoComplete="off" value={form.nama_driver} onChange={e=>setForm({...form,nama_driver:e.target.value})}/></label>
    <label>NRP<input required autoComplete="off" value={form.nrp_driver} onChange={e=>setForm({...form,nrp_driver:e.target.value})}/></label>
    <label>Email<input type="email" name="driver_email_no_autofill" autoComplete="new-email" value={form.email} onChange={e=>setForm({...form,email:e.target.value})}/></label>
    <label>Password<input type="password" name="driver_password_no_autofill" autoComplete="new-password" value={form.password} onChange={e=>setForm({...form,password:e.target.value})} placeholder={editingId?'Kosongkan jika tidak ubah akun':'Password awal'}/></label>
    <label>Vendor<select value={form.vendor_id} onChange={e=>setForm({...form,vendor_id:e.target.value})}><option value="">Tanpa vendor</option>{vendors.map(v=><option key={v.id} value={v.id}>{v.vendor_name}</option>)}</select></label>
    <label>Mulai Dinas<input type="date" value={form.mulai_dinas} onChange={e=>setForm({...form,mulai_dinas:e.target.value})}/></label>
    <label>End Masa Dinas<input type="date" value={form.end_masa_dinas} onChange={e=>setForm({...form,end_masa_dinas:e.target.value})}/></label>
    <label>Status<select value={form.status} onChange={e=>setForm({...form,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label>
    <button>{submitLabel}</button>{cancelLabel && <button type="button" className="secondary" onClick={resetForm}>{cancelLabel}</button>}
  </form>

  return <div className="stack">
    <Panel title="Master Driver" desc="Tambah driver baru dan masa dinas. Untuk mengubah data existing, klik Edit pada tabel agar form terbuka dalam modal.">
      {!editingId && renderDriverForm({ submitLabel: 'Simpan Driver + Akun' })}
      {editingId && <p className="message">Sedang mengedit driver di modal. Tutup modal untuk kembali tambah driver baru.</p>}
      {message && <p className="message">{message}</p>}
    </Panel>
    {modalOpen && <div className="modal-backdrop" onClick={resetForm}>
      <div className="modal-card" onClick={e=>e.stopPropagation()}>
        <div className="modal-head">
          <div><h2>Edit Master Driver</h2><p className="muted">Ubah data driver di sini agar tidak tercampur dengan form tambah driver.</p></div>
          <button type="button" className="secondary small" onClick={resetForm}>Tutup</button>
        </div>
        {renderDriverForm({ submitLabel: 'Update Driver', cancelLabel: 'Batal Edit' })}
        {message && <p className="message">{message}</p>}
      </div>
    </div>}
    <Panel title="Import Driver Excel" desc="Template tidak memakai kode unik. Cukup isi site_code, nama_driver, nrp_driver, email/password bila perlu akun, vendor_name bila ada, serta tanggal masa dinas.">
      <style>{'@keyframes srgsSpin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}'}</style>
      <div className="import-actions"><button className="secondary" disabled={driverImporting} onClick={()=>downloadTemplate('template-master-driver.xlsx', [{site_code:'BAYA',nama_driver:'Budi',nrp_driver:'D-001',email:'budi@company.co.id',password:'password123',vendor_name:'PBM',mulai_dinas:'2026-04-01',end_masa_dinas:'2026-05-01',status:'Aktif'}])}><Download size={16}/> Download Template Excel</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" disabled={driverImporting} onChange={e=>previewDrivers(e.target.files?.[0])}/></label>{preview.length>0 && <button onClick={submitDriverImport} disabled={driverImporting || preview.some(r=>r.error)}>{driverImporting ? 'Mengimpor...' : 'Submit Import Valid'}</button>}</div>
      {driverImporting && <div className="message" style={{display:'flex',alignItems:'center',gap:10}}><span aria-hidden="true" style={{width:16,height:16,border:'2px solid #bfdbfe',borderTopColor:'#2563eb',borderRadius:'50%',display:'inline-block',animation:'srgsSpin 0.8s linear infinite'}}/><span>Import sedang diproses... {driverImportProgress.current}/{driverImportProgress.total} baris. Mohon tunggu untuk file besar.</span></div>}
      {preview.length>0 && <PreviewTable rows={preview}/>} 
    </Panel>
    <Panel title="Row Data Master Driver" desc="Aksi edit/delete dibuat rapi dalam satu tabel. Delete akan mengubah status menjadi Nonaktif agar histori tetap aman." action={<button onClick={()=>downloadXlsx('master-driver.xlsx', rows.map(({aksi,...r})=>r))}><Download size={16}/> Export Excel</button>}>
      <DataTable rows={rows} customActions={(idx)=>{const d=drivers[idx]; return <div className="row-actions"><button className="secondary small" onClick={()=>editDriver(d)}>Edit</button><button className="secondary small" onClick={()=>toggleDriver(d)}>{d.status==='Aktif'?'Nonaktifkan':'Aktifkan'}</button><button className="danger small" onClick={()=>deleteDriver(d)}>Delete</button></div>}} />
    </Panel>
  </div>
}

function FatiguePlans({ context, profile }) {
  const [drivers, setDrivers] = useState([])
  const [plans, setPlans] = useState([])
  const [sites, setSites] = useState([])
  const [form, setForm] = useState({ driver_id: '', site_id: context.site_id, bulan: String(new Date().getMonth()+1), tahun: String(new Date().getFullYear()) })
  const [message, setMessage] = useState('')
  const [preview, setPreview] = useState([])
  const adminHO = isHeadOfficeAdmin(context)

  useEffect(() => { load() }, [context.id])
  async function load() {
    let driverQ = supabase.from('drivers').select('*, sites(site_code,site_name)').eq('status', 'Aktif').order('nama_driver')
    let planQ = supabase.from('fatigue_plans').select('*, drivers(nama_driver,nrp_driver), sites(site_name,site_code)').order('created_at', { ascending:false })
    if (!adminHO) { driverQ = driverQ.eq('site_id', context.site_id); planQ = planQ.eq('site_id', context.site_id) }
    const [{ data: d }, { data: p }, { data: s }] = await Promise.all([
      driverQ,
      planQ,
      supabase.from('sites').select('*').neq('site_code','JIEP').order('site_code')
    ])
    setDrivers(d || [])
    setPlans(p || [])
    setSites(s || [])
  }
  const visibleDrivers = adminHO ? drivers.filter(d => !form.site_id || d.site_id === form.site_id) : drivers
  async function save(e) {
    e.preventDefault(); setMessage('')
    const siteId = adminHO ? form.site_id : context.site_id
    const { error } = await supabase.from('fatigue_plans').insert({
      site_id: siteId,
      driver_id: form.driver_id,
      bulan: Number(form.bulan),
      tahun: Number(form.tahun),
      created_by: profile.id,
      status: 'Planned'
    })
    if (error) setMessage(error.message)
    else { setMessage('Plan berhasil dibuat.'); setForm({...form, driver_id:''}); load() }
  }
  async function previewPlans(file){
    if (!file) return
    setMessage('')
    try {
      const rows = (await parseExcelOrCsv(file)).map(normalizeRow)
      const [{ data: allSites }, { data: allDrivers }] = await Promise.all([
        supabase.from('sites').select('id,site_code'),
        supabase.from('drivers').select('id,nrp_driver,site_id')
      ])
      const mapped = rows.map((r,idx)=>{
        const siteCode = adminHO ? cleanText(r.site_code).toUpperCase() : context.sites?.site_code
        const nrp = cleanText(r.nrp_driver)
        const site = (allSites || []).find(s => s.site_code === siteCode)
        const driver = (allDrivers || []).find(d => d.nrp_driver === nrp && (!site || d.site_id === site.id))
        let error=''
        const bulan=Number(r.bulan), tahun=Number(r.tahun)
        if(!site) error='site_code tidak ditemukan'
        else if(!driver) error='nrp_driver tidak ditemukan di site tersebut'
        else if(!bulan || bulan<1 || bulan>12) error='bulan wajib 1-12'
        else if(!tahun) error='tahun wajib'
        return { row:idx+2, site_code:siteCode, nrp_driver:nrp, bulan, tahun, status:cleanText(r.status)||'Planned', site_id:site?.id, driver_id:driver?.id, error }
      })
      setPreview(mapped)
    }catch(e){ setMessage(e.message) }
  }
  async function submitPlanImport(){
    const valid=preview.filter(r=>!r.error)
    if(!valid.length) return setMessage('Tidak ada baris valid untuk diimport.')
    const payload=valid.map(r=>({ site_id:r.site_id, driver_id:r.driver_id, bulan:Number(r.bulan), tahun:Number(r.tahun), status:r.status || 'Planned', created_by: profile.id }))
    const { error } = await supabase.from('fatigue_plans').insert(payload)
    if(error) setMessage(error.message)
    else { setMessage(`Import plan berhasil: ${payload.length} baris.`); setPreview([]); load() }
  }
  const rows = plans.map(p => ({ id:p.id, bulan:p.bulan, tahun:p.tahun, driver:p.drivers?.nama_driver, nrp:p.drivers?.nrp_driver, site:p.sites?.site_name, status:p.status }))
  return <div className="stack">
    <Panel title="Buat Plan Sidak Bulanan" desc="Planning berbasis objek driver dalam bulan berjalan, bukan per tanggal.">
      <form className="form-grid" onSubmit={save}>
        {adminHO && <label>Site<select required value={form.site_id} onChange={e=>setForm({...form, site_id:e.target.value, driver_id:''})}>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>}
        <label>Driver<select required value={form.driver_id} onChange={e=>setForm({...form, driver_id:e.target.value})}><option value="">Pilih driver</option>{visibleDrivers.map(d=><option key={d.id} value={d.id}>{d.nama_driver} - {d.nrp_driver} {adminHO?`(${d.sites?.site_code})`:''}</option>)}</select></label>
        <label>Bulan<select value={form.bulan} onChange={e=>setForm({...form, bulan:e.target.value})}>{MONTHS.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select></label>
        <label>Tahun<input value={form.tahun} onChange={e=>setForm({...form, tahun:e.target.value})}/></label>
        <button>Simpan Plan</button>
      </form>
      {message && <p className="message">{message}</p>}
    </Panel>
    <Panel title="Import Plan Sidak Excel" desc="Download template .xlsx, isi daftar driver yang direncanakan, upload, preview, lalu submit.">
      <div className="import-actions"><button className="secondary" onClick={()=>downloadTemplate('template-plan-sidak.xlsx', templateRows('plan'))}><Download size={16}/> Download Template Excel</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewPlans(e.target.files?.[0])}/></label>{preview.length>0 && <button onClick={submitPlanImport} disabled={preview.some(r=>r.error)}>Submit Import Valid</button>}</div>
      {preview.length>0 && <PreviewTable rows={preview}/>} 
    </Panel>
    <Panel title="Row Data Plan Sidak" desc="Semua objek driver yang direncanakan untuk inspeksi. Status In Review/Done tidak tampil lagi di menu Mulai Inspeksi." action={<button onClick={()=>downloadXlsx('plan-sidak-fatigue.xlsx', rows)}><Download size={16}/> Export Excel</button>}>
      <DataTable rows={rows}/>
    </Panel>
  </div>
}

function FatigueInspections({ context, profile }) {
  const [plans, setPlans] = useState([])
  const [inspections, setInspections] = useState([])
  const [parameters, setParameters] = useState([])
  const [search, setSearch] = useState('')
  const [selectedPlan, setSelectedPlan] = useState(null)
  const [details, setDetails] = useState([])
  const [message, setMessage] = useState('')
  const [photo, setPhoto] = useState(null)
  const adminHO = isHeadOfficeAdmin(context)

  useEffect(() => { load() }, [context.id])
  async function load() {
    let planQ = supabase.from('fatigue_plans').select('*, drivers(nama_driver,nrp_driver), sites(site_name,site_code)').eq('status', 'Planned').order('created_at', { ascending:false })
    let inspQ = supabase.from('fatigue_inspections').select('*, drivers(nama_driver,nrp_driver), sites(site_name,site_code)').neq('status','Draft').order('created_at', { ascending:false })
    if (!adminHO) { planQ = planQ.eq('site_id', context.site_id); inspQ = inspQ.eq('site_id', context.site_id) }
    const [{ data: p }, { data: i }, { data: params }] = await Promise.all([
      planQ,
      inspQ,
      supabase.from('fatigue_parameters').select('*').eq('status','Aktif').order('urutan')
    ])
    setPlans((p || []).filter(x => x.sites?.site_code !== 'JIEP'))
    setInspections((i || []).filter(x => x.sites?.site_code !== 'JIEP'))
    setParameters(params || [])
  }
  function openInspection(plan){
    setMessage('')
    setPhoto(null)
    setSelectedPlan(plan)
    setDetails(parameters.map(param => ({ parameter_id: param.id, parameter_name: param.parameter_name, hasil: 'Aman', note: '' })))
  }
  function updateLocalDetail(index, patch){
    setDetails(prev => prev.map((d,i)=> i===index ? { ...d, ...patch } : d))
  }
  function closeModal(){
    setSelectedPlan(null)
    setDetails([])
    setPhoto(null)
  }
  async function submitInspection() {
    setMessage('')
    if (!selectedPlan) return
    if (!photo) return setMessage('Foto inspeksi wajib diupload sebelum submit.')
    if (details.some(d => d.hasil === 'Tidak Aman' && !String(d.note || '').trim())) return setMessage('Note wajib diisi untuk semua item Tidak Aman.')
    const { data: insp, error: inspErr } = await supabase.from('fatigue_inspections').insert({
      plan_id: selectedPlan.id,
      inspector_id: profile.id,
      site_id: selectedPlan.site_id,
      driver_id: selectedPlan.driver_id,
      tanggal_inspeksi: new Date().toISOString().slice(0,10),
      status: 'Submitted'
    }).select('*, drivers(nama_driver,nrp_driver), sites(site_name,site_code)').single()
    if (inspErr) return setMessage(inspErr.message)

    const compressed = await compressImage(photo)
    const path = `fatigue-inspection/${insp.id}-${Date.now()}.jpg`
    const { error: upErr } = await supabase.storage.from('inspection-photos').upload(path, compressed, { upsert: true, contentType: 'image/jpeg' })
    if (upErr) return setMessage(upErr.message)
    const { data: pub } = supabase.storage.from('inspection-photos').getPublicUrl(path)

    const detailRows = details.map(d => ({ inspection_id: insp.id, parameter_id: d.parameter_id, hasil: d.hasil, note: d.note || null }))
    const { error: detErr } = await supabase.from('fatigue_inspection_details').insert(detailRows)
    if (detErr) return setMessage(detErr.message)

    const { error: updErr } = await supabase.from('fatigue_inspections').update({ foto_inspeksi_url: pub.publicUrl }).eq('id', insp.id)
    if (updErr) return setMessage(updErr.message)
    const { error: planErr } = await supabase.from('fatigue_plans').update({ status: 'In Review' }).eq('id', selectedPlan.id)
    if (planErr) return setMessage(planErr.message)

    setMessage('Inspeksi submitted. Plan masuk status In Review dan menunggu approval atasan site.')
    closeModal()
    load()
  }

  const q = search.toLowerCase()
  const filteredPlans = plans.filter(p => `${p.drivers?.nama_driver||''} ${p.drivers?.nrp_driver||''}`.toLowerCase().includes(q))
  const rows = inspections.map(i => ({ id:i.id, tanggal:i.tanggal_inspeksi, driver:i.drivers?.nama_driver, nrp:i.drivers?.nrp_driver, site:i.sites?.site_name, status:i.status }))
  return <div className="stack">
    <Panel title="Pilih Plan Aktif" desc="Cari nama/NRP driver, lalu klik Mulai Inspeksi pada card yang sesuai. Plan Submitted/In Review/Done tidak ditampilkan di sini.">
      <div className="toolbar-inline"><input placeholder="Search nama / NRP driver..." value={search} onChange={e=>setSearch(e.target.value)} /></div>
      <div className="plan-card-list">
        {filteredPlans.map(p => <div className="plan-card" key={p.id}>
          <div><b>{p.drivers?.nama_driver}</b><small>NRP: {p.drivers?.nrp_driver} · {p.sites?.site_name} · {p.bulan}-{p.tahun}</small></div>
          <StatusPill value={p.status}/>
          <button onClick={()=>openInspection(p)}>Mulai Inspeksi</button>
        </div>)}
        {filteredPlans.length === 0 && <p className="muted">Tidak ada plan aktif yang cocok.</p>}
      </div>
      {message && <p className="message">{message}</p>}
    </Panel>
    <Panel title="Row Data Inspeksi" desc="Draft tidak ditampilkan. Hanya status Submitted, Approved, dan Rejected yang muncul di tabel ini." action={<button onClick={()=>exportToCsv('inspeksi-fatigue.csv', rows)}><Download size={16}/> Export CSV</button>}>
      <DataTable rows={rows}/>
    </Panel>
    {selectedPlan && <div className="modal-backdrop"><div className="modal-card wide-modal">
      <div className="row-between"><div><h3>Inspeksi Fatigue Driver</h3><p className="muted">{selectedPlan.drivers?.nama_driver} / {selectedPlan.drivers?.nrp_driver} · {selectedPlan.sites?.site_name} · {selectedPlan.bulan}-{selectedPlan.tahun}</p></div><button className="secondary" onClick={closeModal}><X size={16}/> Batal</button></div>
      <div className="form-grid single-line"><label>Upload Foto Inspeksi<input type="file" accept="image/*" onChange={e=>setPhoto(e.target.files?.[0])}/></label><div className="message">Foto akan dikompres sebelum upload. Batal tidak membuat draft.</div></div>
      <div className="checklist-list">
        {details.map((d,idx) => <div className="check-row" key={d.parameter_id}>
          <div><b>{d.parameter_name}</b><small>{d.hasil === 'Tidak Aman' ? 'Note wajib diisi' : 'Aman'}</small></div>
          <select value={d.hasil} onChange={e=>updateLocalDetail(idx, { hasil:e.target.value })}><option>Aman</option><option>Tidak Aman</option></select>
          <input placeholder="Note / temuan" value={d.note || ''} onChange={e=>updateLocalDetail(idx, { note:e.target.value })}/>
        </div>)}
      </div>
      <div className="modal-actions"><button className="secondary" onClick={closeModal}>Batal</button><button onClick={submitInspection}>Submit Inspeksi</button></div>
      {message && <p className="message">{message}</p>}
    </div></div>}
  </div>
}

function FatigueOutstanding({ context, profile }) {
  const [rows, setRows] = useState([])
  const [message, setMessage] = useState('')
  const [closing, setClosing] = useState(null)
  const [closeNote, setCloseNote] = useState('')
  const [file, setFile] = useState(null)

  useEffect(() => { load() }, [context.id])
  async function load() {
    let q = supabase.from('fatigue_outstandings').select('*, drivers(nama_driver,nrp_driver), fatigue_parameters(parameter_name), sites(site_name,site_code)').order('created_at', { ascending:false })
    if (!isHeadOfficeAdmin(context)) q = q.eq('site_id', context.site_id)
    const { data } = await q
    setRows(data || [])
  }
  async function closeOutstanding() {
    setMessage('')
    if (!closing || !file || !closeNote) return setMessage('Foto evidence dan catatan close wajib diisi.')
    const compressed = await compressImage(file)
    const path = `fatigue/${closing.id}-${Date.now()}.jpg`
    const { error: upErr } = await supabase.storage.from('evidence-photos').upload(path, compressed, { upsert: true, contentType: 'image/jpeg' })
    if (upErr) return setMessage(upErr.message)
    const { data: pub } = supabase.storage.from('evidence-photos').getPublicUrl(path)
    const { error } = await supabase.from('fatigue_outstandings').update({
      status: 'Closed', evidence_photo_url: pub.publicUrl, close_note: closeNote, closed_by: profile.id, closed_at: new Date().toISOString()
    }).eq('id', closing.id)
    if (error) setMessage(error.message)
    else { setMessage('Outstanding closed, menunggu approval atasan.'); setClosing(null); setCloseNote(''); setFile(null); load() }
  }
  const exportRows = rows.map(o => ({ id:o.id, driver:o.drivers?.nama_driver, nrp:o.drivers?.nrp_driver, parameter:o.fatigue_parameters?.parameter_name, description:o.description, status:o.status, site:o.sites?.site_name }))
  return <div className="stack">
    <Panel title="Outstanding Sidak Fatigue" desc="Close wajib foto evidence. Status Closed belum final sebelum approval atasan." action={<button onClick={()=>exportToCsv('outstanding-fatigue.csv', exportRows)}><Download size={16}/> Export CSV</button>}>
      <div className="cards-grid">{rows.map(o => <div className="mini-card" key={o.id}>
        <div className="row-between"><b>{o.fatigue_parameters?.parameter_name}</b><StatusPill value={o.status}/></div>
        <p>{o.description}</p><small>{o.drivers?.nama_driver} · {o.sites?.site_name}</small>
        {o.status === 'Open' && <button onClick={()=>setClosing(o)}>Close Temuan</button>}
        {o.evidence_photo_url && <a href={o.evidence_photo_url} target="_blank">Lihat Foto Evidence</a>}
      </div>)}</div>
      {rows.length === 0 && <p className="muted">Belum ada outstanding.</p>}
    </Panel>
    {closing && <Panel title="Close Outstanding" desc="Foto akan dikompres sebelum upload agar hemat storage Supabase.">
      <div className="form-grid">
        <label>Foto Evidence<input type="file" accept="image/*" onChange={e=>setFile(e.target.files?.[0])}/></label>
        <label>Catatan Close<input value={closeNote} onChange={e=>setCloseNote(e.target.value)} placeholder="Tindak lanjut yang dilakukan"/></label>
        <button onClick={closeOutstanding}><Upload size={16}/> Upload & Close</button>
        <button className="secondary" onClick={()=>setClosing(null)}>Batal</button>
      </div>
      {message && <p className="message">{message}</p>}
    </Panel>}
  </div>
}

function FatigueApproval({ context, profile }) {
  const [inspections, setInspections] = useState([])
  const [outs, setOuts] = useState([])
  const [review, setReview] = useState(null)
  const [reviewDetails, setReviewDetails] = useState([])
  const [note, setNote] = useState('')
  const hasAccess = canApprove(context.role)
  useEffect(() => { load() }, [context.id])
  async function load() {
    if (!hasAccess) return
    const [{ data: i }, { data: o }] = await Promise.all([
      (isHeadOfficeAdmin(context) ? supabase.from('fatigue_inspections').select('*, drivers(nama_driver,nrp_driver), sites(site_name,site_code)').eq('status','Submitted') : supabase.from('fatigue_inspections').select('*, drivers(nama_driver,nrp_driver), sites(site_name,site_code)').eq('site_id', context.site_id).eq('status','Submitted')),
      (isHeadOfficeAdmin(context) ? supabase.from('fatigue_outstandings').select('*, drivers(nama_driver,nrp_driver), fatigue_parameters(parameter_name), sites(site_name,site_code)').eq('status','Closed') : supabase.from('fatigue_outstandings').select('*, drivers(nama_driver,nrp_driver), fatigue_parameters(parameter_name), sites(site_name,site_code)').eq('site_id', context.site_id).eq('status','Closed'))
    ])
    setInspections(i || [])
    setOuts(o || [])
  }
  async function openInspectionReview(insp){
    const { data } = await supabase.from('fatigue_inspection_details').select('*, fatigue_parameters(parameter_name, urutan)').eq('inspection_id', insp.id).order('created_at')
    setReview({ type:'inspection', data: insp }); setReviewDetails(data || []); setNote('')
  }
  function openOutstandingReview(o){ setReview({ type:'outstanding', data:o }); setReviewDetails([]); setNote('') }
  async function approveInspection(){
    const insp = review.data
    const unsafe = reviewDetails.filter(d => d.hasil === 'Tidak Aman')
    for (const d of unsafe) {
      const { data: existing } = await supabase.from('fatigue_outstandings').select('id').eq('detail_id', d.id).maybeSingle()
      if (!existing) await supabase.from('fatigue_outstandings').insert({ inspection_id: insp.id, detail_id: d.id, site_id: insp.site_id, driver_id: insp.driver_id, parameter_id: d.parameter_id, description: d.note, status: 'Open' })
    }
    await supabase.from('fatigue_inspections').update({ status:'Approved', approved_by: profile.id, approved_at: new Date().toISOString(), approval_note: note }).eq('id', insp.id)
    await supabase.from('fatigue_plans').update({ status:'Done' }).eq('id', insp.plan_id)
    setReview(null); load()
  }
  async function rejectInspection(){
    const insp = review.data
    await supabase.from('fatigue_inspections').update({ status:'Rejected', approved_by: profile.id, approved_at: new Date().toISOString(), approval_note: note }).eq('id', insp.id)
    await supabase.from('fatigue_plans').update({ status:'Planned' }).eq('id', insp.plan_id)
    setReview(null); load()
  }
  async function approveOutstanding(){
    await supabase.from('fatigue_outstandings').update({ status:'Approved', approved_by: profile.id, approved_at: new Date().toISOString(), approval_note: note }).eq('id', review.data.id)
    setReview(null); load()
  }
  async function rejectOutstanding(){
    await supabase.from('fatigue_outstandings').update({ status:'Open', approved_by: profile.id, approved_at: new Date().toISOString(), approval_note: note }).eq('id', review.data.id)
    setReview(null); load()
  }
  return <div className="stack">
    {!hasAccess && <div className="warning"><XCircle/> Role ini tidak memiliki akses approval. Login sebagai Atasan Site/App Admin/Platform Admin.</div>}
    <Panel title="Approval Inspeksi Submitted" desc="Klik Review untuk melihat foto, checklist, dan catatan sebelum approve/reject.">
      {inspections.map(i => <div className="approval-row" key={i.id}><div><b>{i.drivers?.nama_driver}</b><small>{i.tanggal_inspeksi} · {i.sites?.site_name}</small></div><StatusPill value={i.status}/>{hasAccess && <button onClick={()=>openInspectionReview(i)}><Eye size={16}/> Review</button>}</div>)}
      {inspections.length === 0 && <p className="muted">Tidak ada inspeksi menunggu approval.</p>}
    </Panel>
    <Panel title="Approval Close Outstanding" desc="Close outstanding belum final sebelum disetujui atasan.">
      {outs.map(o => <div className="approval-row" key={o.id}><div><b>{o.fatigue_parameters?.parameter_name}</b><small>{o.drivers?.nama_driver} · {o.description}</small></div><StatusPill value={o.status}/>{hasAccess && <button onClick={()=>openOutstandingReview(o)}><Eye size={16}/> Review</button>}</div>)}
      {outs.length === 0 && <p className="muted">Tidak ada close outstanding menunggu approval.</p>}
    </Panel>
    {review && <div className="modal-backdrop"><div className="modal-card">
      <div className="row-between"><h3>{review.type === 'inspection' ? 'Review Inspeksi' : 'Review Close Outstanding'}</h3><button className="secondary" onClick={()=>setReview(null)}><X size={16}/> Tutup</button></div>
      {review.type === 'inspection' ? <div className="stack">
        <p><b>Driver:</b> {review.data.drivers?.nama_driver} / {review.data.drivers?.nrp_driver}</p>
        <p><b>Site:</b> {review.data.sites?.site_name} · <b>Tanggal:</b> {review.data.tanggal_inspeksi}</p>
        {review.data.foto_inspeksi_url && <a href={review.data.foto_inspeksi_url} target="_blank">Lihat Foto Inspeksi</a>}
        <div className="checklist-list">{reviewDetails.map(d => <div className="check-row" key={d.id}><div><b>{d.fatigue_parameters?.parameter_name}</b><small>{d.note || '-'}</small></div><StatusPill value={d.hasil}/></div>)}</div>
      </div> : <div className="stack">
        <p><b>Parameter:</b> {review.data.fatigue_parameters?.parameter_name}</p><p><b>Driver:</b> {review.data.drivers?.nama_driver}</p><p><b>Temuan:</b> {review.data.description}</p><p><b>Catatan Close:</b> {review.data.close_note}</p>{review.data.evidence_photo_url && <a href={review.data.evidence_photo_url} target="_blank">Lihat Foto Evidence</a>}
      </div>}
      <label className="modal-note">Catatan Atasan<input value={note} onChange={e=>setNote(e.target.value)} placeholder="Catatan approve/reject"/></label>
      <div className="modal-actions">{review.type === 'inspection' ? <><button className="danger" onClick={rejectInspection}><X size={16}/> Reject</button><button onClick={approveInspection}><Check size={16}/> Approve</button></> : <><button className="danger" onClick={rejectOutstanding}><X size={16}/> Reject Close</button><button onClick={approveOutstanding}><Check size={16}/> Approve Close</button></>}</div>
    </div></div>}
  </div>
}
function AdminPanel({ context, profile }) {
  const [tab, setTab] = useState('vendors')
  const [sites, setSites] = useState([])
  const [vendors, setVendors] = useState([])
  const [message, setMessage] = useState('')
  const [vendorForm, setVendorForm] = useState({ id:null, vendor_code:'', vendor_name:'', status:'Aktif' })
  const [siteForm, setSiteForm] = useState({ id:null, site_code:'', site_name:'', region:'Operation', status:'Aktif' })
  const [vendorPreview, setVendorPreview] = useState([])
  const [sitePreview, setSitePreview] = useState([])

  useEffect(()=>{ load() }, [])
  async function load(){
    const [{data:s},{data:v}] = await Promise.all([
      supabase.from('sites').select('*').order('site_code'),
      supabase.from('vendors').select('*').order('vendor_code')
    ])
    setSites(s||[]); setVendors(v||[])
  }
  function resetVendor(){ setVendorForm({ id:null, vendor_code:'', vendor_name:'', status:'Aktif' }) }
  function resetSite(){ setSiteForm({ id:null, site_code:'', site_name:'', region:'Operation', status:'Aktif' }) }
  async function saveVendor(e){
    e.preventDefault(); setMessage('')
    const payload = { vendor_code: vendorForm.id && cleanText(vendorForm.vendor_code) ? cleanText(vendorForm.vendor_code).toUpperCase() : makeVendorCode(), vendor_name: cleanText(vendorForm.vendor_name), status: vendorForm.status || 'Aktif' }
    const q = vendorForm.id ? supabase.from('vendors').update(payload).eq('id', vendorForm.id) : supabase.from('vendors').insert(payload)
    const {error}=await q
    if(error)setMessage(error.message); else{resetVendor(); setMessage('Vendor tersimpan.'); load()}
  }
  async function saveSite(e){
    e.preventDefault(); setMessage('')
    const payload = { site_code: cleanText(siteForm.site_code).toUpperCase(), site_name: cleanText(siteForm.site_name), region: cleanText(siteForm.region) || 'Operation', status: siteForm.status || 'Aktif' }
    const q = siteForm.id ? supabase.from('sites').update(payload).eq('id', siteForm.id) : supabase.from('sites').upsert(payload, { onConflict:'site_code' })
    const {error}=await q
    if(error)setMessage(error.message); else{resetSite(); setMessage('Site tersimpan.'); load()}
  }
  async function toggleVendor(v){ const next=v.status==='Aktif'?'Nonaktif':'Aktif'; const {error}=await supabase.from('vendors').update({status:next}).eq('id',v.id); if(error)setMessage(error.message); else{setMessage(`Vendor ${next}.`); load()} }
  async function toggleSite(v){ const next=v.status==='Aktif'?'Nonaktif':'Aktif'; const {error}=await supabase.from('sites').update({status:next}).eq('id',v.id); if(error)setMessage(error.message); else{setMessage(`Site ${next}.`); load()} }
  async function previewMaster(type, file){
    if(!file) return
    setMessage('')
    try{
      const rows = (await parseExcelOrCsv(file)).map(normalizeRow)
      if(type==='vendors'){
        setVendorPreview(rows.map((r,idx)=>({ row: idx+2, vendor_code: makeVendorCode(idx), vendor_name: cleanText(r.vendor_name), status: cleanText(r.status)||'Aktif', error: !cleanText(r.vendor_name) ? 'vendor_name wajib. vendor_code dibuat otomatis oleh sistem' : '' })))
      }
      if(type==='sites'){
        setSitePreview(rows.map((r,idx)=>({ row: idx+2, site_code: cleanText(r.site_code).toUpperCase(), site_name: cleanText(r.site_name||r.site_code), region: cleanText(r.region)||'Operation', status: cleanText(r.status)||'Aktif', error: !cleanText(r.site_code) ? 'site_code wajib' : '' })))
      }
    }catch(e){ setMessage(e.message) }
  }
  async function submitVendorImport(){
    const valid = vendorPreview.filter(r=>!r.error).map((r,idx)=>({ vendor_code:r.vendor_code || makeVendorCode(idx), vendor_name:r.vendor_name, status:r.status }))
    const {error}=await supabase.from('vendors').insert(valid)
    if(error)setMessage(error.message); else{setMessage(`Import vendor berhasil: ${valid.length} baris.`); setVendorPreview([]); load()}
  }
  async function submitSiteImport(){
    const valid = sitePreview.filter(r=>!r.error).map(r=>({ site_code:r.site_code, site_name:r.site_name, region:r.region, status:r.status }))
    const {error}=await supabase.from('sites').upsert(valid,{onConflict:'site_code'})
    if(error)setMessage(error.message); else{setMessage(`Import site berhasil: ${valid.length} baris.`); setSitePreview([]); load()}
  }
  const vendorRows = vendors.map(v=>({vendor_code:v.vendor_code,vendor_name:v.vendor_name,status:v.status}))
  const siteRows = sites.map(s=>({site_code:s.site_code,site_name:s.site_name,region:s.region,status:s.status}))
  return <div className="stack">
    <div className="tab-row"><button className={tab==='vendors'?'active':''} onClick={()=>setTab('vendors')}>Master Vendor</button><button className={tab==='sites'?'active':''} onClick={()=>setTab('sites')}>Master Site</button><button className={tab==='access'?'active':''} onClick={()=>setTab('access')}>Mapping Akses</button></div>
    {message && <p className="message">{message}</p>}
    {tab==='vendors' && <><Panel title={vendorForm.id?'Edit Vendor':'Tambah Vendor'} desc="Admin dapat menambahkan vendor dari UI atau import Excel. Kode vendor dibuat otomatis oleh sistem."><form className="form-grid" onSubmit={saveVendor}><label>Kode Vendor<input disabled value={vendorForm.id ? vendorForm.vendor_code : 'Otomatis oleh sistem'} title="Kode vendor dibuat otomatis oleh sistem"/></label><label>Nama Vendor<input required value={vendorForm.vendor_name} onChange={e=>setVendorForm({...vendorForm,vendor_name:e.target.value})}/></label><label>Status<select value={vendorForm.status} onChange={e=>setVendorForm({...vendorForm,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label><button>{vendorForm.id?'Update Vendor':'Simpan Vendor'}</button>{vendorForm.id && <button type="button" className="secondary" onClick={resetVendor}>Batal Edit</button>}</form></Panel><Panel title="Import Vendor Excel" desc="Download template .xlsx, isi nama vendor, upload, preview, lalu submit. vendor_code dibuat otomatis oleh sistem."><div className="import-actions"><button onClick={()=>downloadTemplate('template-master-vendor.xlsx', templateRows('vendor'))}><Download size={16}/> Download Template Excel</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewMaster('vendors',e.target.files?.[0])}/></label>{vendorPreview.length>0 && <button onClick={submitVendorImport} disabled={vendorPreview.some(r=>r.error)}>Submit Import Valid</button>}</div>{vendorPreview.length>0 && <PreviewTable rows={vendorPreview}/>}</Panel><Panel title="Row Data Vendor" desc="Data dapat diedit/nonaktifkan. Export menggunakan Excel." action={<button onClick={()=>downloadXlsx('master-vendor.xlsx', vendorRows)}><Download size={16}/> Export Excel</button>}><DataTable rows={vendorRows.map(v=>({...v, aksi:''}))} customActions={(idx)=>{const v=vendors[idx]; return <div className="row-actions"><button className="secondary" onClick={()=>setVendorForm(v)}>Edit</button><button className="secondary" onClick={()=>toggleVendor(v)}>{v.status==='Aktif'?'Nonaktifkan':'Aktifkan'}</button></div>}} /></Panel></>}
    {tab==='sites' && <><Panel title={siteForm.id?'Edit Site':'Tambah Site'} desc="JIEP adalah Head Office. Site operation dipakai untuk dashboard achievement."><form className="form-grid" onSubmit={saveSite}><label>Kode Site<input required value={siteForm.site_code} onChange={e=>setSiteForm({...siteForm,site_code:e.target.value})}/></label><label>Nama Site<input required value={siteForm.site_name} onChange={e=>setSiteForm({...siteForm,site_name:e.target.value})}/></label><label>Region<input value={siteForm.region} onChange={e=>setSiteForm({...siteForm,region:e.target.value})}/></label><label>Status<select value={siteForm.status} onChange={e=>setSiteForm({...siteForm,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label><button>{siteForm.id?'Update Site':'Simpan Site'}</button>{siteForm.id && <button type="button" className="secondary" onClick={resetSite}>Batal Edit</button>}</form></Panel><Panel title="Import Site Excel" desc="Download template .xlsx, isi data, upload, preview, lalu submit."><div className="import-actions"><button onClick={()=>downloadTemplate('template-master-site.xlsx', templateRows('site'))}><Download size={16}/> Download Template Excel</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewMaster('sites',e.target.files?.[0])}/></label>{sitePreview.length>0 && <button onClick={submitSiteImport} disabled={sitePreview.some(r=>r.error)}>Submit Import Valid</button>}</div>{sitePreview.length>0 && <PreviewTable rows={sitePreview}/>}</Panel><Panel title="Row Data Site" desc="Data bisa diedit/nonaktifkan. JIEP tetap tersedia sebagai Head Office." action={<button onClick={()=>downloadXlsx('master-site.xlsx', siteRows)}><Download size={16}/> Export Excel</button>}><DataTable rows={siteRows.map(v=>({...v, aksi:''}))} customActions={(idx)=>{const v=sites[idx]; return <div className="row-actions"><button className="secondary" onClick={()=>setSiteForm(v)}>Edit</button><button className="secondary" onClick={()=>toggleSite(v)}>{v.status==='Aktif'?'Nonaktifkan':'Aktifkan'}</button></div>}} /></Panel></>}
    {tab==='access' && <AccessMapping context={context} profile={profile}/>} 
  </div>
}

function UserExistingSearch({ profiles, value, onPick }){
  const [open, setOpen] = useState(false)
  const [q, setQ] = useState('')
  const selected = profiles.find(p => p.id === value)
  const list = profiles.filter(p => `${p.nama || ''} ${p.email || ''} ${p.nrp || ''}`.toLowerCase().includes(q.toLowerCase())).slice(0, 50)
  function choose(id){ onPick(id); setOpen(false); setQ('') }
  return <label>User Existing
    <div className="user-picker-field">
      <button type="button" className="secondary user-picker-trigger" onClick={()=>setOpen(true)}>{selected ? `${selected.nama || selected.email} · ${selected.email || ''}${selected.nrp ? ' · ' + selected.nrp : ''}` : 'Cari dan pilih user existing'}</button>
      {selected && <button type="button" className="ghost-btn small" onClick={()=>onPick('')}>Clear</button>}
    </div>
    {open && <div className="modal-backdrop" onMouseDown={()=>setOpen(false)}><div className="modal-card user-picker-modal" onMouseDown={e=>e.stopPropagation()}><div className="modal-head"><div><h3>Pilih User Existing</h3><p>Cari berdasarkan nama, email, atau NRP.</p></div><button type="button" className="secondary small" onClick={()=>setOpen(false)}>Tutup</button></div><div className="table-search"><Search size={16}/><input autoFocus placeholder="Ketik nama / email / NRP..." value={q} onChange={e=>setQ(e.target.value)} /></div><div className="user-picker-list">{list.map(p=><button type="button" key={p.id} onClick={()=>choose(p.id)}><b>{p.nama || p.email}</b><span>{p.email}{p.nrp ? ` · ${p.nrp}` : ''}</span></button>)}{!list.length && <p className="muted">User tidak ditemukan.</p>}</div></div></div>}
  </label>
}

function AccessMapping({ context, profile }){
  const [profiles,setProfiles]=useState([])
  const [apps,setApps]=useState([])
  const [sites,setSites]=useState([])
  const [accessRows,setAccessRows]=useState([])
  const [message,setMessage]=useState('')
  const [loading,setLoading]=useState(false)
  const [newUser,setNewUser]=useState({ nama:'', nrp:'', email:'', password:'', app_id:'', role:'GL', site_id:'' })
  const [existing,setExisting]=useState({ user_id:'', app_id:'', role:'GL', site_id:'' })
  const [editing,setEditing]=useState(null)
  const [bulkPreview,setBulkPreview]=useState([])
  const [existingBulkPreview,setExistingBulkPreview]=useState([])
  useEffect(()=>{load()},[])
  async function load(){
    const [{data:p},{data:a},{data:s},{data:m}] = await Promise.all([
      supabase.from('users_profile').select('*').neq('status','Nonaktif').order('nama'),
      supabase.from('applications').select('*').eq('status','Aktif').order('app_name'),
      supabase.from('sites').select('*').eq('status','Aktif').order('site_code'),
      supabase.from('user_app_access').select('*, users_profile(id,nama,email,nrp,status), applications(id,app_name,app_code), sites(id,site_name,site_code)').order('created_at',{ascending:false})
    ])
    const profilesData=p||[], appsData=a||[], sitesData=s||[], mapData=m||[]
    setProfiles(profilesData); setApps(appsData); setSites(sitesData); setAccessRows(mapData)
    const firstApp=appsData[0]?.id || ''
    const firstSite=sitesData[0]?.id || ''
    setNewUser(f=>({ ...f, app_id:f.app_id || firstApp, site_id:f.site_id || firstSite }))
    setExisting(f=>({ ...f, app_id:f.app_id || firstApp, site_id:f.site_id || firstSite }))
  }
  function normalizeEmail(v){ return cleanText(v).toLowerCase() }
  function resetNewUser(){ setNewUser({ nama:'', nrp:'', email:'', password:'', app_id:apps[0]?.id||'', role:'GL', site_id:sites[0]?.id||'' }) }
  function resetExisting(){ setExisting({ user_id:'', app_id:apps[0]?.id||'', role:'GL', site_id:sites[0]?.id||'' }); setEditing(null) }
  function findProfileByEmail(email){ return profiles.find(p=>normalizeEmail(p.email)===normalizeEmail(email)) }
  function findProfileByNrp(nrp){ const key=cleanText(nrp).toLowerCase(); return key ? profiles.find(p=>cleanText(p.nrp).toLowerCase()===key) : null }
  function findProfileByName(nama){ const key=cleanText(nama).toLowerCase(); return key ? profiles.find(p=>cleanText(p.nama).toLowerCase()===key) : null }
  function findAccessDuplicate(x, ignoreId=null){ return accessRows.find(r=>String(r.id)!==String(ignoreId||'') && String(r.user_id||r.users_profile?.id||'')===String(x.user_id||'') && String(r.app_id||r.applications?.id||'')===String(x.app_id||'') && String(r.role||'')===String(x.role||'') && String(r.site_id||r.sites?.id||'')===String(x.site_id||'')) }
  async function createUserWithMapping(e){
    e.preventDefault(); setMessage(''); setLoading(true)
    try{
      const payload={ nama: cleanText(newUser.nama) || normalizeEmail(newUser.email), nrp: cleanText(newUser.nrp), email: normalizeEmail(newUser.email), password: newUser.password, app_id: newUser.app_id, role: newUser.role, site_id: newUser.site_id || null }
      if(!payload.email) throw new Error('Email user baru wajib diisi')
      if(payload.email === normalizeEmail(profile?.email)) throw new Error('Email user baru tidak boleh sama dengan email admin yang sedang login. Pilih User Existing jika ingin menambah mapping untuk akun admin.')
      if(!payload.password || payload.password.length < 6) throw new Error('Password minimal 6 karakter')
      if(!payload.app_id || !payload.role) throw new Error('Aplikasi dan role wajib dipilih')
      if(findProfileByEmail(payload.email)) throw new Error('Email sudah terdaftar. Silahkan mapping pada User Existing.')
      if(payload.nrp && findProfileByNrp(payload.nrp)) throw new Error('NRP sudah terdaftar. Silahkan mapping pada User Existing.')
      if(cleanText(newUser.nama) && findProfileByName(newUser.nama)) throw new Error('Nama user sudah terdaftar. Silahkan mapping pada User Existing.')
      const { data, error } = await supabase.functions.invoke('admin-create-user', { body: payload })
      if(error) throw error
      if(data?.error) throw new Error(data.error)
      setMessage(`User ${payload.email} berhasil dibuat dan dimapping.`)
      resetNewUser(); await load()
    }catch(err){ setMessage(err.message || String(err)) }
    finally{ setLoading(false) }
  }
  async function saveExistingMapping(e){
    e.preventDefault(); setMessage(''); setLoading(true)
    try{
      if(!existing.user_id) throw new Error('Pilih user existing terlebih dahulu')
      if(!existing.app_id || !existing.role) throw new Error('Aplikasi dan role wajib dipilih')
      const payload={ user_id:existing.user_id, app_id:existing.app_id, role:existing.role, site_id:existing.site_id || null, status:'Aktif' }
      if(findAccessDuplicate(payload, editing?.id)) throw new Error('Mapping akses ini sudah ada. User, aplikasi, role, dan site tidak boleh double.')
      let error
      if(editing){ ({ error } = await supabase.from('user_app_access').update(payload).eq('id',editing.id)) }
      else { ({ error } = await supabase.from('user_app_access').insert(payload)) }
      if(error) throw error
      setMessage(editing ? 'Mapping akses berhasil diupdate.' : 'Mapping akses user existing berhasil disimpan.')
      resetExisting(); await load()
    }catch(err){ setMessage(err.message || String(err)) }
    finally{ setLoading(false) }
  }
  function downloadAccessTemplate(){
    downloadXlsx('template-bulk-mapping-user-existing.xlsx', [
      { nama:'Nama User', nrp:'NRP001', email:'user@company.co.id', password:'password123', app_code:'sidak_fatigue', role:'GL', site_code:'BAYA', status:'Aktif' },
      { nama:'Nama User', nrp:'NRP001', email:'user@company.co.id', password:'password123', app_code:'drd_driver', role:'GL', site_code:'BAYA', status:'Aktif' },
      { nama:'Nama User', nrp:'NRP001', email:'user@company.co.id', password:'password123', app_code:'INSPEKSI', role:'GL', site_code:'BAYA', status:'Aktif' }
    ])
  }
  async function previewAccessUpload(file){
    if(!file) return
    setMessage('')
    try{
      const raw = (await parseExcelOrCsv(file)).map(normalizeRow)
      const seenMap = new Set()
      const mapped = raw.map((r,idx)=>{
        const email = normalizeEmail(r.email)
        const nrp = cleanText(r.nrp)
        const appCode = cleanText(r.app_code || r.aplikasi)
        const role = cleanText(r.role) || 'GL'
        const hasPasswordColumn = cleanText(r.password)
        const siteCode = cleanText(r.site_code).toUpperCase()
        const app = apps.find(a => String(a.status||'').toLowerCase()==='aktif' && String(a.app_code||'').toLowerCase() === String(appCode).toLowerCase())
        const site = siteCode ? sites.find(s => s.site_code === siteCode) : null
        const prof = profiles.find(p => normalizeEmail(p.email) === email)
        const nrpProf = nrp ? findProfileByNrp(nrp) : null
        const nameProf = cleanText(r.nama) ? findProfileByName(r.nama) : null
        const userKey = prof?.id || email
        const mapKey = [userKey, app?.id||appCode, role, site?.id||''].join('|')
        let error = ''
        let mode = prof ? 'Mapping existing user' : 'Create user baru + mapping'
        if(!email) error = 'email wajib'
        else if(!app) error = 'app_code tidak ditemukan / tidak aktif'
        else if(!role) error = 'role wajib'
        else if(siteCode && !site) error = 'site_code tidak ditemukan'
        else if(seenMap.has(mapKey)) error = 'mapping akses double di file import'
        else if(prof && nrpProf && nrpProf.id !== prof.id) error = 'NRP sudah dipakai email/user lain'
        else if(prof && findAccessDuplicate({ user_id:prof.id, app_id:app.id, role, site_id:site?.id||null })) error = 'mapping akses sudah ada'
        else if(!prof && nrpProf) error = 'NRP sudah terdaftar. Gunakan email user existing tersebut untuk bulk mapping.'
        else if(!prof && nameProf) error = 'nama sudah terdaftar. Gunakan email user existing tersebut untuk bulk mapping.'
        else if(!prof && (!cleanText(r.password) || cleanText(r.password).length < 6)) error = 'user baru wajib password minimal 6 karakter'
        seenMap.add(mapKey)
        return { row:idx+2, mode, nama:cleanText(r.nama)||prof?.nama||email, nrp:nrp || prof?.nrp || '', email, password:cleanText(r.password), app_code:app?.app_code||appCode, app_id:app?.id, role, site_code:siteCode||'ALL', site_id:site?.id||null, user_id:prof?.id||null, status:cleanText(r.status)||'Aktif', error }
      })
      setBulkPreview(mapped)
    }catch(e){ setMessage(e.message || String(e)) }
  }
  async function submitAccessUpload(){
    setLoading(true); setMessage('')
    let ok = 0; const fail = []
    for (const r of bulkPreview.filter(x=>!x.error)){
      try{
        if(r.user_id){
          if(findAccessDuplicate({ user_id:r.user_id, app_id:r.app_id, role:r.role, site_id:r.site_id })) throw new Error('mapping akses sudah ada')
          const { error } = await supabase.from('user_app_access').insert({ user_id:r.user_id, app_id:r.app_id, role:r.role, site_id:r.site_id, status:'Aktif' })
          if(error) throw error
        } else {
          if(findProfileByEmail(r.email)) throw new Error('email sudah terdaftar. Silahkan mapping pada User Existing.')
          if(r.nrp && findProfileByNrp(r.nrp)) throw new Error('NRP sudah terdaftar. Silahkan mapping pada User Existing.')
          if(r.nama && findProfileByName(r.nama)) throw new Error('nama sudah terdaftar. Silahkan mapping pada User Existing.')
          const { data, error } = await supabase.functions.invoke('admin-create-user', { body:{ nama:r.nama, nrp:r.nrp, email:r.email, password:r.password, app_id:r.app_id, role:r.role, site_id:r.site_id } })
          if(error) throw error
          if(data?.error) throw new Error(data.error)
        }
        ok++
      }catch(e){ fail.push('Baris '+r.row+': '+(e.message || e)) }
    }
    setLoading(false)
    setMessage('Import mapping berhasil: '+ok+'. '+(fail.length ? fail.join(' | ') : ''))
    if(!fail.length) setBulkPreview([])
    await load()
  }

  async function previewExistingMappingUpload(file){
    if(!file) return
    setMessage('')
    try{
      const raw = (await parseExcelOrCsv(file)).map(normalizeRow)
      const seenMap = new Set()
      const mapped = raw.map((r,idx)=>{
        const email = normalizeEmail(r.email)
        const appCode = cleanText(r.app_code || r.aplikasi)
        const role = cleanText(r.role) || 'GL'
        const siteCode = cleanText(r.site_code).toUpperCase()
        const app = apps.find(a => String(a.status||'').toLowerCase()==='aktif' && String(a.app_code||'').toLowerCase() === String(appCode).toLowerCase())
        const site = siteCode ? sites.find(st => st.site_code === siteCode) : null
        const prof = profiles.find(pr => normalizeEmail(pr.email) === email)
        const mapKey = [prof?.id || email, app?.id || appCode, role, site?.id || ''].join('|')
        let error = ''
        if(!email) error = 'email wajib'
        else if(hasPasswordColumn) error = 'Mapping user existing tidak membutuhkan password. Hapus isi/kolom password pada file ini.'
        else if(!prof) error = 'email belum terdaftar. Gunakan Upload Excel User Mapping untuk membuat user baru.'
        else if(!app) error = 'app_code tidak ditemukan / tidak aktif'
        else if(!role) error = 'role wajib'
        else if(siteCode && !site) error = 'site_code tidak ditemukan'
        else if(seenMap.has(mapKey)) error = 'mapping akses double di file import'
        else if(findAccessDuplicate({ user_id:prof.id, app_id:app.id, role, site_id:site?.id||null })) error = 'mapping akses sudah ada'
        seenMap.add(mapKey)
        return { row:idx+2, nama:prof?.nama || cleanText(r.nama) || email, nrp:prof?.nrp || cleanText(r.nrp) || '', email, app_code:app?.app_code||appCode, app_id:app?.id, role, site_code:siteCode||'ALL', site_id:site?.id||null, user_id:prof?.id||null, status:cleanText(r.status)||'Aktif', error }
      })
      setExistingBulkPreview(mapped)
    }catch(e){ setMessage(e.message || String(e)) }
  }
  async function submitExistingMappingUpload(){
    setLoading(true); setMessage('')
    let ok = 0; const fail = []
    for (const r of existingBulkPreview.filter(x=>!x.error)){
      try{
        if(!r.user_id) throw new Error('email belum terdaftar')
        if(findAccessDuplicate({ user_id:r.user_id, app_id:r.app_id, role:r.role, site_id:r.site_id })) throw new Error('mapping akses sudah ada')
        const { error } = await supabase.from('user_app_access').insert({ user_id:r.user_id, app_id:r.app_id, role:r.role, site_id:r.site_id, status:'Aktif' })
        if(error) throw error
        ok++
      }catch(e){ fail.push('Baris '+r.row+': '+(e.message || e)) }
    }
    setLoading(false)
    setMessage('Import mapping existing berhasil: '+ok+'. '+(fail.length ? fail.join(' | ') : ''))
    if(!fail.length) setExistingBulkPreview([])
    await load()
  }

  async function toggleAccess(row){ const next=row.status==='Aktif'?'Nonaktif':'Aktif'; const {error}=await supabase.from('user_app_access').update({status:next}).eq('id',row.id); if(error)setMessage(error.message); else{setMessage(`Akses ${next}.`); load()} }
  async function deleteAccess(row){ if(!confirm('Hapus mapping akses ini?')) return; const {error}=await supabase.from('user_app_access').delete().eq('id',row.id); if(error)setMessage(error.message); else{setMessage('Mapping akses dihapus.'); load()} }
  function editAccess(row){ setEditing(row); setExisting({ user_id:row.user_id || row.users_profile?.id || '', app_id:row.app_id || row.applications?.id || '', role:row.role || 'GL', site_id:row.site_id || row.sites?.id || '' }); window.scrollTo({top:0,behavior:'smooth'}) }
  const rows=accessRows.map(r=>({ user:r.users_profile?.nama, nrp:r.users_profile?.nrp || '-', email:r.users_profile?.email, app:r.applications?.app_name, role:r.role, site:r.sites?.site_name || 'All Site', status:r.status, aksi:'' }))
  const roles=['Platform Admin','App Admin','GL','Atasan Site','Driver','Viewer']
  return <div className="stack">
    <Panel title="Buat User Baru + Mapping Akses" desc="Form ini membuat akun Auth baru, profile, dan mapping akses sekaligus. Email tidak boleh sama dengan email admin yang sedang login."><form className="form-grid" onSubmit={createUserWithMapping} autoComplete="off"><label>Nama<input autoComplete="off" value={newUser.nama} onChange={e=>setNewUser({...newUser,nama:e.target.value})}/></label><label>NRP<input autoComplete="off" value={newUser.nrp} onChange={e=>setNewUser({...newUser,nrp:e.target.value})}/></label><label>Email User Baru<input type="email" name="new_user_email_no_autofill" autoComplete="new-email" value={newUser.email} onChange={e=>setNewUser({...newUser,email:e.target.value})}/></label><label>Password<input type="password" name="new_user_password_no_autofill" autoComplete="new-password" placeholder="Minimal 6 karakter" value={newUser.password} onChange={e=>setNewUser({...newUser,password:e.target.value})}/></label><label>Aplikasi<select value={newUser.app_id} onChange={e=>setNewUser({...newUser,app_id:e.target.value})}>{apps.map(a=><option key={a.id} value={a.id}>{a.app_name}</option>)}</select></label><label>Role<select value={newUser.role} onChange={e=>setNewUser({...newUser,role:e.target.value})}>{roles.map(r=><option key={r}>{r}</option>)}</select></label><label>Site<select value={newUser.site_id} onChange={e=>setNewUser({...newUser,site_id:e.target.value})}>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label><button disabled={loading}>{loading?'Memproses...':'Buat User Baru'}</button></form></Panel>
    <Panel title={editing?'Edit Mapping User Existing':'Tambah Mapping User Existing'} desc="Gunakan section ini untuk menambah role/site/aplikasi pada user yang sudah ada. Tidak membuat password baru."><form className="form-grid" onSubmit={saveExistingMapping}><UserExistingSearch profiles={profiles} value={existing.user_id} onPick={id=>setExisting({...existing,user_id:id})}/><label>Aplikasi<select value={existing.app_id} onChange={e=>setExisting({...existing,app_id:e.target.value})}>{apps.map(a=><option key={a.id} value={a.id}>{a.app_name}</option>)}</select></label><label>Role<select value={existing.role} onChange={e=>setExisting({...existing,role:e.target.value})}>{roles.map(r=><option key={r}>{r}</option>)}</select></label><label>Site<select value={existing.site_id} onChange={e=>setExisting({...existing,site_id:e.target.value})}>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label><button disabled={loading}>{editing?'Update Mapping':'Simpan Mapping'}</button>{editing && <button type="button" className="secondary" onClick={resetExisting}>Batal Edit</button>}</form>{message && <p className="message">{message}</p>}<p className="muted">Pembuatan password memakai Supabase Edge Function <code>admin-create-user</code>. Setelah update function, deploy ulang dengan <code>supabase functions deploy admin-create-user</code>.</p></Panel>
    <Panel title="Upload Excel User Mapping" desc="Upload untuk membuat user baru + mapping akses. Jika email sudah ada, sistem akan menambahkan mapping sesuai baris Excel."><div className="import-actions"><button className="secondary" onClick={downloadAccessTemplate}><Download size={16}/> Download Template</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewAccessUpload(e.target.files?.[0])}/></label>{bulkPreview.length>0 && <button disabled={loading || bulkPreview.some(r=>r.error)} onClick={submitAccessUpload}>Submit Import Valid</button>}</div>{bulkPreview.length>0 && <PreviewTable rows={bulkPreview.map(({app_id,site_id,user_id,password,...r})=>({...r,status_validasi:r.error ? 'ERROR: '+r.error : 'VALID'}))}/>}</Panel>
    <Panel title="Upload Mapping User Existing" desc="Upload khusus untuk menambahkan mapping aplikasi/role/site kepada user yang sudah ada. Tidak membuat user Auth baru dan tidak memakai password."><div className="import-actions"><button className="secondary" onClick={()=>downloadXlsx('template-mapping-user-existing.xlsx',[{ nama:'Nama User', nrp:'NRP001', email:'user@company.co.id', app_code:'sidak_fatigue', role:'GL', site_code:'BAYA', status:'Aktif' },{ nama:'Nama User', nrp:'NRP001', email:'user@company.co.id', app_code:'drd_driver', role:'GL', site_code:'BAYA', status:'Aktif' },{ nama:'Nama User', nrp:'NRP001', email:'user@company.co.id', app_code:'INSPEKSI', role:'GL', site_code:'BAYA', status:'Aktif' }])}><Download size={16}/> Download Template Tanpa Password</button><label className="upload-line"><FileSpreadsheet/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewExistingMappingUpload(e.target.files?.[0])}/></label>{existingBulkPreview.length>0 && <button disabled={loading || existingBulkPreview.some(r=>r.error)} onClick={submitExistingMappingUpload}>Submit Mapping Existing</button>}</div>{existingBulkPreview.length>0 && <PreviewTable rows={existingBulkPreview.map(({app_id,site_id,user_id,...r})=>({...r,status_validasi:r.error ? 'ERROR: '+r.error : 'VALID'}))}/>}</Panel>
    <Panel title="Row Data Mapping Akses" desc="Akses bisa diedit, dihapus, atau diaktif/nonaktifkan. User dengan beberapa site/role akan memilih kombinasi saat login." action={<button onClick={()=>downloadXlsx('mapping-akses.xlsx', rows)}><Download size={16}/> Export Excel</button>}><DataTable rows={rows} customActions={(idx)=>{const r=accessRows[idx]; return <div className="row-actions"><button className="secondary" onClick={()=>editAccess(r)}>Edit</button><button className="secondary" onClick={()=>toggleAccess(r)}>{r.status==='Aktif'?'Nonaktifkan':'Aktifkan'}</button><button className="secondary danger" onClick={()=>deleteAccess(r)}>Hapus</button></div>}} /></Panel>
  </div>
}
function PreviewTable({ rows }){
  const mapped = rows.map(r=>({ ...r, status_validasi: r.error ? `ERROR: ${r.error}` : 'VALID' }))
  return <div className="preview-box"><h4>Preview Import</h4><DataTable rows={mapped}/></div>
}

async function parseExcelOrCsv(file){
  const buf = await file.arrayBuffer()
  const workbook = XLSX.read(buf, { type:'array' })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  return XLSX.utils.sheet_to_json(sheet, { defval:'' })
}

function PlaceholderApp({ context }) {
  return <div className="stack">
    <Panel title={`${context.applications?.app_name} - Phase berikutnya`} desc="Core login, mapping role/site, dashboard row data, dan export sudah siap.">
      <p>Phase ini fokus menyelesaikan Sidak Fatigue end-to-end terlebih dahulu. Setelah stabil, pola yang sama dipakai untuk DRD Driver dan Inspeksi.</p>
    </Panel>
  </div>
}

function Kpi({ title, value, icon }) { return <div className="kpi"><div><small>{title}</small><strong>{value}</strong></div><div className="kpi-icon">{icon}</div></div> }
function Panel({ title, desc, action, children }) { return <div className="panel"><div className="panel-head"><div><h3>{title}</h3><p>{desc}</p></div>{action}</div>{children}</div> }
function StatusPill({ value }) {
  const red = ['Open','Tidak Aman','Rejected','Expired','Tidak Lulus'].includes(value)
  const green = ['Approved','Done','Lulus','Aman'].includes(value)
  const amber = ['Planned','Draft','Submitted','Closed','Belum Test','Proses'].includes(value)
  return <span className={`pill ${red?'red':green?'green':amber?'amber':'blue'}`}>{value}</span>
}
function DataTable({ rows, customActions, actions }) {
  const [q, setQ] = useState('')
  if (!rows || rows.length === 0) return <p className="muted">Belum ada data.</p>
  const headers = Object.keys(rows[0])
  const filtered = rows.filter(r => !q || Object.values(r).some(v => String(v ?? '').toLowerCase().includes(q.toLowerCase())))
  return <div className="data-table-block"><div className="table-search"><Search size={16}/><input placeholder="Search row data..." value={q} onChange={e=>setQ(e.target.value)} /></div><div className="table-wrap"><table><thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{filtered.map((r,i)=>{ const originalIdx = rows.indexOf(r); return <tr key={i}>{headers.map(h=><td key={h}>{h === 'aksi' && customActions ? customActions(originalIdx) : String(r[h] ?? '')}</td>)}</tr>})}</tbody></table></div>{filtered.length===0&&<p className="muted">Tidak ada data sesuai pencarian.</p>}</div>
}

function FullCenter({ text, action }) { return <div className="full-center"><ShieldCheck size={42}/><h2>{text}</h2>{action}</div> }

export default App
