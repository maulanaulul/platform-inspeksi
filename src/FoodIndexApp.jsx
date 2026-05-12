import React, { useEffect, useMemo, useState } from 'react'
import { supabase } from './lib/supabase'
import {
  AlertTriangle, BarChart3, Building2, CalendarCheck, CheckCircle2, ClipboardCheck,
  Download, FileSpreadsheet, LayoutDashboard, LogOut, Menu, Search, ShieldCheck,
  Store, Upload, Users, X
} from 'lucide-react'
import * as XLSX from 'xlsx'

const ADMIN_ROLES = ['Platform Admin', 'App Admin']
const ATASAN_ROLES = ['Atasan Site']
const GL_ROLES = ['GL', 'Site Admin', 'Admin Site']
function canAdmin(role){ return ADMIN_ROLES.includes(role) }
function canApprove(role){ return ADMIN_ROLES.includes(role) || ATASAN_ROLES.includes(role) }
function canInspect(role){ return ADMIN_ROLES.includes(role) || GL_ROLES.includes(role) }
function cleanText(v){ return String(v ?? '').trim() }
function today(){ return new Date().toISOString().slice(0, 10) }
function startOfWeekMonday(date = new Date()){
  const d = new Date(date)
  const day = d.getDay() || 7
  if (day !== 1) d.setDate(d.getDate() - day + 1)
  return d.toISOString().slice(0, 10)
}
function endOfWeekSunday(monday){
  const d = new Date(monday)
  d.setDate(d.getDate() + 6)
  return d.toISOString().slice(0, 10)
}
function statusClass(value=''){
  const v = String(value).toLowerCase()
  if (v.includes('approved') || v.includes('closed') || v.includes('aktif')) return 'ok'
  if (v.includes('expired') || v.includes('rejected') || v.includes('non')) return 'bad'
  if (v.includes('waiting') || v.includes('need') || v.includes('action plan')) return 'warn'
  return ''
}
function downloadXlsx(filename, rows){
  const ws = XLSX.utils.json_to_sheet(rows || [])
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Data')
  XLSX.writeFile(wb, filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`)
}
function normalizeRow(row){
  const out = {}
  Object.keys(row || {}).forEach(k => { out[String(k).trim()] = row[k] })
  return out
}
function getVal(row, keys){
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') return row[k]
  }
  return ''
}
function contextSiteName(context){ return context?.sites?.site_name || context?.sites?.site_code || 'All Site' }
function isGlobalView(context){ return canAdmin(context?.role) && !context?.site_id }
function adminCanSeeAll(context){ return canAdmin(context?.role) }

const MENUS = [
  ['dashboard', 'Dashboard', LayoutDashboard],
  ['vendors', 'Master Vendor Catering', Store],
  ['parameters', 'Master Parameter', FileSpreadsheet],
  ['tasks', 'Tasklist Inspeksi', ClipboardCheck],
  ['outstanding', 'Outstanding', AlertTriangle],
  ['approval', 'Approval', CheckCircle2],
  ['report', 'Report / Export', Download]
]

export default function FoodIndexApp({ embeddedProfile, embeddedContext, embeddedWork, onChangeApp }) {
  const profile = embeddedProfile
  const context = embeddedContext || embeddedWork
  const [page, setPage] = useState('dashboard')
  const [sidebar, setSidebar] = useState(false)

  const menu = MENUS.filter(([key]) => {
    if (['vendors','parameters'].includes(key)) return canAdmin(context?.role)
    if (key === 'approval') return canApprove(context?.role)
    if (key === 'tasks') return canInspect(context?.role)
    return true
  })

  useEffect(() => {
    if (!menu.find(([key]) => key === page)) setPage(menu[0]?.[0] || 'dashboard')
  }, [context?.id])

  const pageTitle = menu.find(([key]) => key === page)?.[1] || 'Food Index'
  const desc = {
    dashboard: 'Monitoring task mingguan, achievement, alert keterlambatan, dan outstanding Food Index.',
    vendors: 'Master vendor catering khusus Food Index. Site menggunakan master site platform existing.',
    parameters: 'Master parameter inspeksi Food Index. Parameter dapat ditambah manual atau upload Excel.',
    tasks: 'Tasklist inspeksi otomatis mingguan berdasarkan vendor catering aktif per site.',
    outstanding: 'Tindak lanjut temuan Food Index sampai close dan approval atasan site.',
    approval: 'Approval hasil inspeksi dan approval close outstanding.',
    report: 'Export data Food Index untuk kebutuhan monitoring dan laporan.'
  }

  return <div className="app-shell redesign-shell real-ui-shell">
    {sidebar && <button className="mobile-nav-backdrop" aria-label="Tutup navigasi" onClick={()=>setSidebar(false)} />}
    <aside className={sidebar ? 'sidebar open' : 'sidebar'}>
      <button className="sidebar-close" onClick={()=>setSidebar(false)}>×</button>
      <div className="brand dark sidebar-brand">
        <div className="logo">🍽️</div>
        <div><b>Food Index</b><span>{context?.role} · {contextSiteName(context)}</span></div>
      </div>
      <nav>{menu.map(([key,label,Icon]) => <button key={key} className={page===key?'active':''} onClick={()=>{setPage(key); setSidebar(false)}}><Icon size={18}/>{label}</button>)}</nav>
      <div className="sidebar-info-cards">
        <div className="sidebar-info-card"><b>{contextSiteName(context)}</b><span>Lokasi kerja aktif</span></div>
        <div className="sidebar-info-card"><b>Food Index</b><span>Inspeksi vendor catering mingguan</span></div>
      </div>
      <div className="sidebar-card"><b>Flow Mingguan</b><p>Task dibuat setiap Senin. Mulai Kamis muncul alert bila belum inspeksi. Senin berikutnya task lama menjadi expired.</p></div>
    </aside>
    <main className="main redesign-main">
      <header className="app-header">
        <button className="icon mobile" onClick={()=>setSidebar(!sidebar)}><Menu /></button>
        <div className="header-copy"><h2>{pageTitle}</h2><p>{desc[page] || 'Kelola Food Index.'}</p></div>
        <div className="header-actions redesign-actions">
          <div className="user-chip"><Users size={16}/><span>{profile?.nama || profile?.email} · {contextSiteName(context)}</span></div>
          <button className="secondary" onClick={onChangeApp}>Ganti Aplikasi</button>
          <button className="secondary" onClick={()=>supabase.auth.signOut()}><LogOut size={16}/> Logout</button>
        </div>
      </header>
      <section className="content">
        {page === 'dashboard' && <FoodDashboard context={context} />}
        {page === 'vendors' && <FoodVendors context={context} />}
        {page === 'parameters' && <FoodParameters context={context} />}
        {page === 'tasks' && <FoodTasks context={context} profile={profile} />}
        {page === 'outstanding' && <FoodOutstanding context={context} profile={profile} />}
        {page === 'approval' && <FoodApproval context={context} profile={profile} />}
        {page === 'report' && <FoodReport context={context} />}
      </section>
    </main>
  </div>
}

function Panel({ title, desc, action, children }){
  return <div className="panel"><div className="panel-head"><div><h3>{title}</h3>{desc && <p>{desc}</p>}</div>{action}</div>{children}</div>
}
function Kpi({ title, value, icon }){ return <div className="kpi"><div><small>{title}</small><strong>{value}</strong></div><div className="kpi-icon">{icon}</div></div> }
function StatusPill({ value }){ return <span className={`status-pill ${statusClass(value)}`}>{value || '-'}</span> }
function Table({ rows, columns, empty='Belum ada data.' }){
  const [q, setQ] = useState('')
  const cols = columns || (rows[0] ? Object.keys(rows[0]) : [])
  const filtered = rows.filter(r => !q || JSON.stringify(r).toLowerCase().includes(q.toLowerCase()))
  return <div>
    <div className="table-toolbar"><div className="searchbox"><Search size={18}/><input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search row data..." /></div></div>
    <div className="table-wrap" style={{maxHeight: 420, overflow: 'auto'}}><table><thead><tr>{cols.map(c => <th key={c}>{c}</th>)}</tr></thead><tbody>{filtered.map((r,i)=><tr key={r.id || i}>{cols.map(c => <td key={c}>{renderCell(r[c])}</td>)}</tr>)}</tbody></table>{!filtered.length && <p className="muted table-empty">{empty}</p>}</div>
  </div>
}
function renderCell(v){
  if (v === null || v === undefined || v === '') return '-'
  if (['Open','In Progress','Need Action Plan','Waiting Approval','Approved','Expired','Rejected','Closed','Aktif','Nonaktif'].includes(String(v))) return <StatusPill value={v} />
  return String(v)
}

async function fetchSites(){
  const { data, error } = await supabase.from('sites').select('id,site_code,site_name,status').order('site_code')
  if (error) throw error
  return data || []
}
function applySiteFilter(query, context, field='site_id'){
  if (adminCanSeeAll(context)) return query
  return query.eq(field, context.site_id)
}

function FoodDashboard({ context }){
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [siteFilter, setSiteFilter] = useState('')
  const [sites, setSites] = useState([])
  const [tasks, setTasks] = useState([])
  const [outs, setOuts] = useState([])
  const monday = startOfWeekMonday()
  const thursdayAlert = new Date().getDay() >= 4 || new Date().getDay() === 0

  useEffect(() => { load() }, [context?.id, siteFilter])
  async function load(){
    setLoading(true); setError('')
    try {
      if (adminCanSeeAll(context)) setSites(await fetchSites())
      let tq = supabase.from('food_weekly_tasks').select('*, sites(site_code,site_name), food_vendors(vendor_name)').eq('week_start_date', monday).order('generated_at', { ascending:false })
      let oq = supabase.from('food_outstandings').select('*, food_findings(corrective_action, preventive_action, food_inspection_answers(finding_note), food_weekly_tasks(sites(site_code,site_name), food_vendors(vendor_name)))').order('created_at', { ascending:false }).limit(500)
      tq = applySiteFilter(tq, context)
      if (!adminCanSeeAll(context)) oq = oq.eq('task_id', '00000000-0000-0000-0000-000000000000') // fallback if nested filter is not supported; loaded by tasks below in later versions
      if (adminCanSeeAll(context) && siteFilter) tq = tq.eq('site_id', siteFilter)
      const [{ data:t, error:te }, { data:o, error:oe }] = await Promise.all([tq, oq])
      if (te) throw te
      if (oe && adminCanSeeAll(context)) throw oe
      setTasks(t || [])
      setOuts(adminCanSeeAll(context) ? (o || []) : [])
    } catch(e){ setError(e.message) }
    setLoading(false)
  }

  const total = tasks.length
  const approved = tasks.filter(t => t.status === 'Approved').length
  const open = tasks.filter(t => ['Open','In Progress','Need Action Plan'].includes(t.status)).length
  const waiting = tasks.filter(t => t.status === 'Waiting Approval').length
  const expired = tasks.filter(t => t.status === 'Expired').length
  const achievement = total ? Math.round((approved / total) * 100) : 0
  const alertRows = tasks.filter(t => thursdayAlert && ['Open','In Progress','Need Action Plan'].includes(t.status)).map(t => ({ site:t.sites?.site_code, vendor:t.food_vendors?.vendor_name, status:t.status }))
  const taskRows = tasks.map(t => ({ site:t.sites?.site_code, vendor:t.food_vendors?.vendor_name, minggu:t.week_start_date, status:t.status, mulai:t.started_at?.slice(0,10), submit:t.submitted_at?.slice(0,10), approved:t.approved_at?.slice(0,10) }))

  return <div className="stack">
    {error && <div className="error">{error}</div>}
    <Panel title="Filter Dashboard" desc="Dashboard mengikuti minggu berjalan. Admin dapat filter site." action={adminCanSeeAll(context) && <select value={siteFilter} onChange={e=>setSiteFilter(e.target.value)}><option value="">Semua Site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select>}>
      <div className="summary-strip"><span>Minggu berjalan: <b>{monday}</b> s/d <b>{endOfWeekSunday(monday)}</b></span>{thursdayAlert && <span><AlertTriangle size={14}/> Alert Kamis aktif</span>}</div>
    </Panel>
    <div className="kpi-grid">
      <Kpi title="Task Minggu Ini" value={total} icon={<CalendarCheck/>} />
      <Kpi title="Approved" value={approved} icon={<CheckCircle2/>} />
      <Kpi title="Belum Selesai" value={open} icon={<AlertTriangle/>} />
      <Kpi title="Achievement" value={`${achievement}%`} icon={<BarChart3/>} />
      <Kpi title="Waiting Approval" value={waiting} icon={<ShieldCheck/>} />
      <Kpi title="Expired" value={expired} icon={<X/>} />
      <Kpi title="Outstanding Open" value={outs.filter(o=>o.status==='Open').length} icon={<AlertTriangle/>} />
    </div>
    <Panel title="Alert Site Belum Inspeksi" desc="Mulai Kamis, task yang belum selesai akan muncul sebagai alert dashboard admin.">
      {loading ? <p>Memuat...</p> : <Table rows={alertRows} columns={['site','vendor','status']} empty="Tidak ada alert keterlambatan." />}
    </Panel>
    <Panel title="Row Data Task Minggu Ini" desc="Task otomatis berdasarkan vendor catering aktif per site." action={<button onClick={()=>downloadXlsx('food-index-task-minggu-ini.xlsx', taskRows)}><Download size={16}/> Export</button>}>
      {loading ? <p>Memuat...</p> : <Table rows={taskRows} />}
    </Panel>
  </div>
}

function FoodVendors({ context }){
  const [sites, setSites] = useState([])
  const [vendors, setVendors] = useState([])
  const [msg, setMsg] = useState('')
  const [error, setError] = useState('')
  const [form, setForm] = useState({ site_id: context.site_id || '', vendor_name:'', status:'Aktif' })
  const [preview, setPreview] = useState([])

  useEffect(() => { load() }, [context?.id])
  async function load(){
    setError('')
    try {
      setSites(await fetchSites())
      let q = supabase.from('food_vendors').select('*, sites(site_code,site_name)').order('created_at', { ascending:false })
      q = applySiteFilter(q, context)
      const { data, error } = await q
      if (error) throw error
      setVendors(data || [])
    } catch(e){ setError(e.message) }
  }
  async function saveVendor(e){
    e.preventDefault(); setMsg(''); setError('')
    try {
      if (!form.site_id) throw new Error('Site wajib dipilih.')
      if (!cleanText(form.vendor_name)) throw new Error('Nama vendor wajib diisi.')
      const row = { site_id:form.site_id, vendor_name:cleanText(form.vendor_name), status:form.status || 'Aktif' }
      const { error } = await supabase.from('food_vendors').insert(row)
      if (error) throw error
      setForm({ site_id: context.site_id || '', vendor_name:'', status:'Aktif' })
      setMsg('Vendor catering berhasil disimpan.'); load()
    } catch(e){ setError(e.message) }
  }
  function downloadTemplate(){ downloadXlsx('template-master-vendor-catering.xlsx', [{ site_code:'BAYA', vendor_name:'Vendor Catering A', status:'Aktif' }]) }
  async function parseVendorExcel(file){
    setMsg(''); setError('')
    const buf = await file.arrayBuffer()
    const wb = XLSX.read(buf)
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]).map(normalizeRow)
    const siteMap = new Map((await fetchSites()).map(s => [String(s.site_code).toUpperCase(), s]))
    const mapped = rows.map((r, idx) => {
      const siteCode = cleanText(getVal(r, ['site_code','SITE_CODE','Site Code'])).toUpperCase()
      const site = siteMap.get(siteCode)
      const vendor_name = cleanText(getVal(r, ['vendor_name','VENDOR_NAME','nama_vendor','Nama Vendor']))
      const status = cleanText(getVal(r, ['status','STATUS'])) || 'Aktif'
      const err = !site ? 'site_code tidak ditemukan' : !vendor_name ? 'vendor_name wajib diisi' : ''
      return { row: idx+2, site_code: siteCode, site_id: site?.id, vendor_name, status, error: err }
    })
    setPreview(mapped)
  }
  async function submitImport(){
    setMsg(''); setError('')
    try {
      const valid = preview.filter(r => !r.error).map(r => ({ site_id:r.site_id, vendor_name:r.vendor_name, status:r.status }))
      if (!valid.length) throw new Error('Tidak ada data valid untuk diimport.')
      const { error } = await supabase.from('food_vendors').insert(valid)
      if (error) throw error
      setMsg(`Import selesai: ${valid.length} vendor catering tersimpan.`); setPreview([]); load()
    } catch(e){ setError(e.message) }
  }

  const rows = vendors.map(v => ({ site:v.sites?.site_code, vendor_code:v.vendor_code, vendor_name:v.vendor_name, status:v.status, created_at:v.created_at?.slice(0,10) }))
  return <div className="stack">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title="Tambah Vendor Catering" desc="Vendor catering khusus Food Index. Tidak memakai master vendor dari app lain.">
      <form className="form-grid" onSubmit={saveVendor}>
        <label>Site<select value={form.site_id} onChange={e=>setForm({...form, site_id:e.target.value})} disabled={!adminCanSeeAll(context)}><option value="">Pilih site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>
        <label>Nama Vendor Catering<input value={form.vendor_name} onChange={e=>setForm({...form, vendor_name:e.target.value})} placeholder="Nama vendor catering" /></label>
        <label>Status<select value={form.status} onChange={e=>setForm({...form, status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label>
        <button>Simpan Vendor</button>
      </form>
    </Panel>
    <Panel title="Import Vendor Catering Excel" desc="Kolom: site_code, vendor_name, status." action={<button className="secondary" onClick={downloadTemplate}><Download size={16}/> Download Template</button>}>
      <div className="import-row"><label className="upload-line"><Upload size={20}/><span>Upload Excel</span><input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0] && parseVendorExcel(e.target.files[0])} hidden /></label>{preview.length > 0 && <button onClick={submitImport}>Submit Import Valid</button>}</div>
      {preview.length > 0 && <Table rows={preview.map(r=>({ row:r.row, site_code:r.site_code, vendor_name:r.vendor_name, status:r.status, valid:r.error || 'Valid' }))} />}
    </Panel>
    <Panel title="Row Data Vendor Catering" action={<button onClick={()=>downloadXlsx('food-index-vendor-catering.xlsx', rows)}><Download size={16}/> Export</button>}>
      <Table rows={rows} />
    </Panel>
  </div>
}

function FoodParameters({ context }){
  const [rows, setRows] = useState([])
  const [form, setForm] = useState({ category:'General', parameter_text:'', sort_order:1, status:'Aktif' })
  const [preview, setPreview] = useState([])
  const [msg, setMsg] = useState('')
  const [error, setError] = useState('')
  useEffect(() => { load() }, [])
  async function load(){
    const { data, error } = await supabase.from('food_parameters').select('*').order('sort_order')
    if (error) setError(error.message)
    setRows(data || [])
  }
  async function save(e){
    e.preventDefault(); setMsg(''); setError('')
    try {
      if (!cleanText(form.parameter_text)) throw new Error('Parameter wajib diisi.')
      const { error } = await supabase.from('food_parameters').insert({ ...form, sort_order: Number(form.sort_order) || 1 })
      if (error) throw error
      setForm({ category:'General', parameter_text:'', sort_order:(rows.length+2), status:'Aktif' })
      setMsg('Parameter berhasil disimpan.'); load()
    } catch(e){ setError(e.message) }
  }
  function downloadTemplate(){ downloadXlsx('template-master-parameter-food-index.xlsx', [{ category:'Kebersihan', parameter_text:'Area penyajian bersih dan rapi', sort_order:1, status:'Aktif' }]) }
  async function parseExcel(file){
    const buf = await file.arrayBuffer()
    const wb = XLSX.read(buf)
    const items = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]).map(normalizeRow).map((r, idx) => {
      const parameter_text = cleanText(getVal(r, ['parameter_text','PARAMETER_TEXT','parameter','Parameter','soal','Soal']))
      return { row:idx+2, category:cleanText(getVal(r, ['category','CATEGORY','kategori','Kategori'])) || 'General', parameter_text, sort_order:Number(getVal(r, ['sort_order','SORT_ORDER','urutan','Urutan'])) || idx+1, status:cleanText(getVal(r, ['status','STATUS'])) || 'Aktif', error: parameter_text ? '' : 'parameter_text wajib diisi' }
    })
    setPreview(items)
  }
  async function submitImport(){
    setMsg(''); setError('')
    try {
      const valid = preview.filter(r => !r.error).map(({row,error,...r}) => r)
      if (!valid.length) throw new Error('Tidak ada parameter valid.')
      const { error } = await supabase.from('food_parameters').insert(valid)
      if (error) throw error
      setMsg(`Import selesai: ${valid.length} parameter tersimpan.`); setPreview([]); load()
    } catch(e){ setError(e.message) }
  }
  const tableRows = rows.map(r => ({ category:r.category, parameter_code:r.parameter_code, parameter_text:r.parameter_text, sort_order:r.sort_order, status:r.status }))
  return <div className="stack">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title="Tambah Parameter Manual" desc="Jawaban inspeksi nanti hanya 1 atau 0. Jika 0, evidence foto wajib diupload.">
      <form className="form-grid" onSubmit={save}>
        <label>Kategori<input value={form.category} onChange={e=>setForm({...form, category:e.target.value})} /></label>
        <label>Parameter<input value={form.parameter_text} onChange={e=>setForm({...form, parameter_text:e.target.value})} placeholder="Isi parameter inspeksi" /></label>
        <label>Urutan<input type="number" value={form.sort_order} onChange={e=>setForm({...form, sort_order:e.target.value})} /></label>
        <label>Status<select value={form.status} onChange={e=>setForm({...form, status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label>
        <button>Simpan Parameter</button>
      </form>
    </Panel>
    <Panel title="Upload Parameter Excel" desc="Kolom: category, parameter_text, sort_order, status." action={<button className="secondary" onClick={downloadTemplate}><Download size={16}/> Download Template</button>}>
      <div className="import-row"><label className="upload-line"><Upload size={20}/><span>Upload Excel</span><input type="file" accept=".xlsx,.xls" hidden onChange={e=>e.target.files?.[0] && parseExcel(e.target.files[0])}/></label>{preview.length > 0 && <button onClick={submitImport}>Submit Import Valid</button>}</div>
      {preview.length > 0 && <Table rows={preview.map(r=>({ row:r.row, category:r.category, parameter_text:r.parameter_text, sort_order:r.sort_order, status:r.status, valid:r.error || 'Valid' }))} />}
    </Panel>
    <Panel title="Row Data Parameter Food Index" action={<button onClick={()=>downloadXlsx('food-index-parameters.xlsx', tableRows)}><Download size={16}/> Export</button>}>
      <Table rows={tableRows} />
    </Panel>
  </div>
}

async function uploadFoodImage(file, folder='evidence'){
  if (!file) return ''
  const ext = (file.name?.split('.').pop() || 'jpg').toLowerCase()
  const safeExt = ['jpg','jpeg','png','webp'].includes(ext) ? ext : 'jpg'
  const path = `${folder}/${new Date().toISOString().slice(0,10)}/${crypto.randomUUID()}.${safeExt}`
  const { error } = await supabase.storage.from('food-index-assets').upload(path, file, { upsert:false, contentType:file.type || 'image/jpeg' })
  if (error) throw error
  const { data } = supabase.storage.from('food-index-assets').getPublicUrl(path)
  return data?.publicUrl || ''
}
async function ensureFoodTasksGenerated(){
  await supabase.rpc('expire_old_food_tasks')
  await supabase.rpc('generate_food_weekly_tasks')
}
function toTaskRow(t){
  return {
    id: t.id,
    site: t.sites?.site_code || '-',
    vendor: t.food_vendors?.vendor_name || '-',
    minggu: `${t.week_start_date} s/d ${t.week_end_date}`,
    status: t.status,
    started_at: t.started_at?.slice(0,16)?.replace('T',' ') || '-',
    submitted_at: t.submitted_at?.slice(0,16)?.replace('T',' ') || '-',
    approved_at: t.approved_at?.slice(0,16)?.replace('T',' ') || '-'
  }
}

function FoodTasks({ context, profile }){
  const [loading, setLoading] = useState(true)
  const [msg, setMsg] = useState('')
  const [error, setError] = useState('')
  const [tasks, setTasks] = useState([])
  const [parameters, setParameters] = useState([])
  const [activeTask, setActiveTask] = useState(null)
  const [answers, setAnswers] = useState({})
  const [submitting, setSubmitting] = useState(false)
  const monday = startOfWeekMonday()

  useEffect(() => { load() }, [context?.id])

  async function load(){
    setLoading(true); setError('')
    try {
      await ensureFoodTasksGenerated()
      const [{ data:p, error:pe }, taskRes] = await Promise.all([
        supabase.from('food_parameters').select('*').eq('status','Aktif').order('sort_order'),
        fetchTasks()
      ])
      if (pe) throw pe
      setParameters(p || [])
      setTasks(taskRes || [])
    } catch(e){ setError(e.message) }
    setLoading(false)
  }
  async function fetchTasks(){
    let q = supabase.from('food_weekly_tasks')
      .select('*, sites(site_code,site_name), food_vendors(vendor_name)')
      .eq('week_start_date', monday)
      .order('generated_at', { ascending:false })
    q = applySiteFilter(q, context)
    const { data, error } = await q
    if (error) throw error
    return data || []
  }
  async function openTask(task){
    setMsg(''); setError('')
    try {
      let currentTask = task
      if (task.status === 'Open') {
        const { data, error } = await supabase.from('food_weekly_tasks')
          .update({ status:'In Progress', started_at:new Date().toISOString(), gl_user_id:profile?.id, updated_at:new Date().toISOString() })
          .eq('id', task.id)
          .select('*, sites(site_code,site_name), food_vendors(vendor_name)')
          .single()
        if (error) throw error
        currentTask = data
      }
      const { data: existing, error: ae } = await supabase.from('food_inspection_answers')
        .select('*, food_findings(corrective_action, preventive_action)')
        .eq('task_id', currentTask.id)
      if (ae) throw ae
      const next = {}
      parameters.forEach(p => { next[p.id] = { score:'', note:'', evidenceUrl:'', evidenceFile:null } })
      ;(existing || []).forEach(a => {
        next[a.parameter_id] = {
          score: String(a.score),
          note: a.finding_note || '',
          evidenceUrl: a.evidence_photo_url || '',
          evidenceFile: null
        }
      })
      setAnswers(next)
      setActiveTask(currentTask)
      setTasks(await fetchTasks())
    } catch(e){ setError(e.message) }
  }
  function setAnswer(parameterId, patch){
    setAnswers(prev => ({ ...prev, [parameterId]: { ...(prev[parameterId] || {}), ...patch } }))
  }
  async function submitInspection(){
    setSubmitting(true); setMsg(''); setError('')
    try {
      if (!activeTask) throw new Error('Task belum dipilih.')
      if (!parameters.length) throw new Error('Master parameter aktif belum ada.')
      for (const p of parameters) {
        const a = answers[p.id] || {}
        if (a.score !== '0' && a.score !== '1') throw new Error(`Parameter wajib diisi: ${p.parameter_text}`)
        if (a.score === '0') {
          if (!cleanText(a.note)) throw new Error(`Catatan temuan wajib untuk parameter: ${p.parameter_text}`)
          if (!a.evidenceFile && !a.evidenceUrl) throw new Error(`Foto evidence wajib untuk temuan: ${p.parameter_text}`)
        }
      }
      const answerRows = []
      const findingPayload = []
      for (const p of parameters) {
        const a = answers[p.id]
        let evidenceUrl = a.evidenceUrl || ''
        if (a.score === '0' && a.evidenceFile) evidenceUrl = await uploadFoodImage(a.evidenceFile, 'evidence')
        answerRows.push({
          task_id: activeTask.id,
          parameter_id: p.id,
          score: Number(a.score),
          finding_note: a.score === '0' ? cleanText(a.note) : null,
          evidence_photo_url: a.score === '0' ? evidenceUrl : null,
          updated_at: new Date().toISOString()
        })
        if (a.score === '0') {
          findingPayload.push({ parameter_id:p.id })
        }
      }
      const { data:savedAnswers, error:upsertError } = await supabase.from('food_inspection_answers')
        .upsert(answerRows, { onConflict:'task_id,parameter_id' })
        .select('id, parameter_id')
      if (upsertError) throw upsertError

      const { error:deleteOutstandingError } = await supabase.from('food_outstandings').delete().eq('task_id', activeTask.id)
      if (deleteOutstandingError) throw deleteOutstandingError

      const { error:deleteFindingsError } = await supabase.from('food_findings').delete().eq('task_id', activeTask.id)
      if (deleteFindingsError) throw deleteFindingsError

      const savedMap = new Map((savedAnswers || []).map(a => [a.parameter_id, a.id]))
      if (findingPayload.length) {
        const findingRows = findingPayload.map(f => ({
          task_id: activeTask.id,
          answer_id: savedMap.get(f.parameter_id),
          corrective_action: null,
          preventive_action: null,
          validated_by: null,
          validated_at: null,
          status: 'Need Action Plan',
          updated_at: new Date().toISOString()
        })).filter(f => f.answer_id)
        const { data:findings, error:fe } = await supabase.from('food_findings').insert(findingRows).select('id, task_id')
        if (fe) throw fe
        const outRows = (findings || []).map(f => ({ finding_id:f.id, task_id:activeTask.id, status:'Open' }))
        if (outRows.length) {
          const { error:oe } = await supabase.from('food_outstandings').insert(outRows)
          if (oe) throw oe
        }
      }
      const nextTaskStatus = findingPayload.length ? 'Need Action Plan' : 'Waiting Approval'
      const { error:te } = await supabase.from('food_weekly_tasks')
        .update({ status:nextTaskStatus, submitted_at:new Date().toISOString(), gl_user_id:profile?.id, updated_at:new Date().toISOString() })
        .eq('id', activeTask.id)
      if (te) throw te
      setActiveTask(null)
      setMsg(findingPayload.length ? `Inspeksi tersubmit dengan ${findingPayload.length} temuan. Lanjut isi corrective & preventive di menu Outstanding sebelum approval Atasan Site.` : 'Inspeksi tersubmit tanpa temuan. Menunggu approval Atasan Site.')
      setTasks(await fetchTasks())
    } catch(e){ setError(e.message) }
    setSubmitting(false)
  }

  const rows = tasks.map(toTaskRow)
  const canOpen = t => ['Open','In Progress','Rejected'].includes(t.status)
  return <div className="stack">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title="Tasklist Mingguan Food Index" desc="Task otomatis dibuat berdasarkan vendor catering aktif di site. Klik Start untuk mulai inspeksi." action={<button className="secondary" onClick={load} disabled={loading}>{loading ? 'Memuat...' : 'Refresh / Generate Task'}</button>}>
      {loading ? <p>Memuat task...</p> : <div className="table-wrap" style={{maxHeight:500, overflow:'auto'}}><table><thead><tr><th>Site</th><th>Vendor</th><th>Minggu</th><th>Status</th><th>Mulai</th><th>Submit</th><th>Approval</th><th>Aksi</th></tr></thead><tbody>{tasks.map(t => <tr key={t.id}><td>{t.sites?.site_code || '-'}</td><td>{t.food_vendors?.vendor_name || '-'}</td><td>{t.week_start_date} s/d {t.week_end_date}</td><td>{renderCell(t.status)}</td><td>{t.started_at?.slice(0,16)?.replace('T',' ') || '-'}</td><td>{t.submitted_at?.slice(0,16)?.replace('T',' ') || '-'}</td><td>{t.approved_at?.slice(0,16)?.replace('T',' ') || '-'}</td><td>{canOpen(t) ? <button onClick={()=>openTask(t)}>{t.status === 'Open' ? 'Start Inspeksi' : 'Lanjut / Edit'}</button> : <span className="muted">-</span>}</td></tr>)}</tbody></table>{!tasks.length && <p className="muted table-empty">Belum ada task minggu ini. Pastikan Master Vendor Catering aktif sudah tersedia.</p>}</div>}
    </Panel>
    <Panel title="Export Tasklist" action={<button onClick={()=>downloadXlsx('food-index-tasklist.xlsx', rows)}><Download size={16}/> Export</button>}>
      <Table rows={rows} empty="Belum ada task untuk diexport." />
    </Panel>
    {activeTask && <div className="modal-backdrop"><div className="modal-card wide-modal">
      <div className="modal-head"><div><h3>Inspeksi Food Index</h3><p>{activeTask.sites?.site_code} · {activeTask.food_vendors?.vendor_name} · {activeTask.week_start_date}</p></div><button className="icon" onClick={()=>setActiveTask(null)}><X size={18}/></button></div>
      <div className="info-note"><b>Rules:</b> Pilihan hanya 1 atau 0. Jika 0, catatan temuan dan foto evidence wajib diisi. Corrective & preventive action diisi setelah submit melalui menu Outstanding.</div>
      <div className="table-wrap" style={{maxHeight:'62vh', overflow:'auto'}}><table><thead><tr><th style={{minWidth:80}}>Nilai</th><th style={{minWidth:280}}>Parameter</th><th>Evidence / Temuan</th></tr></thead><tbody>{parameters.map(p => {
        const a = answers[p.id] || {}
        return <tr key={p.id}><td><select value={a.score ?? ''} onChange={e=>setAnswer(p.id,{score:e.target.value})}><option value="">Pilih</option><option value="1">1</option><option value="0">0</option></select></td><td><b>{p.category}</b><br/>{p.parameter_text}</td><td>{a.score === '0' ? <div className="stack compact"><input placeholder="Catatan temuan wajib" value={a.note || ''} onChange={e=>setAnswer(p.id,{note:e.target.value})}/><label className="upload-line small"><Upload size={16}/><span>{a.evidenceFile?.name || (a.evidenceUrl ? 'Evidence sudah ada' : 'Upload foto evidence')}</span><input type="file" accept="image/*" hidden onChange={e=>setAnswer(p.id,{evidenceFile:e.target.files?.[0] || null})}/></label>{a.evidenceUrl && <a href={a.evidenceUrl} target="_blank" rel="noreferrer">Lihat evidence</a>}<span className="muted">Corrective & preventive action diisi di menu Outstanding.</span></div> : <span className="muted">Tidak perlu jika nilai 1</span>}</td></tr>
      })}</tbody></table></div>
      <div className="modal-actions"><button className="secondary" onClick={()=>setActiveTask(null)} disabled={submitting}>Batal</button><button onClick={submitInspection} disabled={submitting}>{submitting ? 'Menyimpan...' : 'Submit Inspeksi'}</button></div>
    </div></div>}
  </div>
}
function FoodOutstanding({ context, profile }){
  const [loading, setLoading] = useState(true)
  const [rows, setRows] = useState([])
  const [active, setActive] = useState(null)
  const [form, setForm] = useState({ corrective_action:'', preventive_action:'' })
  const [msg, setMsg] = useState('')
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)

  useEffect(() => { load() }, [context?.id])

  async function load(){
    setLoading(true); setError('')
    try {
      let q = supabase.from('food_outstandings')
        .select(`
          *,
          food_findings(
            id,
            corrective_action,
            preventive_action,
            status,
            validated_at,
            food_inspection_answers(
              finding_note,
              evidence_photo_url,
              food_parameters(category, parameter_text)
            )
          ),
          food_weekly_tasks!inner(
            id,
            status,
            week_start_date,
            site_id,
            vendor_id,
            sites(site_code, site_name),
            food_vendors(vendor_name)
          )
        `)
        .order('created_at', { ascending:false })
      if (!adminCanSeeAll(context)) q = q.eq('food_weekly_tasks.site_id', context.site_id)
      const { data, error } = await q
      if (error) throw error
      setRows(data || [])
    } catch(e){ setError(e.message) }
    setLoading(false)
  }

  function openActionPlan(row){
    setActive(row)
    setForm({
      corrective_action: row.food_findings?.corrective_action || '',
      preventive_action: row.food_findings?.preventive_action || ''
    })
    setMsg(''); setError('')
  }

  async function saveActionPlan(){
    if (!active) return
    setSaving(true); setMsg(''); setError('')
    try {
      if (!cleanText(form.corrective_action)) throw new Error('Corrective Action wajib diisi.')
      if (!cleanText(form.preventive_action)) throw new Error('Preventive Action wajib diisi.')
      const now = new Date().toISOString()
      const { error: fe } = await supabase.from('food_findings')
        .update({
          corrective_action: cleanText(form.corrective_action),
          preventive_action: cleanText(form.preventive_action),
          validated_by: profile?.id,
          validated_at: now,
          status: 'Action Plan Submitted',
          updated_at: now
        })
        .eq('id', active.finding_id)
      if (fe) throw fe

      const taskId = active.task_id
      const { data: findings, error: checkError } = await supabase.from('food_findings')
        .select('id, corrective_action, preventive_action')
        .eq('task_id', taskId)
      if (checkError) throw checkError
      const allReady = (findings || []).length > 0 && (findings || []).every(f => cleanText(f.corrective_action) && cleanText(f.preventive_action))
      if (allReady) {
        const { error: te } = await supabase.from('food_weekly_tasks')
          .update({ status:'Waiting Approval', updated_at:now })
          .eq('id', taskId)
        if (te) throw te
      }
      setActive(null)
      setMsg(allReady ? 'Corrective & preventive tersimpan. Inspeksi masuk Waiting Approval Atasan Site.' : 'Corrective & preventive tersimpan.')
      await load()
    } catch(e){ setError(e.message) }
    setSaving(false)
  }

  const tableRows = rows.map(r => ({
    id: r.id,
    site: r.food_weekly_tasks?.sites?.site_code || '-',
    vendor: r.food_weekly_tasks?.food_vendors?.vendor_name || '-',
    minggu: r.food_weekly_tasks?.week_start_date || '-',
    parameter: r.food_findings?.food_inspection_answers?.food_parameters?.parameter_text || '-',
    catatan_temuan: r.food_findings?.food_inspection_answers?.finding_note || '-',
    corrective_action: r.food_findings?.corrective_action || '-',
    preventive_action: r.food_findings?.preventive_action || '-',
    status_task: r.food_weekly_tasks?.status || '-',
    status_outstanding: r.status || '-'
  }))

  return <div className="stack">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title="Outstanding / Temuan Food Index" desc="Isi corrective dan preventive action dari temuan inspeksi. Setelah semua temuan pada task terisi, inspeksi masuk approval Atasan Site." action={<button className="secondary" onClick={load} disabled={loading}>{loading ? 'Memuat...' : 'Refresh'}</button>}>
      {loading ? <p>Memuat outstanding...</p> : <div className="table-wrap" style={{maxHeight:520, overflow:'auto'}}><table><thead><tr><th>Site</th><th>Vendor</th><th>Minggu</th><th>Parameter</th><th>Catatan</th><th>Evidence</th><th>Corrective</th><th>Preventive</th><th>Status Task</th><th>Aksi</th></tr></thead><tbody>{rows.map(r => {
        const answer = r.food_findings?.food_inspection_answers
        const hasPlan = cleanText(r.food_findings?.corrective_action) && cleanText(r.food_findings?.preventive_action)
        return <tr key={r.id}><td>{r.food_weekly_tasks?.sites?.site_code || '-'}</td><td>{r.food_weekly_tasks?.food_vendors?.vendor_name || '-'}</td><td>{r.food_weekly_tasks?.week_start_date || '-'}</td><td>{answer?.food_parameters?.parameter_text || '-'}</td><td>{answer?.finding_note || '-'}</td><td>{answer?.evidence_photo_url ? <a href={answer.evidence_photo_url} target="_blank" rel="noreferrer">Lihat foto</a> : '-'}</td><td>{r.food_findings?.corrective_action || '-'}</td><td>{r.food_findings?.preventive_action || '-'}</td><td>{renderCell(r.food_weekly_tasks?.status)}</td><td>{hasPlan ? <span className="muted">Sudah diisi</span> : <button onClick={()=>openActionPlan(r)}>Isi Tindakan</button>}</td></tr>
      })}</tbody></table>{!rows.length && <p className="muted table-empty">Belum ada outstanding / temuan.</p>}</div>}
    </Panel>
    <Panel title="Export Outstanding" action={<button onClick={()=>downloadXlsx('food-index-outstanding.xlsx', tableRows)}><Download size={16}/> Export</button>}>
      <Table rows={tableRows} empty="Belum ada data outstanding." />
    </Panel>
    {active && <div className="modal-backdrop"><div className="modal-card">
      <div className="modal-head"><div><h3>Isi Corrective & Preventive</h3><p>{active.food_weekly_tasks?.sites?.site_code} · {active.food_weekly_tasks?.food_vendors?.vendor_name}</p></div><button className="icon" onClick={()=>setActive(null)}><X size={18}/></button></div>
      <div className="stack compact">
        <label>Catatan Temuan<input value={active.food_findings?.food_inspection_answers?.finding_note || '-'} disabled /></label>
        <label>Corrective Action<textarea rows={4} value={form.corrective_action} onChange={e=>setForm({...form, corrective_action:e.target.value})} placeholder="Isi tindakan corrective" /></label>
        <label>Preventive Action<textarea rows={4} value={form.preventive_action} onChange={e=>setForm({...form, preventive_action:e.target.value})} placeholder="Isi tindakan preventive" /></label>
      </div>
      <div className="modal-actions"><button className="secondary" onClick={()=>setActive(null)} disabled={saving}>Batal</button><button onClick={saveActionPlan} disabled={saving}>{saving ? 'Menyimpan...' : 'Simpan & Kirim ke Approval'}</button></div>
    </div></div>}
  </div>
}
function FoodApproval(){ return <Placeholder title="Approval" desc="Tahap berikutnya: approval hasil inspeksi dan approval close outstanding oleh Atasan Site. Inspeksi dengan temuan baru muncul setelah corrective & preventive action diisi di menu Outstanding." /> }
function FoodReport(){ return <Placeholder title="Report / Export" desc="Tahap berikutnya: export task, temuan, outstanding, dan achievement Food Index." /> }
function Placeholder({ title, desc }){
  return <Panel title={title} desc={desc}><div className="card"><ClipboardCheck size={42}/><h3>{title} sedang disiapkan</h3><p>{desc}</p></div></Panel>
}
