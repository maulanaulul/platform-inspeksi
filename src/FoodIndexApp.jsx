import React, { useEffect, useMemo, useRef, useState } from 'react'
import { supabase } from './lib/supabase'
import {
  AlertTriangle, BarChart3, Building2, CalendarCheck, CheckCircle2, ClipboardCheck,
  Download, FileSpreadsheet, LayoutDashboard, LogOut, Menu, Search, ShieldCheck,
  Store, Upload, Users, X, Camera, Image as ImageIcon
} from 'lucide-react'
import * as XLSX from 'xlsx'

const ADMIN_ROLES = ['Platform Admin', 'App Admin']
const ATASAN_ROLES = ['Atasan Site']
const GL_ROLES = ['GL', 'Site Admin', 'Admin Site']
// Role Food Index yang tampil di UI. Nilai ini disimpan di food_role_level, bukan enum app_role.
const FOOD_ACCESS_ROLES = ['Administrator', 'GL', 'Section Head', 'Dept Head']
// Role enum app_role yang benar-benar valid di database. Jangan kirim role UI langsung ke kolom role.
const FOOD_DB_ACCESS_ROLES = ['Platform Admin', 'App Admin', 'GL', 'Atasan Site']
const FOOD_ROLE_TARGETS = { GL: 4, 'Section Head': 2, 'Dept Head': 1 }
// Site berikut disembunyikan dari semua perhitungan dan tabel Dashboard Food Index.
// Site tetap ada di master/admin untuk kebutuhan akses Head Office, tetapi tidak masuk achievement operasional.
const FOOD_DASHBOARD_HIDDEN_SITE_CODES = ['JIEP']
function canAdmin(role){ return ADMIN_ROLES.includes(role) }
function canApprove(role){ return ADMIN_ROLES.includes(role) || ATASAN_ROLES.includes(role) }
function canInspect(role){ return ADMIN_ROLES.includes(role) || GL_ROLES.includes(role) || ATASAN_ROLES.includes(role) }
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
  if (v.includes('approved') || v.includes('closed') || v === 'close' || v.includes('aktif')) return 'ok'
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
function isFoodDashboardHiddenSiteCode(siteCode){
  return FOOD_DASHBOARD_HIDDEN_SITE_CODES.includes(String(siteCode || '').trim().toUpperCase())
}
function isFoodDashboardVisibleSite(site){
  return !isFoodDashboardHiddenSiteCode(site?.site_code)
}

const MENUS = [
  ['dashboard', 'Dashboard', LayoutDashboard],
  ['admin', 'Admin Panel', Users],
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
  const [selfPasswordOpen, setSelfPasswordOpen] = useState(false)

  const menu = MENUS.filter(([key]) => {
    if (['admin','vendors','parameters'].includes(key)) return canAdmin(context?.role)
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
    admin: 'Kelola akses user Food Index, role, site, dan pembuatan password login.',
    vendors: 'Master vendor catering khusus Food Index. Site menggunakan master site platform existing.',
    parameters: 'Master parameter inspeksi Food Index. Parameter dapat ditambah manual atau upload Excel.',
    tasks: 'Tasklist inspeksi otomatis mingguan berdasarkan vendor catering aktif per site.',
    outstanding: 'Tindak lanjut temuan Food Index sampai close dan approval atasan site.',
    approval: 'Approval hasil inspeksi dan approval close outstanding.',
    report: 'Export data Food Index untuk kebutuhan monitoring dan laporan.'
  }

  return <div className="food-index-app app-shell redesign-shell real-ui-shell">
    <FoodIndexScopedStyles />
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
          <button className="secondary" onClick={()=>setSelfPasswordOpen(true)}>Ganti Password Saya</button>
          <button className="secondary" onClick={onChangeApp}>Ganti Aplikasi</button>
          <button className="secondary" onClick={()=>supabase.auth.signOut()}><LogOut size={16}/> Logout</button>
        </div>
      </header>
      <section className="content">
        {page === 'dashboard' && <FoodDashboard context={context} />}
        {page === 'admin' && <FoodAdminPanel context={context} profile={profile} />}
        {page === 'vendors' && <FoodVendors context={context} />}
        {page === 'parameters' && <FoodParameters context={context} />}
        {page === 'tasks' && <FoodTasks context={context} profile={profile} />}
        {page === 'outstanding' && <FoodOutstanding context={context} profile={profile} />}
        {page === 'approval' && <FoodApproval context={context} profile={profile} />}
        {page === 'report' && <FoodReport context={context} />}
      </section>
    </main>
    {selfPasswordOpen && <SelfPasswordModal profile={profile} onClose={()=>setSelfPasswordOpen(false)} />}
  </div>
}


function SelfPasswordModal({ profile, onClose }){
  const [form, setForm] = useState({ oldPassword:'', newPassword:'', confirmPassword:'' })
  const [saving, setSaving] = useState(false)
  const [msg, setMsg] = useState('')
  const [err, setErr] = useState('')
  async function submit(e){
    e.preventDefault()
    setErr(''); setMsg('')
    const email = cleanText(profile?.email)
    const oldPassword = cleanText(form.oldPassword)
    const newPassword = cleanText(form.newPassword)
    const confirmPassword = cleanText(form.confirmPassword)
    try {
      if (!email) throw new Error('Email akun tidak ditemukan. Silakan login ulang.')
      if (!oldPassword) throw new Error('Password lama wajib diisi.')
      if (!newPassword || newPassword.length < 6) throw new Error('Password baru minimal 6 karakter.')
      if (newPassword !== confirmPassword) throw new Error('Konfirmasi password baru tidak sama.')
      if (oldPassword === newPassword) throw new Error('Password baru tidak boleh sama dengan password lama.')
      setSaving(true)
      const { error: verifyError } = await supabase.auth.signInWithPassword({ email, password: oldPassword })
      if (verifyError) throw new Error('Password lama tidak sesuai.')
      const { error: updateError } = await supabase.auth.updateUser({ password: newPassword })
      if (updateError) throw updateError
      setMsg('Password berhasil diganti. Gunakan password baru pada login berikutnya.')
      setForm({ oldPassword:'', newPassword:'', confirmPassword:'' })
      setTimeout(()=>onClose?.(), 1200)
    } catch (e2) {
      setErr(e2.message || 'Gagal mengganti password.')
    } finally {
      setSaving(false)
    }
  }
  return <div className="modal-backdrop food-modal-backdrop self-password-backdrop" role="dialog" aria-modal="true">
    <div className="modal-card self-password-modal">
      <div className="modal-head compact-head">
        <div>
          <span className="modal-eyebrow">Akun Saya</span>
          <h3>Ganti Password Saya</h3>
          <p>Password akan diperbarui hanya untuk akun yang sedang login.</p>
        </div>
        <button className="secondary close-btn" onClick={onClose}>×</button>
      </div>
      <form className="self-password-form" onSubmit={submit}>
        {err && <div className="error">{err}</div>}
        {msg && <div className="success">{msg}</div>}
        <label>Email akun<input value={profile?.email || ''} disabled /></label>
        <label>Password lama<input type="password" autoComplete="current-password" value={form.oldPassword} onChange={e=>setForm({...form, oldPassword:e.target.value})} placeholder="Masukkan password lama" /></label>
        <label>Password baru<input type="password" autoComplete="new-password" value={form.newPassword} onChange={e=>setForm({...form, newPassword:e.target.value})} placeholder="Minimal 6 karakter" /></label>
        <label>Konfirmasi password baru<input type="password" autoComplete="new-password" value={form.confirmPassword} onChange={e=>setForm({...form, confirmPassword:e.target.value})} placeholder="Ulangi password baru" /></label>
        <div className="modal-actions inline-actions">
          <button type="button" className="secondary" onClick={onClose}>Batal</button>
          <button disabled={saving}>{saving ? 'Menyimpan...' : 'Simpan Password'}</button>
        </div>
      </form>
    </div>
  </div>
}

function Panel({ title, desc, action, children }){
  return <div className="panel"><div className="panel-head"><div><h3>{title}</h3>{desc && <p>{desc}</p>}</div>{action}</div>{children}</div>
}
function Kpi({ title, value, icon }){ return <div className="kpi"><div><small>{title}</small><strong>{value}</strong></div><div className="kpi-icon">{icon}</div></div> }
function StatusPill({ value }){ return <span className={`status-pill ${statusClass(value)}`}>{value || '-'}</span> }
function TaskStatusPill({ value }){ return <span className={`status-pill task-status-pill ${statusClass(value)}`}>{value || '-'}</span> }

function SignaturePad({ value, onChange }){
  const canvasRef = useRef(null)
  const drawingRef = useRef(false)
  const lastRef = useRef({ x:0, y:0 })

  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    const ctx = canvas.getContext('2d')
    ctx.lineCap = 'round'
    ctx.lineJoin = 'round'
    ctx.lineWidth = 2.6
    ctx.strokeStyle = '#0f172a'
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    if (value) {
      const img = new Image()
      img.onload = () => ctx.drawImage(img, 0, 0, canvas.width, canvas.height)
      img.src = value
    }
  }, [])

  function point(evt){
    const canvas = canvasRef.current
    const rect = canvas.getBoundingClientRect()
    const client = evt.touches?.[0] || evt
    return {
      x: (client.clientX - rect.left) * (canvas.width / rect.width),
      y: (client.clientY - rect.top) * (canvas.height / rect.height)
    }
  }
  function start(evt){
    evt.preventDefault()
    drawingRef.current = true
    lastRef.current = point(evt)
  }
  function move(evt){
    if (!drawingRef.current) return
    evt.preventDefault()
    const canvas = canvasRef.current
    const ctx = canvas.getContext('2d')
    const next = point(evt)
    ctx.beginPath()
    ctx.moveTo(lastRef.current.x, lastRef.current.y)
    ctx.lineTo(next.x, next.y)
    ctx.stroke()
    lastRef.current = next
    onChange?.(canvas.toDataURL('image/png'))
  }
  function end(){
    if (!drawingRef.current) return
    drawingRef.current = false
    const canvas = canvasRef.current
    onChange?.(canvas.toDataURL('image/png'))
  }
  function clear(){
    const canvas = canvasRef.current
    const ctx = canvas.getContext('2d')
    ctx.clearRect(0, 0, canvas.width, canvas.height)
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    onChange?.('')
  }
  return <div className="signature-pad-wrap">
    <canvas
      ref={canvasRef}
      width={620}
      height={220}
      className="signature-canvas"
      onMouseDown={start}
      onMouseMove={move}
      onMouseUp={end}
      onMouseLeave={end}
      onTouchStart={start}
      onTouchMove={move}
      onTouchEnd={end}
    />
    <div className="signature-actions"><span>Gunakan mouse/touchscreen untuk tanda tangan.</span><button type="button" className="secondary mini-btn" onClick={clear}>Clear TTD</button></div>
  </div>
}

function dataUrlToFile(dataUrl, filename){
  const arr = String(dataUrl || '').split(',')
  const mime = arr[0]?.match(/:(.*?);/)?.[1] || 'image/png'
  const binary = atob(arr[1] || '')
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return new File([bytes], filename, { type: mime })
}

function FoodIndexScopedStyles(){
  return <style>{`
    .food-index-app, .food-index-app * { font-family: Arial, Helvetica, sans-serif !important; }
    .food-index-app h2, .food-index-app h3, .food-index-app b, .food-index-app strong { letter-spacing: -0.02em; }
    .food-index-app h2 { font-weight: 700; }
    .food-index-app h3 { font-weight: 700; }
    .food-index-app .period-note { margin-top: 6px; font-size: 13px; font-weight: 700; color: #2563eb; }
    .food-index-app button { font-weight: 600; }
    .food-index-app input, .food-index-app select, .food-index-app textarea { font-family: Arial, Helvetica, sans-serif !important; font-weight: 500; }
    .food-index-app .panel, .food-index-app .kpi, .food-index-app .modal-card { border-radius: 22px; }
    .food-index-app .table-wrap { border-radius: 18px; }
    .food-index-app .food-modal-backdrop { align-items: flex-start; justify-content: center; padding: 22px; overflow-y: auto; overscroll-behavior: contain; }
    .food-index-app .food-inspection-modal { width: min(1060px, calc(100vw - 34px)); max-height: none; display: block; overflow: visible; padding: 0; margin: 0 auto 22px; }
    .food-index-app .inspection-entry-modal { height: calc(100dvh - 44px); max-height: calc(100dvh - 44px) !important; display: flex !important; flex-direction: column !important; overflow: hidden !important; margin: 0 auto !important; }
    .food-index-app .inspection-modal-body { flex: 1 1 auto; min-height: 0; overflow-y: auto; overflow-x: hidden; padding: 0; scroll-behavior: smooth; overscroll-behavior: contain; }
    .food-index-app .inspection-modal-body::-webkit-scrollbar { width: 10px; }
    .food-index-app .inspection-modal-body::-webkit-scrollbar-thumb { background:#94a3b8; border-radius:999px; }
    .food-index-app .inspection-entry-modal .food-modal-head, .food-index-app .inspection-entry-modal .sticky-actions { flex: 0 0 auto; }
    .food-index-app .inspection-entry-modal .sticky-actions { position: static !important; }
    .food-index-app .food-modal-head { padding: 26px 30px 18px; border-bottom: 1px solid #e8eef8; }

    .food-index-app .self-password-backdrop {
      position: fixed !important;
      inset: 0 !important;
      z-index: 9999 !important;
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      padding: 18px !important;
      background: rgba(15, 23, 42, .56) !important;
      backdrop-filter: blur(8px) !important;
      overflow: auto !important;
    }
    .food-index-app .self-password-modal {
      width: min(520px, calc(100vw - 32px));
      max-height: calc(100dvh - 36px);
      overflow: auto;
      margin: auto;
      position: relative;
      z-index: 10000;
      background: #fff;
    }
    .food-index-app .compact-head h3 { font-size: 24px; margin: 4px 0 6px; }
    .food-index-app .self-password-form { display:grid; gap: 12px; padding: 18px 22px 22px; }
    .food-index-app .self-password-form label { display:grid; gap: 7px; color:#334155; font-size: 13px; font-weight: 700; }
    .food-index-app .self-password-form input { width:100%; min-height:46px; border:1px solid #dbe7f7; border-radius: 14px; padding: 10px 12px; background:#fff; color:#0f172a; }
    .food-index-app .self-password-form input:disabled { background:#f8fafc; color:#64748b; }
    .food-index-app .inline-actions { display:flex; justify-content:flex-end; gap:10px; margin-top: 4px; }
    .food-index-app .food-modal-head h3 { margin: 4px 0 6px; font-size: 30px; line-height: 1.08; }
    .food-index-app .food-modal-head p { margin: 0; color: #667694; font-weight: 500; }
    .food-index-app .modal-eyebrow { color: #2563eb; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .08em; }
    .food-index-app .close-btn { width: 46px; height: 46px; border-radius: 14px; }
    .food-index-app .food-modal-backdrop::-webkit-scrollbar { width: 10px; }
    .food-index-app .food-modal-backdrop::-webkit-scrollbar-thumb { background:#94a3b8; border-radius:999px; }
    .food-index-app .inspection-summary { display: grid; grid-template-columns: 1fr 1fr 2.2fr; gap: 12px; padding: 16px 30px 0; }
    .food-index-app .inspection-summary > div { background: #f8fbff; border: 1px solid #dfe9f8; border-radius: 16px; padding: 12px 14px; min-height: 58px; }
    .food-index-app .inspection-summary small { display:block; color:#64748b; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing:.04em; }
    .food-index-app .inspection-summary b { color:#0f172a; font-size: 22px; }
    .food-index-app .inspection-summary span { display:block; margin-top:4px; color:#334155; font-weight: 600; }
    .food-index-app .inspection-note { margin: 14px 30px 0; border-radius: 14px; line-height: 1.45; }
    .food-index-app .inspection-scroll { overflow: visible; padding: 16px 30px 22px; display: grid; gap: 14px; }
    .food-index-app .inspection-item { background:#fff; border:1px solid #e2eaf6; border-radius: 18px; padding: 16px; box-shadow: 0 10px 28px rgba(15, 23, 42, .05); }
    .food-index-app .inspection-item.finding { border-color:#fecaca; background: linear-gradient(180deg, #fff, #fffafa); }
    .food-index-app .inspection-item.passed { border-color:#bbf7d0; }
    .food-index-app .inspection-param { display:flex; gap: 12px; align-items:flex-start; margin-bottom: 14px; }
    .food-index-app .param-index { width:34px; height:34px; border-radius: 12px; display:inline-flex; align-items:center; justify-content:center; background:#eff6ff; color:#1d4ed8; font-weight:700; flex: 0 0 auto; }
    .food-index-app .inspection-param small { display:block; color:#64748b; font-weight:700; margin-bottom: 4px; }
    .food-index-app .inspection-param strong { display:block; font-size: 17px; color:#0f172a; line-height:1.35; }
    .food-index-app .score-options { display:grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; margin-bottom: 12px; }
    .food-index-app .score-option { border:1px solid #dbe7f7; background:#f8fbff; color:#0f172a; border-radius: 16px; padding: 14px; display:flex; align-items:center; justify-content:center; gap: 10px; min-height: 58px; box-shadow:none; }
    .food-index-app .score-option b { width:32px; height:32px; border-radius: 999px; display:inline-flex; align-items:center; justify-content:center; background:#e2e8f0; color:#0f172a; }
    .food-index-app .score-option span { font-weight:700; }
    .food-index-app .score-option.active.good { border-color:#22c55e; background:#ecfdf5; color:#166534; }
    .food-index-app .score-option.active.good b { background:#22c55e; color:white; }
    .food-index-app .score-option.active.bad { border-color:#ef4444; background:#fef2f2; color:#991b1b; }
    .food-index-app .score-option.active.bad b { background:#ef4444; color:white; }
    .food-index-app .finding-fields { display:grid; grid-template-columns: minmax(260px, 1.3fr) minmax(220px, .9fr); gap: 12px; align-items: stretch; }
    .food-index-app .field-block span { display:block; color:#334155; font-size:13px; font-weight:700; margin-bottom:8px; }
    .food-index-app .field-block em { color:#dc2626; font-style:normal; }
    .food-index-app .field-block textarea { width:100%; min-height: 106px; resize: vertical; border-radius: 14px; border:1px solid #dbe7f7; padding: 12px; line-height:1.45; }
    .food-index-app .upload-card { border:1px dashed #93c5fd; background:#eff6ff; border-radius: 16px; padding: 16px; display:flex; flex-direction:column; justify-content:center; align-items:center; gap: 6px; text-align:center; cursor:pointer; color:#1d4ed8; min-height: 106px; }
    .food-index-app .upload-card span { color:#64748b; font-size: 12px; line-height:1.35; }
    .food-index-app .evidence-link { display:inline-flex; align-items:center; color:#2563eb; font-weight:700; margin-top: 8px; }
    .food-index-app .no-evidence-note { background:#f8fafc; border:1px dashed #e2e8f0; border-radius: 14px; color:#64748b; padding: 12px 14px; font-size: 14px; }
    .food-index-app .signature-section { margin: 0 30px 20px; padding: 18px; border: 1px solid #dbe7f7; border-radius: 20px; background: #f8fbff; max-height: none; overflow: visible; }
    .food-index-app .signature-section-head { display:flex; align-items:flex-start; justify-content:space-between; gap: 14px; margin-bottom: 14px; }
    .food-index-app .signature-section-head h4 { margin:0 0 6px; font-size:20px; color:#0f172a; }
    .food-index-app .signature-section-head p { margin:0; color:#64748b; line-height:1.4; }
    .food-index-app .signature-mode-tabs { display:flex; flex-wrap:wrap; gap: 10px; margin-bottom: 14px; }
    .food-index-app .signature-mode-tabs button { background:#fff !important; color:#1d4ed8; border:1px solid #bfdbfe; box-shadow:none; padding: 10px 14px; border-radius: 999px; }
    .food-index-app .signature-mode-tabs button.active { background:#2563eb !important; color:#fff; border-color:#2563eb; }
    .food-index-app .signature-search-box { display:grid; grid-template-columns: minmax(260px,1fr) auto; gap: 10px; align-items:end; margin-bottom: 12px; }
    .food-index-app .signature-search-box label, .food-index-app .signature-form-grid label { display:block; color:#334155; font-size:13px; font-weight:700; }
    .food-index-app .signature-search-box input, .food-index-app .signature-form-grid input { width:100%; margin-top:7px; border:1px solid #dbe7f7; border-radius: 14px; padding: 12px; min-height: 46px; background:#fff; }
    .food-index-app .driver-results { grid-column: 1 / -1; display:grid; gap: 8px; max-height: 190px; overflow:auto; }
    .food-index-app .driver-results button { background:#fff !important; color:#0f172a; box-shadow:none; border:1px solid #dbe7f7; justify-content:flex-start; text-align:left; display:flex; flex-direction:column; align-items:flex-start; gap: 3px; padding: 10px 12px; }
    .food-index-app .driver-results button span { color:#64748b; font-size:12px; }
    .food-index-app .signature-form-grid { display:grid; grid-template-columns: repeat(4, minmax(150px, 1fr)); gap: 10px; margin-bottom: 12px; }
    .food-index-app .signature-pad-wrap { background:#fff; border:1px solid #dbe7f7; border-radius: 18px; padding: 12px; }
    .food-index-app .signature-empty-hint { background:#fff; border:1px dashed #bfdbfe; border-radius:18px; padding:20px; text-align:center; color:#64748b; font-weight:700; margin-bottom:12px; }
    .food-index-app .signature-canvas { width:100%; height: 190px; display:block; border-radius: 14px; border:1px dashed #93c5fd; background:#fff; touch-action:none; }
    .food-index-app .signature-actions { display:flex; justify-content:space-between; gap: 10px; align-items:center; margin-top: 8px; color:#64748b; font-size: 12px; }
    .food-index-app .mini-btn { padding: 8px 12px; min-height: auto; font-size: 12px; }
    .food-index-app .signature-add-row { display:flex; justify-content:flex-end; margin: 12px 0; }
    .food-index-app .signature-list { display:grid; gap: 10px; }
    .food-index-app .signature-list-item { display:grid; grid-template-columns: 1fr 160px auto; gap: 12px; align-items:center; background:#fff; border:1px solid #dbe7f7; border-radius: 16px; padding: 10px 12px; }
    .food-index-app .signature-list-item b { display:block; color:#0f172a; }
    .food-index-app .signature-list-item span { display:block; color:#64748b; font-size:12px; margin-top:3px; }
    .food-index-app .signature-list-item img { max-height:64px; max-width: 150px; object-fit:contain; border:1px solid #e2e8f0; border-radius: 10px; background:#fff; }
    .food-index-app .signature-list-item a { color:#2563eb; font-weight:700; }
    .food-index-app .approval-signature-box { border:1px solid #dbe7f7; background:#f8fbff; border-radius:18px; padding:16px; margin-top: 8px; }
    @media (max-width: 760px) {
      .food-index-app .signature-section { margin-left: 18px; margin-right: 18px; }
      .food-index-app .signature-search-box, .food-index-app .signature-form-grid, .food-index-app .signature-list-item { grid-template-columns: 1fr; }
      .food-index-app .signature-canvas { height: 160px; }
    }

    .food-index-app .approval-detail-grid { display:grid; grid-template-columns: repeat(5, minmax(160px, 1fr)); gap: 12px; }
    .food-index-app .approval-detail-grid > div { background:#f8fafc; border:1px solid #e2e8f0; border-radius: 14px; padding: 12px; min-height: 74px; }
    .food-index-app .approval-detail-grid small { display:block; color:#64748b; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; margin-bottom: 6px; }
    .food-index-app .approval-detail-grid b, .food-index-app .approval-detail-grid span, .food-index-app .approval-detail-grid a { font-weight:700; color:#0f172a; line-height:1.35; }
    .food-index-app .approval-detail-grid a { color:#2563eb; }
    .food-index-app button.danger { background:#ef4444; color:white; border-color:#ef4444; }
    .food-index-app button.danger:hover { background:#dc2626; }
    .food-index-app .sticky-actions { position: sticky; bottom:0; background:#fff; border-top:1px solid #e8eef8; padding: 16px 30px 22px; margin: 0; }
    @media (max-width: 760px) {
      .food-index-app .inspection-summary, .food-index-app .finding-fields, .food-index-app .approval-detail-grid { grid-template-columns: 1fr; }
      .food-index-app .food-modal-head, .food-index-app .inspection-summary, .food-index-app .inspection-note, .food-index-app .inspection-scroll, .food-index-app .sticky-actions { padding-left: 18px; padding-right: 18px; margin-left:0; margin-right:0; }
      .food-index-app .score-options { grid-template-columns: 1fr; }
    }

    /* v40 modern polish */
    .food-index-app { background: radial-gradient(circle at top left, #eef6ff 0, #f7fbff 28%, #eef4fb 100%); color:#0f172a; }
    .food-index-app .main { background: transparent; }
    .food-index-app .sidebar { background: linear-gradient(180deg, #071a55 0%, #0d2e83 52%, #08205f 100%); box-shadow: 12px 0 34px rgba(15, 23, 42, .16); }
    .food-index-app .sidebar .logo { background: linear-gradient(135deg, #4f7cff, #7c3aed); box-shadow: 0 18px 42px rgba(37, 99, 235, .34); }
    .food-index-app .sidebar nav button { border-radius: 18px; margin: 5px 14px; transition: transform .16s ease, background .16s ease; }
    .food-index-app .sidebar nav button:hover { transform: translateX(3px); background: rgba(255,255,255,.10); }
    .food-index-app .sidebar nav button.active { background: linear-gradient(135deg, #2f74ff, #2857d8); box-shadow: 0 16px 36px rgba(37, 99, 235, .38); }
    .food-index-app .app-header { background: rgba(255,255,255,.72); backdrop-filter: blur(14px); border-bottom: 1px solid rgba(148, 163, 184, .22); }
    .food-index-app .header-copy h2 { font-size: clamp(36px, 4vw, 58px); line-height: .95; }
    .food-index-app .header-copy p { color:#64748b; font-weight: 500; max-width: 760px; }
    .food-index-app .user-chip, .food-index-app .secondary, .food-index-app button { border-radius: 16px; }
    .food-index-app button:not(.secondary):not(.icon):not(.score-option):not(.sidebar-close) { background: linear-gradient(135deg, #2563eb, #1d4ed8); box-shadow: 0 12px 26px rgba(37, 99, 235, .22); }
    .food-index-app button:not(.secondary):not(.icon):not(.score-option):not(.sidebar-close):hover { filter: brightness(.98); transform: translateY(-1px); }
    .food-index-app .content { padding-top: 26px; }
    .food-index-app .panel { background: rgba(255,255,255,.90); border: 1px solid rgba(191, 219, 254, .75); box-shadow: 0 24px 60px rgba(15, 23, 42, .08); overflow:hidden; }
    .food-index-app .panel-head { gap: 16px; align-items:flex-start; }
    .food-index-app .panel-head h3 { font-size: 28px; }
    .food-index-app .panel-head p { color:#475569; font-weight:500; }
    .food-index-app .kpi { background: linear-gradient(145deg, #ffffff 0%, #f8fbff 100%); border: 1px solid rgba(191, 219, 254, .8); box-shadow: 0 18px 48px rgba(15, 23, 42, .07); }
    .food-index-app .kpi strong { font-size: 36px; }
    .food-index-app .kpi-icon { background: linear-gradient(135deg, #eff6ff, #dbeafe); color:#2563eb; }
    .food-index-app .table-wrap { background:#fff; border: 1px solid #dbe7f7; box-shadow: inset 0 1px 0 rgba(255,255,255,.8); position:relative; }
    .food-index-app .table-wrap table { border-collapse: separate; border-spacing: 0; }
    .food-index-app table th { background: linear-gradient(180deg, #f8fbff, #eef6ff); color:#334155; font-size: 12px; letter-spacing:.08em; }
    .food-index-app table td { vertical-align: top; }
    .food-index-app .table-wrap thead th { position: sticky; top: 0; z-index: 8; background: linear-gradient(180deg, #f8fbff 0%, #eef6ff 100%); box-shadow: 0 1px 0 #dbe7f7, 0 8px 18px rgba(15,23,42,.06); }
    .food-index-app .status-pill { border-radius: 999px; padding: 7px 11px; font-weight: 700; display:inline-flex; align-items:center; justify-content:center; width:max-content; max-width:100%; white-space:nowrap; line-height:1.15; box-sizing:border-box; }
    .food-index-app .status-pill.ok { background:#dcfce7; color:#166534; }
    .food-index-app .status-pill.warn { background:#fef3c7; color:#92400e; }
    .food-index-app .status-pill.bad { background:#fee2e2; color:#991b1b; }
    .food-index-app td.remark-cell { font-weight:700; text-align:center; vertical-align:middle !important; box-shadow: inset 0 1px 0 rgba(255,255,255,.65); }
    .food-index-app td.remark-cell-success { background:#dcfce7 !important; color:#166534 !important; border-left:4px solid #22c55e; }
    .food-index-app td.remark-cell-danger { background:#fee2e2 !important; color:#991b1b !important; border-left:4px solid #ef4444; }
    .food-index-app td.remark-cell-warning { background:#fef3c7 !important; color:#92400e !important; border-left:4px solid #f59e0b; }
    .food-index-app td.remark-cell-neutral { background:#f8fafc !important; color:#475569 !important; border-left:4px solid #94a3b8; }
    .food-index-app .report-table-wrap table { min-width: 1480px; }
    .food-index-app .modal-card { box-shadow: 0 34px 90px rgba(15,23,42,.24); border:1px solid rgba(219,234,254,.95); }
    .food-index-app .modal-head { background: linear-gradient(180deg, #ffffff, #f8fbff); }
    .food-index-app .info-note { background:#f8fbff; border:1px solid #dbeafe; border-radius:16px; padding:14px 16px; color:#334155; line-height:1.55; }
    .food-index-app .upload-line { border:1px dashed #93c5fd; border-radius:16px; background:#eff6ff; color:#1d4ed8; padding:16px; display:flex; gap:10px; align-items:center; cursor:pointer; font-weight:700; }


    /* v43 dashboard design based on the approved dark-sidebar reference */
    .food-index-app.real-ui-shell {
      min-height: 100vh;
      background:
        radial-gradient(circle at 74% 0%, rgba(96,165,250,.32), transparent 28%),
        linear-gradient(135deg, #e7f7ff 0%, #eef6ff 36%, #f8fbff 100%);
      color: #071327;
    }
    .food-index-app .sidebar {
      background:
        linear-gradient(rgba(255,255,255,.045) 1px, transparent 1px),
        linear-gradient(90deg, rgba(255,255,255,.045) 1px, transparent 1px),
        linear-gradient(180deg, #061747 0%, #0d2d7a 50%, #071e5d 100%) !important;
      background-size: 46px 46px, 46px 46px, auto !important;
      border-right: 1px solid rgba(255,255,255,.10) !important;
      box-shadow: 16px 0 44px rgba(8, 23, 70, .22) !important;
    }
    .food-index-app .sidebar .brand,
    .food-index-app .sidebar-brand {
      margin: 26px 20px 22px !important;
      padding: 20px !important;
      border-radius: 26px !important;
      background: rgba(255,255,255,.11) !important;
      border: 1px solid rgba(255,255,255,.16) !important;
      box-shadow: inset 0 1px 0 rgba(255,255,255,.08), 0 18px 42px rgba(0,0,0,.10) !important;
      backdrop-filter: blur(14px);
    }
    .food-index-app .sidebar .logo {
      width: 56px !important;
      height: 56px !important;
      border-radius: 20px !important;
      background: linear-gradient(135deg, #2f74ff, #16c2f3) !important;
      box-shadow: 0 16px 34px rgba(22,194,243,.24) !important;
    }
    .food-index-app .sidebar .brand b { color:#fff; font-size: 26px; font-weight: 700; }
    .food-index-app .sidebar .brand span { color: rgba(255,255,255,.74); font-size: 16px; line-height: 1.42; margin-top: 8px; display:block; }
    .food-index-app .sidebar nav { padding: 0 18px !important; gap: 14px !important; }
    .food-index-app .sidebar nav button {
      min-height: 70px !important;
      border-radius: 24px !important;
      margin: 0 !important;
      padding: 0 18px !important;
      color: rgba(255,255,255,.78) !important;
      background: transparent !important;
      box-shadow: none !important;
      font-size: 17px !important;
      font-weight: 700 !important;
      justify-content: flex-start !important;
      gap: 18px !important;
    }
    .food-index-app .sidebar nav button svg {
      width: 40px;
      height: 40px;
      padding: 10px;
      border-radius: 16px;
      background: rgba(255,255,255,.10);
      color: rgba(255,255,255,.72);
    }
    .food-index-app .sidebar nav button.active {
      background: linear-gradient(135deg, #2b6ff5, #0ea5e9) !important;
      color: #fff !important;
      box-shadow: 0 18px 42px rgba(14,165,233,.32) !important;
    }
    .food-index-app .sidebar nav button.active svg { color:#fff; background: rgba(255,255,255,.16); }
    .food-index-app .sidebar-info-cards { display:none !important; }
    .food-index-app .sidebar-card {
      left: 20px !important;
      right: 20px !important;
      bottom: 26px !important;
      padding: 22px !important;
      border-radius: 26px !important;
      background: rgba(255,255,255,.13) !important;
      border: 1px solid rgba(255,255,255,.18) !important;
      color:#fff !important;
    }
    .food-index-app .sidebar-card b { font-size: 22px !important; color:#fff !important; }
    .food-index-app .sidebar-card p { color: rgba(255,255,255,.74) !important; font-size: 16px !important; line-height: 1.5 !important; }
    .food-index-app .redesign-main { background: transparent !important; padding: 38px 42px !important; }
    .food-index-app .app-header {
      background: transparent !important;
      border: 0 !important;
      box-shadow: none !important;
      padding: 0 !important;
      display: grid !important;
      grid-template-columns: minmax(0,1fr) auto !important;
      align-items: start !important;
      gap: 24px !important;
      margin-bottom: 26px !important;
    }
    .food-index-app .header-copy h2 {
      font-size: clamp(44px, 5vw, 68px) !important;
      line-height: .95 !important;
      letter-spacing: -3px !important;
      color: #071327 !important;
      margin: 0 !important;
    }
    .food-index-app .header-copy p {
      color:#61708a !important;
      font-size: 21px !important;
      line-height: 1.5 !important;
      max-width: 770px !important;
      margin-top: 18px !important;
    }
    .food-index-app .redesign-actions {
      display: grid !important;
      grid-template-columns: 1fr 1fr !important;
      gap: 14px !important;
      min-width: 440px !important;
      justify-content: end !important;
    }
    .food-index-app .user-chip {
      grid-column: 1 / -1;
      min-height: 64px !important;
      border-radius: 22px !important;
      background: rgba(255,255,255,.90) !important;
      border: 1px solid rgba(209,224,246,.95) !important;
      box-shadow: 0 18px 44px rgba(37,99,235,.11) !important;
      justify-content: center !important;
      font-size: 18px !important;
      font-weight: 700 !important;
    }
    .food-index-app .secondary {
      min-height: 64px !important;
      border-radius: 22px !important;
      background: rgba(255,255,255,.90) !important;
      border: 1px solid rgba(209,224,246,.95) !important;
      box-shadow: 0 18px 44px rgba(37,99,235,.11) !important;
      font-size: 18px !important;
      color:#071327 !important;
      font-weight: 700 !important;
    }
    .food-index-app .content { padding: 0 !important; }
    .food-index-app .food-dashboard-v43 { display: grid; gap: 24px; }
    .food-index-app .dashboard-filter-card,
    .food-index-app .dashboard-card,
    .food-index-app .dash-kpi {
      background: rgba(255,255,255,.86);
      border: 1px solid #d8e8ff;
      box-shadow: 0 24px 60px rgba(37,99,235,.11);
      border-radius: 30px;
    }
    .food-index-app .dashboard-filter-card { padding: 26px; display:grid; grid-template-columns: 1fr auto; gap: 20px; align-items:start; }
    .food-index-app .dashboard-filter-card small { color:#64748b; font-weight: 700; display:block; margin-bottom:8px; font-size: 15px; }
    .food-index-app .dashboard-filter-card h3 { margin:0; font-size: 26px; letter-spacing:-.8px; }
    .food-index-app .dashboard-filter-card p { margin: 10px 0 0; color:#64748b; font-size: 16px; line-height:1.45; }
    .food-index-app .dashboard-filter-actions { display:flex; align-items:end; gap: 12px; flex-wrap: wrap; justify-content:flex-end; }
    .food-index-app .dashboard-filter-actions label { color:#334155; font-weight:700; display:grid; gap:8px; }
    .food-index-app .dashboard-filter-actions select { height:54px; min-width:240px; border-radius:18px; border:1px solid #bfd8ff; background:#fff; padding:0 14px; font-weight:700; }
    .food-index-app .dashboard-filter-actions button,
    .food-index-app .dashboard-card-head button,
    .food-index-app .task-row-card button {
      min-height:54px !important;
      border-radius:18px !important;
      padding: 0 22px !important;
      background: linear-gradient(135deg, #2563eb, #1d4ed8) !important;
      color:#fff !important;
      border:0 !important;
      box-shadow: 0 18px 36px rgba(37,99,235,.24) !important;
      font-size:16px !important;
      display:inline-flex !important;
      align-items:center !important;
      gap:8px !important;
    }
    .food-index-app .dashboard-tags { grid-column: 1 / -1; display:flex; gap:10px; flex-wrap:wrap; }
    .food-index-app .dashboard-tags span { padding: 8px 12px; border-radius:999px; color:#1d4ed8; background:#eaf4ff; font-size:13px; font-weight:700; }
    .food-index-app .dashboard-kpi-grid { display:grid; grid-template-columns: repeat(4, minmax(170px, 1fr)); gap: 24px; }
    .food-index-app .dash-kpi { min-height: 190px; padding: 28px; display:flex; flex-direction:column; justify-content:center; }
    .food-index-app .dash-kpi small { color:#667694; font-size:18px; font-weight:700; }
    .food-index-app .dash-kpi strong { display:block; margin-top:18px; font-size:44px; line-height:1; letter-spacing:-2px; }
    .food-index-app .dash-kpi span { display:inline-flex; align-self:flex-start; margin-top:22px; padding:10px 14px; border-radius:999px; background:#dcfce7; color:#15803d; font-size:15px; font-weight:700; }
    .food-index-app .dash-kpi span.orange { background:#ffedd5; color:#c2410c; }
    .food-index-app .dash-kpi span.red { background:#fee2e2; color:#b91c1c; }
    .food-index-app .dashboard-main-grid { display:grid; grid-template-columns: minmax(0, 1.7fr) minmax(320px, .95fr); gap: 24px; align-items:stretch; }
    .food-index-app .dashboard-card { padding: 28px; overflow:hidden; }
    .food-index-app .dashboard-card-head { display:flex; justify-content:space-between; align-items:flex-start; gap:16px; margin-bottom:20px; }
    .food-index-app .dashboard-card h3 { margin:0; font-size:32px; letter-spacing:-1.2px; }
    .food-index-app .dashboard-card p { color:#64748b; font-size:18px; line-height:1.45; margin:10px 0 0; }
    .food-index-app .report-filter-row { display:flex; gap:14px; flex-wrap:wrap; margin-bottom:22px; }
    .food-index-app .report-filter-row input,
    .food-index-app .report-filter-row select { height:64px; border-radius:20px; border:1px solid #cfe0f7; background:#fff; padding:0 18px; font-size:16px; font-weight:700; color:#071327; min-width: 230px; }
    .food-index-app .dashboard-table-wrap { border:1px solid #d8e8ff; border-radius:24px; overflow:auto; max-height:360px; background:#fff; }
    .food-index-app .dashboard-table-wrap table { min-width: 1300px; }
    .food-index-app .dashboard-table-wrap th { background:#f4f9ff !important; color:#334155 !important; font-size:12px !important; }
    .food-index-app .dashboard-table-wrap td { font-size:15px !important; }
    .food-index-app .alert-card-v43 h3 { font-size:32px; }
    .food-index-app .alert-stack-v43 { display:grid; gap:16px; margin-top:24px; max-height:520px; overflow:auto; padding-right:4px; }
    .food-index-app .mini-alert { display:grid; grid-template-columns:58px 1fr; gap:16px; align-items:center; padding:18px; border:1px solid #d8e8ff; border-radius:24px; background:rgba(255,255,255,.70); }
    .food-index-app .mini-alert-icon { width:58px; height:58px; border-radius:20px; display:flex; align-items:center; justify-content:center; background:#dbeeff; color:#0f2f75; font-size:22px; }
    .food-index-app .mini-alert b { font-size:20px; display:block; margin-bottom:6px; }
    .food-index-app .mini-alert span { color:#64748b; font-size:16px; line-height:1.35; }
    .food-index-app .task-row-card .table-wrap { max-height: 380px !important; }
    @media (max-width: 1180px) {
      .food-index-app .app-header { grid-template-columns: 1fr !important; }
      .food-index-app .redesign-actions { min-width: 0 !important; width:100%; }
      .food-index-app .dashboard-main-grid, .food-index-app .dashboard-filter-card { grid-template-columns: 1fr; }
      .food-index-app .dashboard-kpi-grid { grid-template-columns: repeat(2, minmax(0,1fr)); }
    }
    @media (max-width: 700px) {
      .food-index-app .redesign-main { padding: 22px !important; }
      .food-index-app .dashboard-kpi-grid { grid-template-columns: 1fr; }
      .food-index-app .redesign-actions { grid-template-columns: 1fr !important; }
    }


    /* v44 compact sidebar + typography polish: closer to the approved reference, Arial, not oversized */
    .food-index-app.real-ui-shell {
      grid-template-columns: 300px minmax(0, 1fr) !important;
    }
    .food-index-app .sidebar {
      width: 300px !important;
      padding: 20px 16px !important;
    }
    .food-index-app .sidebar .brand,
    .food-index-app .sidebar-brand {
      display: flex !important;
      align-items: center !important;
      gap: 14px !important;
      margin: 18px 4px 22px !important;
      padding: 16px !important;
      border-radius: 24px !important;
      text-align: left !important;
    }
    .food-index-app .sidebar .logo {
      flex: 0 0 auto !important;
      width: 52px !important;
      height: 52px !important;
      border-radius: 18px !important;
      font-size: 24px !important;
    }
    .food-index-app .sidebar .brand b {
      display: block !important;
      font-size: 22px !important;
      line-height: 1.05 !important;
      font-weight: 700 !important;
      letter-spacing: -.4px !important;
    }
    .food-index-app .sidebar .brand span {
      display: block !important;
      margin-top: 6px !important;
      font-size: 14px !important;
      line-height: 1.35 !important;
      color: rgba(255,255,255,.74) !important;
    }
    .food-index-app .sidebar nav {
      padding: 0 4px !important;
      gap: 8px !important;
    }
    .food-index-app .sidebar nav button {
      display: flex !important;
      align-items: center !important;
      justify-content: flex-start !important;
      width: 100% !important;
      min-height: 58px !important;
      border-radius: 18px !important;
      padding: 0 14px !important;
      gap: 13px !important;
      font-size: 15.5px !important;
      line-height: 1.2 !important;
      font-weight: 700 !important;
      text-align: left !important;
      white-space: nowrap !important;
    }
    .food-index-app .sidebar nav button svg {
      flex: 0 0 auto !important;
      width: 34px !important;
      height: 34px !important;
      padding: 8px !important;
      border-radius: 14px !important;
    }
    .food-index-app .sidebar-card {
      left: 16px !important;
      right: 16px !important;
      bottom: 18px !important;
      padding: 18px !important;
      border-radius: 22px !important;
    }
    .food-index-app .sidebar-card b { font-size: 18px !important; }
    .food-index-app .sidebar-card p { font-size: 14px !important; line-height: 1.45 !important; }

    .food-index-app .redesign-main { padding: 34px 36px !important; }
    .food-index-app .header-copy h2 {
      font-size: clamp(40px, 4.2vw, 56px) !important;
      letter-spacing: -2.2px !important;
    }
    .food-index-app .header-copy p {
      font-size: 18px !important;
      line-height: 1.45 !important;
      margin-top: 14px !important;
      max-width: 680px !important;
    }
    .food-index-app .redesign-actions { min-width: 390px !important; gap: 12px !important; }
    .food-index-app .user-chip,
    .food-index-app .secondary {
      min-height: 56px !important;
      border-radius: 18px !important;
      font-size: 15.5px !important;
    }
    .food-index-app .dashboard-filter-card { border-radius: 26px !important; padding: 24px !important; }
    .food-index-app .dashboard-filter-card h3 { font-size: 24px !important; }
    .food-index-app .dashboard-filter-card p { font-size: 15px !important; }
    .food-index-app .dashboard-kpi-grid { gap: 18px !important; }
    .food-index-app .dash-kpi {
      min-height: 158px !important;
      padding: 22px !important;
      border-radius: 24px !important;
    }
    .food-index-app .dash-kpi small { font-size: 16px !important; }
    .food-index-app .dash-kpi strong { font-size: 38px !important; margin-top: 14px !important; }
    .food-index-app .dash-kpi span { font-size: 14px !important; margin-top: 18px !important; }
    .food-index-app .dashboard-card { border-radius: 26px !important; padding: 24px !important; }
    .food-index-app .dashboard-card h3 { font-size: 28px !important; }
    .food-index-app .dashboard-card p { font-size: 16px !important; }
    .food-index-app .report-filter-row input,
    .food-index-app .report-filter-row select {
      height: 56px !important;
      border-radius: 18px !important;
      font-size: 15px !important;
    }
    .food-index-app .dashboard-filter-actions select { height: 52px !important; border-radius: 16px !important; }
    .food-index-app .dashboard-filter-actions button,
    .food-index-app .dashboard-card-head button,
    .food-index-app .task-row-card button {
      min-height: 52px !important;
      border-radius: 16px !important;
      font-size: 15px !important;
    }
    .food-index-app .mini-alert { grid-template-columns: 50px 1fr !important; border-radius: 20px !important; padding: 16px !important; }
    .food-index-app .mini-alert-icon { width: 50px !important; height: 50px !important; border-radius: 17px !important; }
    .food-index-app .mini-alert b { font-size: 17px !important; }
    .food-index-app .mini-alert span { font-size: 14px !important; }
    @media (max-width: 1180px) {
      .food-index-app.real-ui-shell { grid-template-columns: 1fr !important; }
      .food-index-app .sidebar { width: 100% !important; }
    }


    /* v45 compact elegant page header: all Food Index menus, Arial, less bulky */
    .food-index-app {
      font-family: Arial, Helvetica, sans-serif !important;
    }
    .food-index-app .app-header {
      margin-bottom: 22px !important;
      align-items: center !important;
      gap: 18px !important;
      grid-template-columns: minmax(0,1fr) auto !important;
    }
    .food-index-app .header-copy {
      position: relative !important;
      padding: 18px 20px 18px 22px !important;
      border-radius: 26px !important;
      background: linear-gradient(135deg, rgba(255,255,255,.46), rgba(239,247,255,.22)) !important;
      border: 1px solid rgba(216,232,255,.58) !important;
      box-shadow: 0 16px 40px rgba(37,99,235,.055) !important;
      overflow: hidden !important;
    }
    .food-index-app .header-copy::before {
      content: "";
      position: absolute;
      left: 0;
      top: 18px;
      bottom: 18px;
      width: 5px;
      border-radius: 999px;
      background: linear-gradient(180deg, #2563eb, #0ea5e9);
    }
    .food-index-app .header-copy h2 {
      font-family: Arial, Helvetica, sans-serif !important;
      font-size: clamp(30px, 3.2vw, 42px) !important;
      line-height: 1.08 !important;
      letter-spacing: -1.35px !important;
      font-weight: 800 !important;
      color: #071327 !important;
      margin: 0 !important;
    }
    .food-index-app .header-copy p {
      font-family: Arial, Helvetica, sans-serif !important;
      font-size: 15.5px !important;
      line-height: 1.45 !important;
      margin: 8px 0 0 !important;
      max-width: 660px !important;
      color: #64748b !important;
      font-weight: 500 !important;
    }
    .food-index-app .redesign-actions {
      min-width: 360px !important;
      max-width: 420px !important;
      gap: 10px !important;
      align-items: center !important;
    }
    .food-index-app .user-chip {
      min-height: 50px !important;
      border-radius: 17px !important;
      padding: 0 16px !important;
      font-size: 14.5px !important;
      font-weight: 700 !important;
      box-shadow: 0 14px 34px rgba(37,99,235,.075) !important;
      border-color: rgba(207,224,247,.95) !important;
    }
    .food-index-app .secondary {
      min-height: 50px !important;
      border-radius: 17px !important;
      padding: 0 16px !important;
      font-size: 14.5px !important;
      font-weight: 700 !important;
      box-shadow: 0 14px 34px rgba(37,99,235,.075) !important;
      border-color: rgba(207,224,247,.95) !important;
    }
    .food-index-app .dashboard-filter-card {
      margin-top: 2px !important;
    }
    .food-index-app .dashboard-card-head h3,
    .food-index-app .panel-head h3,
    .food-index-app .dashboard-card h3 {
      font-family: Arial, Helvetica, sans-serif !important;
      letter-spacing: -.7px !important;
      font-weight: 800 !important;
    }
    @media (max-width: 1180px) {
      .food-index-app .app-header {
        grid-template-columns: 1fr !important;
      }
      .food-index-app .redesign-actions {
        min-width: 0 !important;
        max-width: none !important;
        width: 100% !important;
      }
    }
    @media (max-width: 700px) {
      .food-index-app .header-copy {
        padding: 16px 16px 16px 20px !important;
        border-radius: 22px !important;
      }
      .food-index-app .header-copy h2 {
        font-size: 30px !important;
        letter-spacing: -1px !important;
      }
      .food-index-app .header-copy p {
        font-size: 14px !important;
      }
    }

    /* v47 targeted fixes */
    .food-index-app .standard-meta { display:flex; flex-wrap:wrap; gap:8px; margin-top:8px; }
    .food-index-app .standard-meta span { display:inline-flex; align-items:center; min-height:28px; padding:6px 10px; border-radius:999px; background:#eef6ff; color:#1e40af; font-size:12px; font-weight:700; }
    .food-index-app .standard-box { margin-top:10px; padding:10px 12px; border-radius:14px; border:1px solid #dbeafe; background:#f8fbff; color:#334155; font-size:13px; line-height:1.45; }
    .food-index-app .camera-options { display:grid; grid-template-columns:1fr 1fr; gap:8px; width:100%; margin-top:8px; }
    .food-index-app .camera-option { min-height:44px !important; border-radius:14px !important; padding:0 10px !important; box-shadow:none !important; background:#ffffff !important; color:#1d4ed8 !important; border:1px solid #bfdbfe !important; font-size:12px !important; display:flex !important; align-items:center !important; justify-content:center !important; gap:6px !important; }
    .food-index-app .outstanding-table-wrap table { min-width: 1820px; table-layout: fixed; }
    .food-index-app .outstanding-table-wrap th,
    .food-index-app .outstanding-table-wrap td,
    .food-index-app .outstanding-export-table .table-wrap th,
    .food-index-app .outstanding-export-table .table-wrap td { white-space: normal; overflow-wrap: anywhere; word-break: normal; line-height: 1.22; }
    .food-index-app .outstanding-table-wrap th,
    .food-index-app .outstanding-export-table .table-wrap th { font-size: 11px; letter-spacing: .07em; vertical-align: middle; }
    .food-index-app .outstanding-table-wrap thead th,
    .food-index-app .outstanding-export-table .table-wrap thead th { position: sticky; top: 0; z-index: 12; background: linear-gradient(180deg, #f8fbff 0%, #eef6ff 100%) !important; box-shadow: 0 1px 0 #dbe7f7, 0 8px 18px rgba(15,23,42,.08); }
    .food-index-app .outstanding-table-wrap th:nth-child(1), .food-index-app .outstanding-table-wrap td:nth-child(1) { width:70px; }
    .food-index-app .outstanding-table-wrap th:nth-child(2), .food-index-app .outstanding-table-wrap td:nth-child(2) { width:155px; }
    .food-index-app .outstanding-table-wrap th:nth-child(3), .food-index-app .outstanding-table-wrap td:nth-child(3) { width:110px; }
    .food-index-app .outstanding-table-wrap th:nth-child(4), .food-index-app .outstanding-table-wrap td:nth-child(4) { width:220px; }
    .food-index-app .outstanding-table-wrap th:nth-child(5), .food-index-app .outstanding-table-wrap td:nth-child(5) { width:240px; }
    .food-index-app .outstanding-table-wrap th:nth-child(6), .food-index-app .outstanding-table-wrap td:nth-child(6) { width:100px; }
    .food-index-app .outstanding-table-wrap th:nth-child(7), .food-index-app .outstanding-table-wrap td:nth-child(7) { width:235px; }
    .food-index-app .outstanding-table-wrap th:nth-child(8), .food-index-app .outstanding-table-wrap td:nth-child(8) { width:250px; }
    .food-index-app .outstanding-table-wrap th:nth-child(9), .food-index-app .outstanding-table-wrap td:nth-child(9) { width:100px; }
    .food-index-app .outstanding-table-wrap th:nth-child(10), .food-index-app .outstanding-table-wrap td:nth-child(10) { width:170px; }
    .food-index-app .outstanding-table-wrap th:nth-child(11), .food-index-app .outstanding-table-wrap td:nth-child(11) { width:170px; }
    .food-index-app .status-task-cell { min-width:170px; max-width:170px; text-align:center; vertical-align:middle !important; }
    .food-index-app .status-task-cell .task-status-pill { min-width:136px; white-space:normal; text-align:center; line-height:1.15; padding:8px 10px; }
    .food-index-app .outstanding-export-table .table-wrap table { min-width: 1840px; table-layout: fixed; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(1), .food-index-app .outstanding-export-table .table-wrap td:nth-child(1) { width:80px; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(2), .food-index-app .outstanding-export-table .table-wrap td:nth-child(2) { width:170px; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(3), .food-index-app .outstanding-export-table .table-wrap td:nth-child(3) { width:110px; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(4), .food-index-app .outstanding-export-table .table-wrap td:nth-child(4) { width:230px; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(5), .food-index-app .outstanding-export-table .table-wrap td:nth-child(5) { width:245px; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(6), .food-index-app .outstanding-export-table .table-wrap td:nth-child(6) { width:245px; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(7), .food-index-app .outstanding-export-table .table-wrap td:nth-child(7) { width:255px; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(8), .food-index-app .outstanding-export-table .table-wrap td:nth-child(8) { width:110px; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(9), .food-index-app .outstanding-export-table .table-wrap td:nth-child(9) { width:180px; text-align:center; vertical-align:middle; }
    .food-index-app .outstanding-export-table .table-wrap th:nth-child(10), .food-index-app .outstanding-export-table .table-wrap td:nth-child(10) { width:185px; text-align:center; vertical-align:middle; }
    .food-index-app .outstanding-export-table .status-pill { white-space:normal; text-align:center; min-width:138px; line-height:1.15; padding:8px 10px; }
    .food-index-app .inspection-item.finding.approval-review-finding { background:linear-gradient(180deg,#fee2e2,#fff5f5) !important; border-color:#ef4444 !important; box-shadow:0 12px 34px rgba(239,68,68,.16) !important; }
    .food-index-app .approval-review-finding .param-index { background:#ef4444 !important; color:white !important; }
    /* v48 dashboard layout polish */
    .food-index-app .dashboard-filter-card { display:grid !important; grid-template-columns:1fr !important; gap:18px !important; align-items:stretch !important; padding:22px !important; }
    .food-index-app .dashboard-filter-card > div:first-child { max-width: none !important; }
    .food-index-app .dashboard-filter-card small { font-size:13px !important; margin-bottom:6px !important; }
    .food-index-app .dashboard-filter-card h3 { font-size:24px !important; line-height:1.15 !important; }
    .food-index-app .dashboard-filter-card p { font-size:15px !important; line-height:1.45 !important; max-width: 760px !important; }
    .food-index-app .dashboard-admin-filters { display:grid; grid-template-columns:repeat(4,minmax(155px,1fr)) minmax(130px,.6fr); gap:12px; align-items:end; margin-top:0; width:100%; }
    .food-index-app .dashboard-admin-filters label { display:grid; gap:7px; font-weight:700; color:#334155; font-size:13px; min-width:0; }
    .food-index-app .dashboard-admin-filters select { height:46px; width:100%; border-radius:15px; border:1px solid #bfdbfe; background:white; padding:0 12px; font-weight:700; min-width:0; }
    .food-index-app .dashboard-admin-filters button { min-height:46px !important; height:46px !important; border-radius:15px !important; padding:0 16px !important; white-space:nowrap; }
    .food-index-app .dashboard-tags { margin-top:0 !important; }
    .food-index-app .dashboard-analysis-grid { display:grid; grid-template-columns:1fr; gap:18px; margin-top:20px; }
    .food-index-app .analysis-table table { min-width:1200px; }
    .food-index-app .resume-table table { min-width:980px; }
    .food-index-app .resume-table .capture-all-wrap {
      max-height: none !important;
      overflow: visible !important;
    }
    .food-index-app .resume-table .capture-all-wrap table {
      width: 100% !important;
      min-width: 0 !important;
      table-layout: fixed !important;
    }
    .food-index-app .resume-table .capture-all-wrap th,
    .food-index-app .resume-table .capture-all-wrap td {
      padding: 8px 10px !important;
      font-size: 13px !important;
      line-height: 1.22 !important;
      white-space: normal !important;
      word-break: break-word !important;
    }
    .food-index-app .resume-table .capture-all-wrap th:nth-child(1),
    .food-index-app .resume-table .capture-all-wrap td:nth-child(1) { width: 4.5%; }
    .food-index-app .resume-table .capture-all-wrap th:nth-child(2),
    .food-index-app .resume-table .capture-all-wrap td:nth-child(2) { width: 8%; }
    .food-index-app .resume-table .capture-all-wrap th:nth-child(3),
    .food-index-app .resume-table .capture-all-wrap td:nth-child(3) { width: 13%; }
    .food-index-app .resume-table .capture-all-wrap th:nth-child(4),
    .food-index-app .resume-table .capture-all-wrap td:nth-child(4) { width: 13%; }
    .food-index-app .resume-table .capture-all-wrap th:nth-child(5),
    .food-index-app .resume-table .capture-all-wrap td:nth-child(5) { width: 16%; }
    .food-index-app .resume-table .capture-all-wrap th:nth-child(6),
    .food-index-app .resume-table .capture-all-wrap td:nth-child(6) { width: 16%; }
    .food-index-app .resume-table .capture-all-wrap th:nth-child(7),
    .food-index-app .resume-table .capture-all-wrap td:nth-child(7) { width: 15%; }
    .food-index-app .resume-table .capture-all-wrap th:nth-child(8),
    .food-index-app .resume-table .capture-all-wrap td:nth-child(8) { width: 14.5%; }
    .food-index-app .dashboard-card.compact-card { padding:22px !important; }
    .food-index-app .dashboard-card.compact-card h3 { font-size:26px !important; }
    .food-index-app .score-badge { display:inline-flex; padding:7px 11px; border-radius:999px; font-weight:800; font-size:12px; }
    .food-index-app .score-badge.ok { background:#dcfce7; color:#166534; }
    .food-index-app .score-badge.bad { background:#fee2e2; color:#991b1b; }
    @media (max-width: 760px) {
      .food-index-app .food-modal-backdrop { align-items:flex-start !important; overflow:auto !important; padding:10px !important; }
      .food-index-app .food-inspection-modal { width:100% !important; max-height:none !important; min-height:calc(100dvh - 20px); overflow:visible !important; }
      .food-index-app .inspection-entry-modal { height:calc(100dvh - 20px) !important; max-height:calc(100dvh - 20px) !important; min-height:0 !important; overflow:hidden !important; }
      .food-index-app .inspection-modal-body { overflow-y:auto !important; min-height:0 !important; }
      .food-index-app .inspection-scroll { max-height:none !important; overflow:visible !important; }
      .food-index-app .sticky-actions { position:sticky; bottom:0; }
      .food-index-app .dashboard-admin-filters { grid-template-columns:1fr; }
      .food-index-app .camera-options { grid-template-columns:1fr; }
    }



    .food-index-app .panel-actions { display:flex; gap:10px; flex-wrap:wrap; justify-content:flex-end; }
    .food-index-app .action-cell { display:flex; gap:8px; align-items:center; flex-wrap:wrap; white-space:nowrap; }
    .food-index-app button.small { min-height:34px; padding:7px 11px; font-size:12px; border-radius:10px; }


    /* v53 PC/mobile single-scroll inspection modal: no separated checklist/signature scroll */
    .food-index-app .food-modal-backdrop {
      align-items: flex-start !important;
      justify-content: center !important;
      overflow-y: auto !important;
      overflow-x: hidden !important;
      padding: 22px !important;
      overscroll-behavior: contain !important;
    }
    .food-index-app .food-inspection-modal.inspection-entry-modal {
      width: min(1060px, calc(100vw - 34px)) !important;
      height: auto !important;
      min-height: 0 !important;
      max-height: none !important;
      display: block !important;
      overflow: visible !important;
      margin: 0 auto 22px !important;
    }
    .food-index-app .inspection-entry-modal .food-modal-head,
    .food-index-app .inspection-entry-modal .sticky-actions {
      flex: none !important;
    }
    .food-index-app .inspection-modal-body {
      display: block !important;
      flex: none !important;
      min-height: 0 !important;
      max-height: none !important;
      overflow: visible !important;
      overflow-y: visible !important;
      overflow-x: visible !important;
      padding: 0 !important;
    }
    .food-index-app .inspection-modal-body::-webkit-scrollbar {
      width: 0 !important;
      height: 0 !important;
    }
    .food-index-app .inspection-scroll,
    .food-index-app .signature-section {
      max-height: none !important;
      overflow: visible !important;
    }
    .food-index-app .sticky-actions {
      position: static !important;
      bottom: auto !important;
      background: #fff !important;
      border-top: 1px solid #e8eef8 !important;
      padding: 16px 30px 22px !important;
      margin: 0 !important;
    }
    @media (max-width: 760px) {
      .food-index-app .food-modal-backdrop {
        padding: 10px !important;
      }
      .food-index-app .food-inspection-modal.inspection-entry-modal {
        width: 100% !important;
        min-height: calc(100dvh - 20px) !important;
        height: auto !important;
        max-height: none !important;
        overflow: visible !important;
      }
      .food-index-app .inspection-modal-body {
        overflow: visible !important;
        max-height: none !important;
      }
      .food-index-app .sticky-actions {
        position: static !important;
        padding-left: 18px !important;
        padding-right: 18px !important;
      }
    }


    /* v54: PC inspection checklist + signature share one large continuous scroll area */
    @media (min-width: 761px) {
      .food-index-app .food-modal-backdrop {
        align-items: center !important;
        padding: 18px !important;
        overflow: hidden !important;
      }
      .food-index-app .food-inspection-modal.inspection-entry-modal {
        width: min(1120px, calc(100vw - 36px)) !important;
        height: calc(100dvh - 36px) !important;
        max-height: calc(100dvh - 36px) !important;
        display: flex !important;
        flex-direction: column !important;
        overflow: hidden !important;
        margin: 0 auto !important;
      }
      .food-index-app .inspection-entry-modal .food-modal-head,
      .food-index-app .inspection-entry-modal .sticky-actions {
        flex: 0 0 auto !important;
      }
      .food-index-app .inspection-modal-body {
        flex: 1 1 auto !important;
        min-height: 0 !important;
        display: flex !important;
        flex-direction: column !important;
        overflow: hidden !important;
        padding: 0 !important;
      }
      .food-index-app .inspection-summary,
      .food-index-app .inspection-note {
        flex: 0 0 auto !important;
      }
      .food-index-app .inspection-scroll {
        flex: 1 1 auto !important;
        min-height: 430px !important;
        max-height: none !important;
        overflow-y: auto !important;
        overflow-x: hidden !important;
        padding: 16px 30px 170px !important;
        display: grid !important;
        gap: 14px !important;
        align-content: start !important;
        scroll-padding-bottom: 180px !important;
      }
      .food-index-app .inspection-scroll::-webkit-scrollbar { width: 10px; }
      .food-index-app .inspection-scroll::-webkit-scrollbar-thumb { background:#94a3b8; border-radius:999px; }
      .food-index-app .signature-section {
        margin: 4px 0 80px !important;
        padding: 18px 18px 26px !important;
        max-height: none !important;
        overflow: visible !important;
      }
      .food-index-app .sticky-actions {
        position: static !important;
        bottom: auto !important;
        flex: 0 0 auto !important;
        background: #fff !important;
        border-top: 1px solid #e8eef8 !important;
        padding: 16px 30px 22px !important;
        margin: 0 !important;
      }
    }


    /* v55: keep signature form/add button visible above footer actions on PC/tablet/mobile */
    .food-index-app .inspection-entry-modal .signature-add-row {
      margin: 18px 0 16px !important;
      padding-bottom: 6px !important;
    }
    .food-index-app .inspection-entry-modal .signature-add-row button {
      min-height: 46px !important;
      padding: 12px 18px !important;
      border-radius: 14px !important;
    }
    .food-index-app .inspection-entry-modal .signature-pad-wrap {
      margin-bottom: 10px !important;
    }
    @media (max-width: 760px) {
      .food-index-app .inspection-scroll {
        padding-bottom: 180px !important;
        scroll-padding-bottom: 190px !important;
      }
      .food-index-app .signature-section {
        margin-bottom: 120px !important;
      }
      .food-index-app .signature-canvas {
        height: 210px !important;
      }
    }


    /* v56 Food Index Admin Panel & role participation dashboard */
    .food-index-app .food-admin-panel .form-grid { align-items:end; }
    .food-index-app .upload-actions { display:flex; align-items:center; gap:14px; flex-wrap:wrap; margin-bottom:14px; }
    .food-index-app .upload-tile { display:inline-flex; align-items:center; gap:10px; border:1px dashed #93c5fd; background:#eff6ff; color:#1d4ed8; padding:16px 18px; border-radius:18px; cursor:pointer; font-weight:800; }
    .food-index-app .role-participation-card { margin-top: 0; }
    .food-index-app .signature-live-search { grid-template-columns: 1fr 70px !important; }
    .food-index-app .signature-live-search > button[disabled] { opacity:.55; cursor:not-allowed; }
    .food-index-app .readonly-signature-grid input[readonly] { background:#f8fbff; color:#334155; }
    .food-index-app .driver-results { grid-column: 1 / -1; display:grid; gap:8px; max-height:220px; overflow:auto; padding:8px; border:1px solid #dbe7f7; border-radius:16px; background:#fff; }
    .food-index-app .driver-results button { text-align:left; border:1px solid #dbe7f7; background:#f8fbff !important; color:#0f172a !important; box-shadow:none !important; border-radius:14px; padding:10px 12px; }
    .food-index-app .driver-results button b { display:block; }
    .food-index-app .driver-results button span { display:block; color:#64748b; font-size:12px; margin-top:3px; }


    /* Food Index report photo preview modal + compact photo button */
    .food-index-app .photo-view-btn {
      display:inline-flex;
      align-items:center;
      justify-content:center;
      gap:6px;
      min-height:32px;
      padding:7px 11px;
      border-radius:999px;
      border:1px solid #bfdbfe;
      background:#eff6ff !important;
      color:#1d4ed8 !important;
      box-shadow:none !important;
      font-size:12px;
      font-weight:800;
      white-space:nowrap;
    }
    .food-index-app .photo-view-btn:hover { background:#dbeafe !important; }
    .food-index-app .inline-photo-view-btn { width:max-content; align-self:flex-start; margin-top:-4px; }
    .food-index-app .photo-view-backdrop {
      position:fixed !important;
      inset:0 !important;
      z-index:10020 !important;
      display:flex !important;
      align-items:center !important;
      justify-content:center !important;
      padding:18px !important;
      background:rgba(15,23,42,.62) !important;
      backdrop-filter:blur(8px) !important;
      overflow:auto !important;
    }
    .food-index-app .photo-view-modal {
      width:min(860px, calc(100vw - 28px));
      max-height:calc(100dvh - 32px);
      display:flex;
      flex-direction:column;
      overflow:hidden;
      border-radius:22px;
      background:#fff;
    }
    .food-index-app .photo-view-head {
      flex:0 0 auto;
      display:flex;
      align-items:flex-start;
      justify-content:space-between;
      gap:14px;
      padding:18px 20px;
      border-bottom:1px solid #e8eef8;
    }
    .food-index-app .photo-view-head h3 { margin:0 0 5px; font-size:22px; line-height:1.12; color:#0f172a; }
    .food-index-app .photo-view-head p { margin:0; color:#64748b; font-size:13px; line-height:1.35; font-weight:600; }
    .food-index-app .photo-view-body {
      flex:1 1 auto;
      min-height:0;
      padding:14px;
      background:#0f172a;
      display:flex;
      align-items:center;
      justify-content:center;
      overflow:auto;
    }
    .food-index-app .photo-view-body img {
      display:block;
      max-width:100%;
      max-height:calc(100dvh - 190px);
      object-fit:contain;
      border-radius:14px;
      background:#fff;
      box-shadow:0 18px 46px rgba(0,0,0,.24);
    }
    .food-index-app .photo-view-failed {
      display:grid;
      justify-items:center;
      gap:7px;
      color:#fff;
      text-align:center;
      padding:28px;
    }
    .food-index-app .photo-view-failed span { color:#cbd5e1; }
    .food-index-app .photo-view-actions {
      flex:0 0 auto;
      display:flex;
      justify-content:flex-end;
      gap:10px;
      padding:14px 18px 18px;
      border-top:1px solid #e8eef8;
      background:#fff;
    }
    @media (max-width:760px) {
      .food-index-app .photo-view-backdrop { padding:10px !important; }
      .food-index-app .photo-view-modal { width:100%; max-height:calc(100dvh - 20px); }
      .food-index-app .photo-view-head { padding:14px 15px; }
      .food-index-app .photo-view-head h3 { font-size:19px; }
      .food-index-app .photo-view-body { padding:10px; }
      .food-index-app .photo-view-body img { max-height:calc(100dvh - 170px); }
    }

  `}</style>
}

function Table({ rows, columns, empty='Belum ada data.', captureAll=false }){
  const [q, setQ] = useState('')
  const cols = columns || (rows[0] ? Object.keys(rows[0]) : [])
  const filtered = rows.filter(r => !q || JSON.stringify(r).toLowerCase().includes(q.toLowerCase()))
  const wrapStyle = captureAll ? { maxHeight: 'none', overflow: 'visible' } : { maxHeight: 420, overflow: 'auto' }
  return <div className={captureAll ? 'capture-table-root' : ''}>
    <div className="table-toolbar"><div className="searchbox"><Search size={18}/><input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search row data..." /></div></div>
    <div className={captureAll ? 'table-wrap capture-all-wrap' : 'table-wrap'} style={wrapStyle}><table><thead><tr>{cols.map(c => <th key={c}>{c}</th>)}</tr></thead><tbody>{filtered.map((r,i)=><tr key={r.id || i}>{cols.map(c => <td key={c} className={tableCellClass(c, r[c])}>{renderCell(r[c], c)}</td>)}</tr>)}</tbody></table>{!filtered.length && <p className="muted table-empty">{empty}</p>}</div>
  </div>
}
function tableCellClass(column, value){
  const key = String(column || '').toLowerCase()
  const val = String(value || '').trim().toLowerCase()
  if (key !== 'remark') return ''
  if (val === 'tercapai') return 'remark-cell remark-cell-success'
  if (val === 'tidak tercapai' || val === 'belum tercapai') return 'remark-cell remark-cell-danger'
  if (val === 'not yet approved') return 'remark-cell remark-cell-warning'
  if (val === 'are not done') return 'remark-cell remark-cell-danger'
  return 'remark-cell remark-cell-neutral'
}
function renderCell(v, column){
  if (v === null || v === undefined || v === '') return '-'
  if (String(column || '').toLowerCase() === 'remark') return String(v)
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

function normalizeRole(value){
  const v = cleanText(value)
  const found = FOOD_ACCESS_ROLES.find(r => r.toLowerCase() === v.toLowerCase())
  if (found) return found
  if (['Platform Admin', 'App Admin'].some(r => r.toLowerCase() === v.toLowerCase())) return 'Administrator'
  if (v.toLowerCase() === 'atasan site') return 'Section Head'
  return v
}
function roleTarget(role){ return FOOD_ROLE_TARGETS[normalizeRole(role)] || 0 }
function foodUiRoleToDbRole(role){
  const r = normalizeRole(role)
  if (r === 'Administrator') return 'App Admin'
  if (r === 'GL') return 'GL'
  if (r === 'Section Head' || r === 'Dept Head') return 'Atasan Site'
  return r
}
function foodRoleFromAccess(row){
  return normalizeRole(row?.food_role_level || row?.role || '')
}
function buildFoodAccessPayload(row){
  const foodRole = normalizeRole(row.role)
  return {
    ...row,
    role: foodUiRoleToDbRole(foodRole),
    food_role_level: foodRole,
    site_id: foodRole === 'Administrator' ? null : (cleanText(row.site_id) || null)
  }
}
async function getFoodIndexApplication(){
  const { data, error } = await supabase
    .from('applications')
    .select('id, app_code, app_name, status')
    .or('app_code.ilike.%food%,app_name.ilike.%food%')
    .limit(5)
  if (error) throw error
  const app = (data || []).find(a => String(a.app_name || '').toLowerCase().includes('food')) || (data || [])[0]
  if (!app) throw new Error('Aplikasi Food Index belum ditemukan di tabel applications.')
  return app
}
function ymFromDate(dateText){ return String(dateText || '').slice(0,7) }
function sameMonth(dateText, year, month){
  if (!dateText) return false
  const [y,m] = String(dateText).split('-')
  return (year === 'ALL' || y === String(year)) && (month === 'ALL' || m === String(month).padStart(2,'0'))
}
function sameDashboardPeriodByDate(dateText, year, month, week='ALL'){
  const date = String(dateText || '').slice(0,10)
  if (!date) return false
  const weekStart = startOfWeekMonday(new Date(date))
  if (week !== 'ALL') return weekStart === week
  const [y,m] = date.split('-')
  return (year === 'ALL' || y === String(year)) && (month === 'ALL' || m === String(month).padStart(2,'0'))
}
function outstandingCreatedWeekStart(row){
  // Dashboard leadtime mengikuti minggu task inspeksi yang melahirkan outstanding,
  // bukan due date, close date, atau timestamp upload/action plan.
  const taskWeek = row?.food_findings?.food_weekly_tasks?.week_start_date || ''
  if (taskWeek) return taskWeek
  const createdDate = String(row?.created_at || '').slice(0,10)
  return createdDate ? startOfWeekMonday(new Date(createdDate)) : ''
}
function outstandingMatchesDashboardPeriod(row, year, month, week='ALL'){
  const outstandingWeekStart = outstandingCreatedWeekStart(row)
  if (!outstandingWeekStart) return false
  if (week !== 'ALL') return outstandingWeekStart === week
  return sameMonth(outstandingWeekStart, year, month)
}
function pctValue(n, d){ return d ? (Number(n || 0) / Number(d || 0)) * 100 : 0 }
function fmtPct(value){
  const n = Number(value || 0)
  return `${Number.isFinite(n) ? n.toFixed(2) : '0.00'}%`
}
function taskIsApproved(task){ return String(task?.status || '') === 'Approved' }
function taskHasInspectionData(task){
  const answers = task?.food_inspection_answers || []
  return Boolean(task?.submitted_at || task?.approved_at || answers.length || ['Approved','Expired','Waiting Approval','Need Action Plan'].includes(String(task?.status || '')))
}
function achievementRemark(score, task){
  if (task && !taskIsApproved(task)) return 'Not Yet Approved'
  return Number(score || 0) >= 95 ? 'Tercapai' : 'Tidak tercapai'
}
function taskWeekIsExpired(task){
  const end = String(task?.week_end_date || '').slice(0,10)
  return Boolean(end && end < today())
}
function taskHasActualInspectionWork(task){
  const answers = task?.food_inspection_answers || []
  return Boolean(task?.submitted_at || task?.approved_at || answers.length)
}
function scoreInspectionDisplay(task, score){
  const hasWork = taskHasActualInspectionWork(task)
  const expiredByStatusOrDate = String(task?.status || '') === 'Expired' || taskWeekIsExpired(task)

  // Khusus tabel Score Food Index per Inspeksi:
  // - sudah inspeksi tetapi belum approved: tampil sebagai Open + Not Yet Approved
  // - belum dilakukan sampai lewat minggu: tampil sebagai Expired + Are Not Done
  if (hasWork && !taskIsApproved(task)) return { status: 'Open', remark: 'Not Yet Approved' }
  if (!hasWork && expiredByStatusOrDate) return { status: 'Expired', remark: 'Are Not Done' }
  if (!hasWork) return { status: String(task?.status || 'Open') || 'Open', remark: 'Are Not Done' }
  return {
    status: String(task?.status || '-') || '-',
    remark: Number(score || 0) >= 95 ? 'Tercapai' : 'Tidak tercapai'
  }
}


function FoodAdminPanel({ context, profile }){
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [msg, setMsg] = useState('')
  const [error, setError] = useState('')
  const [sites, setSites] = useState([])
  const [app, setApp] = useState(null)
  const [accessRows, setAccessRows] = useState([])
  const [manual, setManual] = useState({ nama:'', nrp:'', email:'', password:'', role:'GL', site_id:'' })
  const [uploadRows, setUploadRows] = useState([])
  const [resetForm, setResetForm] = useState({ access_id:'', password:'' })

  useEffect(() => { load() }, [context?.id])
  useEffect(() => { if (!msg) return; const t = setTimeout(()=>setMsg(''), 7000); return () => clearTimeout(t) }, [msg])

  async function load(){
    setLoading(true); setError('')
    try {
      const [siteData, foodApp] = await Promise.all([fetchSites(), getFoodIndexApplication()])
      setSites(siteData || [])
      setApp(foodApp)
      setManual(prev => ({ ...prev, site_id: prev.site_id || context?.site_id || '' }))
      let q = supabase.from('user_app_access')
        .select('id,user_id,app_id,role,food_role_level,site_id,status,created_at,users_profile(id,nama,email,nrp,status),sites(id,site_code,site_name)')
        .eq('app_id', foodApp.id)
        .order('created_at', { ascending:false })
      if (!adminCanSeeAll(context) && context?.site_id) q = q.eq('site_id', context.site_id)
      const { data, error } = await q
      if (error) throw error
      setAccessRows(data || [])
    } catch(e){ setError(e.message) }
    setLoading(false)
  }
  function validateUserPayload(row){
    const foodRole = normalizeRole(row.role)
    if (!FOOD_ACCESS_ROLES.includes(foodRole)) throw new Error(`Role tidak valid: ${row.role}. Gunakan Administrator, GL, Section Head, atau Dept Head.`)
    if (foodRole !== 'Administrator' && !cleanText(row.site_id)) throw new Error('Site wajib dipilih untuk role GL, Section Head, dan Dept Head.')
    if (!cleanText(row.email)) throw new Error('Email wajib diisi.')
    if (!cleanText(row.password) || cleanText(row.password).length < 6) throw new Error('Password minimal 6 karakter.')
    return buildFoodAccessPayload({ ...row, role: foodRole })
  }
  async function createOneUser(e){
    e?.preventDefault?.()
    setSaving(true); setMsg(''); setError('')
    try {
      if (!app?.id) throw new Error('Aplikasi Food Index belum siap.')
      const payload = validateUserPayload({ ...manual, app_id: app.id })
      const { data, error } = await supabase.functions.invoke('admin-create-user', { body: payload })
      if (error) throw error
      if (data?.error) throw new Error(data.error)
      if (!data?.auth_user_id) throw new Error('User belum terbuat di Supabase Auth. Deploy ulang Edge Function admin-create-user versi terbaru lalu coba lagi.')
      setMsg(`User ${manual.email} berhasil dibuat dan dimapping ke Food Index. Auth ID: ${data.auth_user_id}`)
      setManual({ nama:'', nrp:'', email:'', password:'', role:'GL', site_id:context?.site_id || '' })
      await load()
    } catch(e){ setError(e.message) }
    setSaving(false)
  }
  async function handleUploadExcel(file){
    setError(''); setMsg('')
    if (!file) return
    const buf = await file.arrayBuffer()
    const wb = XLSX.read(buf)
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]] || {})
    const parsed = rows.map((raw, idx) => {
      const r = normalizeRow(raw)
      const siteCode = cleanText(getVal(r, ['site_code','SITE_CODE','site','Site']))
      const site = sites.find(s => String(s.site_code).toLowerCase() === siteCode.toLowerCase())
      const role = normalizeRole(getVal(r, ['role','Role','ROLE'])) || 'GL'
      const email = cleanText(getVal(r, ['email','Email','EMAIL']))
      const password = cleanText(getVal(r, ['password','Password','PASSWORD']))
      const nama = cleanText(getVal(r, ['nama','Nama','NAMA','name','Name']))
      const nrp = cleanText(getVal(r, ['nrp','NRP','nrp_driver','NRP_DRIVER']))
      const issues = []
      if (!FOOD_ACCESS_ROLES.includes(role)) issues.push('role tidak valid')
      if (role !== 'Administrator' && !site) issues.push('site_code tidak ditemukan')
      if (!email) issues.push('email kosong')
      if (!password || password.length < 6) issues.push('password minimal 6 karakter')
      return { row: idx + 2, nama, nrp, email, password, role, site_code:siteCode, site_id: role === 'Administrator' ? '' : (site?.id || ''), status: issues.length ? `Error: ${issues.join(', ')}` : 'Valid' }
    })
    setUploadRows(parsed)
    const invalid = parsed.filter(r => r.status !== 'Valid').length
    setMsg(invalid ? `Preview selesai. Ada ${invalid} baris perlu diperbaiki.` : `Preview selesai. ${parsed.length} baris valid dan siap diimport.`)
  }
  async function importExcelUsers(){
    setSaving(true); setError(''); setMsg('')
    try {
      if (!app?.id) throw new Error('Aplikasi Food Index belum siap.')
      const valid = uploadRows.filter(r => r.status === 'Valid')
      if (!valid.length) throw new Error('Tidak ada baris valid untuk diimport.')
      const notes = []
      for (const r of valid) {
        const payload = validateUserPayload({ ...r, app_id: app.id })
        const { data, error } = await supabase.functions.invoke('admin-create-user', { body: payload })
        if (error || data?.error || !data?.auth_user_id) notes.push(`Baris ${r.row}: gagal (${error?.message || data?.error || 'auth_user_id kosong'})`)
        else notes.push(`Baris ${r.row}: sukses ${r.email}`)
      }
      setMsg(`Import selesai. ${notes.join(' | ')}`)
      setUploadRows([])
      await load()
    } catch(e){ setError(e.message) }
    setSaving(false)
  }
  async function resetUserPassword(e){
    e?.preventDefault?.()
    setSaving(true); setError(''); setMsg('')
    try {
      const access = accessRows.find(r => String(r.id) === String(resetForm.access_id))
      if (!access) throw new Error('Pilih akun yang akan diganti passwordnya.')
      if (!cleanText(resetForm.password) || cleanText(resetForm.password).length < 6) throw new Error('Password baru minimal 6 karakter.')
      const email = access.users_profile?.email || ''
      if (!email) throw new Error('Email user tidak ditemukan pada profile.')
      const { data, error } = await supabase.functions.invoke('admin-create-user', {
        body: {
          action: 'reset_password',
          target_user_id: access.users_profile?.id,
          target_email: email,
          password: resetForm.password,
          new_password: resetForm.password
        }
      })
      if (error) throw error
      if (data?.error) throw new Error(data.error)
      if (!data?.auth_user_id) throw new Error('Password belum terupdate di Supabase Auth. Deploy ulang Edge Function admin-create-user versi terbaru.')
      setMsg(`Password ${email} berhasil diganti. Auth ID: ${data.auth_user_id}`)
      setResetForm({ access_id:'', password:'' })
    } catch(e){ setError(e.message) }
    setSaving(false)
  }

  function downloadTemplate(){
    downloadXlsx('template-user-food-index.xlsx', [{ nama:'Budi Santoso', nrp:'NRP001', email:'budi@email.com', password:'Password123', role:'GL', site_code:'JIEP' }])
  }
  const tableRows = accessRows.map(r => ({
    email: r.users_profile?.email || '-',
    nama: r.users_profile?.nama || '-',
    nrp: r.users_profile?.nrp || '-',
    role: foodRoleFromAccess(r),
    db_role: r.role,
    site: r.sites?.site_code || 'All Site',
    status: r.status || '-'
  }))
  return <div className="stack food-admin-panel">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title="Tambah User Food Index" desc="Membuat akun login Supabase Auth sekaligus mapping akses Food Index." action={<button className="secondary" onClick={load} disabled={loading}>{loading ? 'Memuat...' : 'Refresh'}</button>}>
      <form className="form-grid" onSubmit={createOneUser}>
        <label>Nama<input value={manual.nama} onChange={e=>setManual({...manual,nama:e.target.value})} placeholder="Nama user" /></label>
        <label>NRP<input value={manual.nrp} onChange={e=>setManual({...manual,nrp:e.target.value})} placeholder="NRP / ID" /></label>
        <label>Email<input type="email" value={manual.email} onChange={e=>setManual({...manual,email:e.target.value})} placeholder="email@domain.com" /></label>
        <label>Password<input type="text" value={manual.password} onChange={e=>setManual({...manual,password:e.target.value})} placeholder="Minimal 6 karakter" /></label>
        <label>Role<select value={manual.role} onChange={e=>setManual({...manual,role:e.target.value,site_id:e.target.value==='Administrator'?'':manual.site_id})}>{FOOD_ACCESS_ROLES.map(r=><option key={r} value={r}>{r}</option>)}</select></label>
        <label>Site<select value={manual.site_id} onChange={e=>setManual({...manual,site_id:e.target.value})} disabled={manual.role==='Administrator'}><option value="">{manual.role === 'Administrator' ? 'All Site' : 'Pilih Site'}</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>
        <button disabled={saving}>{saving ? 'Menyimpan...' : 'Buat User Food Index'}</button>
      </form>
    </Panel>
    <Panel title="Upload User Food Index Excel" desc="Kolom: nama, nrp, email, password, role, site_code. Role: Administrator, GL, Section Head, Dept Head." action={<button className="secondary" onClick={downloadTemplate}><Download size={16}/> Download Template</button>}>
      <div className="upload-actions"><label className="upload-tile"><Upload size={20}/><b>Upload Excel</b><input type="file" accept=".xlsx,.xls" hidden onChange={e=>handleUploadExcel(e.target.files?.[0])}/></label>{!!uploadRows.length && <button onClick={importExcelUsers} disabled={saving}>{saving ? 'Importing...' : 'Submit Import Valid'}</button>}</div>
      {!!uploadRows.length && <div className="table-wrap" style={{maxHeight:320, overflow:'auto'}}><table><thead><tr><th>Row</th><th>Nama</th><th>NRP</th><th>Email</th><th>Role</th><th>Site Code</th><th>Status</th></tr></thead><tbody>{uploadRows.map(r=><tr key={r.row}><td>{r.row}</td><td>{r.nama}</td><td>{r.nrp}</td><td>{r.email}</td><td>{r.role}</td><td>{r.site_code || '-'}</td><td>{r.status}</td></tr>)}</tbody></table></div>}
    </Panel>
    <Panel title="Ganti Password User Food Index" desc="Administrator dapat mengganti password masing-masing akun Food Index. Password akan langsung diupdate ke Supabase Auth.">
      <form className="form-grid" onSubmit={resetUserPassword}>
        <label>User<select value={resetForm.access_id} onChange={e=>setResetForm({...resetForm, access_id:e.target.value})}><option value="">Pilih user</option>{accessRows.map(r => <option key={r.id} value={r.id}>{r.users_profile?.email || '-'} · {r.users_profile?.nama || '-'} · {foodRoleFromAccess(r)} · {r.sites?.site_code || 'All Site'}</option>)}</select></label>
        <label>Password Baru<input type="text" value={resetForm.password} onChange={e=>setResetForm({...resetForm,password:e.target.value})} placeholder="Minimal 6 karakter" /></label>
        <button disabled={saving || !resetForm.access_id}>{saving ? 'Menyimpan...' : 'Ganti Password'}</button>
      </form>
    </Panel>
    <Panel title="Row Data Akses Food Index" desc="Akses aktif akan muncul sebagai pilihan petugas tanda tangan inspeksi." action={<button className="secondary" onClick={()=>downloadXlsx('akses-user-food-index.xlsx', tableRows)}><Download size={16}/> Export</button>}>
      <Table rows={tableRows} columns={['email','nama','nrp','role','db_role','site','status']} empty="Belum ada akses Food Index." />
    </Panel>
  </div>
}

function FoodDashboard({ context }){
  const now = new Date()
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [siteFilter, setSiteFilter] = useState('')
  const [weekFilter, setWeekFilter] = useState('ALL')
  const [monthFilter, setMonthFilter] = useState('ALL')
  const [yearFilter, setYearFilter] = useState(String(now.getFullYear()))
  const [sites, setSites] = useState([])
  const [vendors, setVendors] = useState([])
  const [tasks, setTasks] = useState([])
  const [outs, setOuts] = useState([])
  const [parameters, setParameters] = useState([])
  const [signatureRows, setSignatureRows] = useState([])
  const weekEnd = weekFilter === 'ALL' ? '' : endOfWeekSunday(weekFilter)
  const weekLabel = weekFilter === 'ALL' ? 'All Week' : `${weekFilter} s/d ${weekEnd}`
  const monthLabel = monthFilter === 'ALL' ? 'All Month' : monthFilter
  const yearLabel = yearFilter === 'ALL' ? 'All Year' : yearFilter
  const selectedSiteLabel = siteFilter ? sites.find(s => s.id === siteFilter)?.site_code || 'Site terpilih' : 'All Site'
  const dashboardPeriodLabel = `Periode data: ${weekLabel} | Bulan: ${monthLabel} | Tahun: ${yearLabel} | Site: ${selectedSiteLabel}`
  const thursdayAlert = new Date().getDay() >= 4 || new Date().getDay() === 0

  useEffect(() => { load() }, [context?.id, siteFilter, weekFilter, monthFilter, yearFilter])

  async function load(){
    setLoading(true); setError('')
    try {
      let allDashboardSites = []
      let visibleDashboardSites = []
      let hiddenDashboardSiteIds = new Set()
      if (adminCanSeeAll(context)) {
        allDashboardSites = await fetchSites()
        visibleDashboardSites = allDashboardSites.filter(isFoodDashboardVisibleSite)
        hiddenDashboardSiteIds = new Set(allDashboardSites.filter(s => !isFoodDashboardVisibleSite(s)).map(s => s.id))
        setSites(visibleDashboardSites)
        if (siteFilter && !visibleDashboardSites.some(s => s.id === siteFilter)) setSiteFilter('')
      }
      let monthStart = null
      let monthEndDate = null
      if (yearFilter !== 'ALL' && monthFilter !== 'ALL') {
        monthStart = `${yearFilter}-${monthFilter}-01`
        monthEndDate = new Date(Number(yearFilter), Number(monthFilter), 0).toISOString().slice(0,10)
      } else if (yearFilter !== 'ALL') {
        monthStart = `${yearFilter}-01-01`
        monthEndDate = `${yearFilter}-12-31`
      }
      let tq = supabase
        .from('food_weekly_tasks')
        .select(`
          *,
          sites(site_code,site_name),
          food_vendors(id,vendor_name),
          food_inspection_answers(id, score, finding_note, evidence_photo_url, food_parameters(category, parameter_text, sort_order)),
          food_findings(id, due_date, corrective_action, preventive_action, status)
        `)
        .order('week_start_date', { ascending:false })
      if (monthStart && monthEndDate) tq = tq.gte('week_start_date', monthStart).lte('week_start_date', monthEndDate)
      let oq = supabase
        .from('food_outstandings')
        .select('*, food_findings(corrective_action, preventive_action, due_date, food_inspection_answers(finding_note, evidence_photo_url, food_parameters(parameter_text)), food_weekly_tasks(id, site_id, week_start_date, week_end_date, submitted_at, approved_at, status, sites(site_code,site_name), food_vendors(vendor_name)))')
        .order('created_at', { ascending:false })
        .limit(1000)
      let vq = supabase.from('food_vendors').select('id, site_id, vendor_name, status').eq('status','Aktif')
      let pq = supabase.from('food_parameters').select('id').eq('status','Aktif')
      let sq = supabase
        .from('food_inspection_signatures')
        .select('id, task_id, signer_user_id, site_id, signer_name, signer_nrp, signer_role, signed_at, food_weekly_tasks!inner(id, site_id, week_start_date, week_end_date, status, sites(site_code,site_name))')
        .order('signed_at', { ascending:false })
        .limit(5000)
      if (monthStart && monthEndDate) sq = sq.gte('food_weekly_tasks.week_start_date', monthStart).lte('food_weekly_tasks.week_start_date', monthEndDate)
      tq = applySiteFilter(tq, context)
      vq = applySiteFilter(vq, context)
      if (!adminCanSeeAll(context)) { oq = oq.eq('task_id', '00000000-0000-0000-0000-000000000000'); sq = sq.eq('food_weekly_tasks.site_id', context.site_id) }
      if (adminCanSeeAll(context) && siteFilter) { tq = tq.eq('site_id', siteFilter); vq = vq.eq('site_id', siteFilter); sq = sq.eq('food_weekly_tasks.site_id', siteFilter) }
      const [{ data:t, error:te }, { data:o, error:oe }, { data:v, error:ve }, { data:p, error:pe }, { data:sigs, error:se }] = await Promise.all([tq, oq, vq, pq, sq])
      if (te) throw te
      if (ve) throw ve
      if (pe) throw pe
      if (oe && adminCanSeeAll(context)) throw oe
      if (se && !String(se.message || '').includes('food_inspection_signatures')) throw se
      const hideDashboardSite = (siteId, siteCode) => hiddenDashboardSiteIds.has(siteId) || isFoodDashboardHiddenSiteCode(siteCode)
      const dashboardTasks = (t || []).filter(row => !hideDashboardSite(row.site_id, row.sites?.site_code))
      const dashboardOuts = (o || []).filter(row => {
        const task = row.food_findings?.food_weekly_tasks || {}
        if (hideDashboardSite(task.site_id, task.sites?.site_code)) return false
        if (adminCanSeeAll(context) && siteFilter && task.site_id !== siteFilter) return false
        return outstandingMatchesDashboardPeriod(row, yearFilter, monthFilter, weekFilter)
      })
      const dashboardVendors = (v || []).filter(row => !hideDashboardSite(row.site_id, ''))
      const dashboardSignatures = (sigs || []).filter(row => {
        const task = row.food_weekly_tasks || {}
        return !hideDashboardSite(row.site_id || task.site_id, task.sites?.site_code)
      })
      setTasks(dashboardTasks)
      setOuts(adminCanSeeAll(context) ? dashboardOuts : [])
      setVendors(dashboardVendors)
      setParameters(p || [])
      setSignatureRows(dashboardSignatures)
    } catch(e){ setError(e.message) }
    setLoading(false)
  }

  const filteredTasks = weekFilter === 'ALL' ? tasks : tasks.filter(t => t.week_start_date === weekFilter)
  const filteredOuts = outs.filter(o => outstandingMatchesDashboardPeriod(o, yearFilter, monthFilter, weekFilter))
  const filteredSignatureRows = weekFilter === 'ALL' ? signatureRows : signatureRows.filter(sig => {
    const task = sig.food_weekly_tasks || {}
    return task.week_start_date === weekFilter
  })
  const weekTasks = filteredTasks
  const total = weekTasks.length
  const approved = weekTasks.filter(t => t.status === 'Approved').length
  const expired = weekTasks.filter(t => t.status === 'Expired').length
  const achievement = pctValue(approved, total)
  const alertRows = weekTasks.filter(t => thursdayAlert && ['Open','In Progress','Need Action Plan'].includes(t.status))
  const reportRows = filteredOuts.slice(0, 8).map((o, idx) => {
    const f = o.food_findings || {}
    const answer = f.food_inspection_answers || {}
    const task = f.food_weekly_tasks || {}
    return {
      no: idx + 1,
      parameter: answer.food_parameters?.parameter_text || '-',
      caption: answer.finding_note || '-',
      tanggal: task.submitted_at?.slice(0,10) || task.week_start_date || '-',
      corrective: f.corrective_action || '-',
      preventive: f.preventive_action || '-',
      due_date: f.due_date || '-',
      foto_corrective: o.corrective_photo_url ? 'Lihat foto' : '-',
      foto_preventive: o.preventive_photo_url ? 'Lihat foto' : '-',
      status: o.status === 'Closed' ? 'Closed' : 'Open'
    }
  })

  const scoreRows = filteredTasks
    .filter(t => taskHasActualInspectionWork(t) || ['Approved','Expired','Waiting Approval','Need Action Plan'].includes(String(t.status || '')))
    .map((t, idx) => {
      const answers = t.food_inspection_answers || []
      const totalParam = answers.length || parameters.length || 0
      const ok = answers.filter(a => Number(a.score) === 1).length
      const score = pctValue(ok, totalParam)
      const display = scoreInspectionDisplay(t, score)
      return {
        no: idx + 1,
        site: t.sites?.site_code || '-',
        vendor: t.food_vendors?.vendor_name || '-',
        minggu: `${t.week_start_date} s/d ${t.week_end_date}`,
        parameter: totalParam,
        sesuai: ok,
        temuan: answers.filter(a => Number(a.score) === 0).length,
        score: fmtPct(score),
        remark: display.remark,
        status: display.status
      }
    })
  const executionScopeTasks = filteredTasks
  const executionTotal = executionScopeTasks.length || (weekFilter !== 'ALL' ? vendors.length : 0)
  const inspectedExecutionTasks = executionScopeTasks.filter(taskHasActualInspectionWork)
  const executionDone = inspectedExecutionTasks.length
  const notInspectedVendorCount = Math.max(executionTotal - executionDone, 0)
  const executionScore = pctValue(executionDone, executionTotal)
  const executionRows = [{
    periode: `${weekLabel} | ${monthLabel}-${yearLabel}`,
    bulan: monthLabel,
    week: weekLabel,
    site: siteFilter ? selectedSiteLabel : adminCanSeeAll(context) ? 'All Site' : contextSiteName(context),
    total_vendor_aktif: executionTotal,
    vendor_sudah_inspeksi: executionDone,
    vendor_belum_inspeksi: notInspectedVendorCount,
    skor_pelaksanaan: fmtPct(executionScore),
    remark: executionScore >= 100 ? 'Tercapai' : 'Tidak tercapai'
  }]
  const leadtimeRows = filteredOuts.filter(o => o.food_findings?.due_date).map((o, idx) => {
    const f = o.food_findings || {}
    const task = f.food_weekly_tasks || {}
    const closedDate = o.closed_at?.slice(0,10) || ''
    const due = f.due_date || ''
    const onTime = closedDate ? closedDate <= due : today() <= due
    return {
      no: idx + 1,
      site: task.sites?.site_code || '-',
      vendor: task.food_vendors?.vendor_name || '-',
      parameter: f.food_inspection_answers?.food_parameters?.parameter_text || '-',
      minggu_outstanding: outstandingCreatedWeekStart(o) || task.week_start_date || '-',
      due_date: due || '-',
      tanggal_close: closedDate || '-',
      status: o.status === 'Closed' ? 'Closed' : 'Open',
      remark: onTime ? 'Tercapai' : 'Tidak tercapai'
    }
  })

  function pct(n, d){ return pctValue(n, d) }
  // sites state pada dashboard admin sudah disaring agar Head Office/JIEP tidak tampil di resume dan achievement operasional.
  const siteSource = adminCanSeeAll(context) ? sites : [{ id: context?.site_id, site_code: contextSiteName(context), site_name: contextSiteName(context) }]
  const roleParticipationRows = siteSource
    .filter(s => !siteFilter || s.id === siteFilter)
    .map((s, idx) => {
      const siteSignatures = filteredSignatureRows.filter(sig => {
        const task = sig.food_weekly_tasks || {}
        return task.site_id === s.id && sameMonth(task.week_start_date, yearFilter, monthFilter)
      })
      const counts = {}
      Object.keys(FOOD_ROLE_TARGETS).forEach(role => { counts[role] = siteSignatures.filter(sig => normalizeRole(sig.signer_role) === role).length })
      const remarks = Object.entries(FOOD_ROLE_TARGETS).map(([role, target]) => `${role}: ${counts[role] || 0}/${target} ${counts[role] >= target ? 'Tercapai' : 'Belum tercapai'}`)
      const allOk = Object.entries(FOOD_ROLE_TARGETS).every(([role, target]) => (counts[role] || 0) >= target)
      return {
        no: idx + 1,
        site: s.site_code || '-',
        site_name: s.site_name || '-',
        gl: `${counts.GL || 0}/4`,
        section_head: `${counts['Section Head'] || 0}/2`,
        dept_head: `${counts['Dept Head'] || 0}/1`,
        status: allOk ? 'Tercapai' : 'Belum tercapai',
        remark: remarks.join(' | ')
      }
    })
  const resumeRows = siteSource
    .filter(s => !siteFilter || s.id === siteFilter)
    .map((s, idx) => {
      const siteTasks = filteredTasks.filter(t => t.site_id === s.id)
      const inspectedSiteTasks = siteTasks.filter(taskHasActualInspectionWork)
      const siteAnswers = inspectedSiteTasks.flatMap(t => t.food_inspection_answers || [])
      const siteOkAnswers = siteAnswers.filter(a => Number(a.score) === 1).length
      const scoreFoodIndex = pct(siteOkAnswers, siteAnswers.length)
      const siteExecutionTotal = siteTasks.length || (weekFilter !== 'ALL' ? vendors.filter(v => v.site_id === s.id).length : 0)
      const skorPelaksanaan = pct(inspectedSiteTasks.length, siteExecutionTotal)
      const siteLeadtime = leadtimeRows.filter(r => r.site === s.site_code)
      const leadtimeScore = siteLeadtime.length ? pct(siteLeadtime.filter(r => r.remark === 'Tercapai').length, siteLeadtime.length) : 0
      const available = [scoreFoodIndex, skorPelaksanaan, ...(siteLeadtime.length ? [leadtimeScore] : [])]
      const achievement = available.length ? available.reduce((a,b)=>a+b,0) / available.length : 0
      return {
        no: idx + 1,
        site: s.site_code || '-',
        site_name: s.site_name || '-',
        score_food_index: fmtPct(scoreFoodIndex),
        skor_pelaksanaan: fmtPct(skorPelaksanaan),
        leadtime_perbaikan: siteLeadtime.length ? fmtPct(leadtimeScore) : '-',
        achievement_site: fmtPct(achievement),
        remark: achievement >= 95 ? 'Tercapai' : 'Tidak tercapai'
      }
    })

  const years = ['ALL', ...Array.from({ length: 6 }, (_, i) => String(now.getFullYear() - i))]
  const monthOptions = ['ALL', ...Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'))]
  const weekOptions = ['ALL']
  if (yearFilter !== 'ALL' && monthFilter !== 'ALL') {
    for (let d = new Date(Number(yearFilter), Number(monthFilter) - 1, 1); d.getMonth() === Number(monthFilter) - 1; d.setDate(d.getDate() + 7)) {
      const w = startOfWeekMonday(new Date(d))
      if (!weekOptions.includes(w)) weekOptions.push(w)
    }
  } else {
    tasks.forEach(t => { if (t.week_start_date && !weekOptions.includes(t.week_start_date)) weekOptions.push(t.week_start_date) })
  }
  if (weekFilter !== 'ALL' && !weekOptions.includes(weekFilter)) weekOptions.push(weekFilter)

  return <div className="food-dashboard-v43">
    {error && <div className="error">{error}</div>}
    <div className="dashboard-filter-card">
      <div>
        <small>Scope dashboard</small>
        <h3>{siteFilter ? sites.find(s => s.id === siteFilter)?.site_code || 'Site terpilih' : 'All Site'}</h3>
        <p>Dashboard mengikuti filter week, bulan, tahun, dan site. Achievement hanya bertambah setelah inspeksi full approved.</p>
      </div>
      <div className="dashboard-admin-filters">
        <label>Week<select value={weekFilter} onChange={e=>setWeekFilter(e.target.value)}>{weekOptions.map(w=><option key={w} value={w}>{w === 'ALL' ? 'All Week' : `${w} s/d ${endOfWeekSunday(w)}`}</option>)}</select></label>
        <label>Bulan<select value={monthFilter} onChange={e=>setMonthFilter(e.target.value)}>{monthOptions.map(m=><option key={m} value={m}>{m === 'ALL' ? 'All Month' : m}</option>)}</select></label>
        <label>Tahun<select value={yearFilter} onChange={e=>setYearFilter(e.target.value)}>{years.map(y=><option key={y} value={y}>{y === 'ALL' ? 'All Year' : y}</option>)}</select></label>
        {adminCanSeeAll(context) && <label>Site<select value={siteFilter} onChange={e=>setSiteFilter(e.target.value)}><option value="">All Site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>}
        <button onClick={load}>⟳ Refresh</button>
      </div>
      <div className="dashboard-tags"><span>Site: {siteFilter ? sites.find(s => s.id === siteFilter)?.site_code || 'Terpilih' : 'All Site'}</span><span>Week: {weekLabel}</span><span>Bulan: {monthLabel}</span><span>Tahun: {yearLabel}</span>{thursdayAlert && <span>Alert Kamis aktif</span>}</div>
    </div>

    <div className="dashboard-kpi-grid">
      <div className="dash-kpi"><small>Task Minggu Ini</small><strong>{total}</strong><span>↗ {sites.length || 1} site aktif</span></div>
      <div className="dash-kpi"><small>Approved</small><strong>{approved}</strong><span>Achievement {fmtPct(achievement)}</span></div>
      <div className="dash-kpi"><small>Outstanding Open</small><strong>{filteredOuts.filter(o => o.status !== 'Closed').length}</strong><span className="orange">Butuh follow up</span></div>
      <div className="dash-kpi"><small>Expired</small><strong>{expired}</strong><span className="red">Impact achievement</span></div>
    </div>

    {adminCanSeeAll(context) && <section className="dashboard-card analysis-table compact-card role-participation-card">
      <div className="dashboard-card-head"><div><h3>Partisipasi Inspeksi per Role</h3><p>Target per site per bulan dihitung dari user yang ikut tanda tangan inspeksi: GL 4x, Section Head 2x, Dept Head 1x.</p></div><button onClick={()=>downloadXlsx('partisipasi-role-food-index.xlsx', roleParticipationRows)}>Export</button></div>
      <Table rows={roleParticipationRows} columns={['no','site','site_name','gl','section_head','dept_head','status','remark']} empty="Belum ada data tanda tangan pada filter ini." />
    </section>}

    {adminCanSeeAll(context) && <section className="dashboard-card analysis-table resume-table compact-card">
      <div className="dashboard-card-head"><div><h3>Resume Achievement per Site</h3><p>Ringkasan 3 dashboard utama per site: Score Food Index, Skor Pelaksanaan Inspeksi, dan Leadtime Perbaikan.</p><p className="period-note">{dashboardPeriodLabel}</p></div><button onClick={()=>downloadXlsx('resume-achievement-food-index-per-site.xlsx', resumeRows)}>Export</button></div>
      <Table rows={resumeRows} columns={['no','site','site_name','score_food_index','skor_pelaksanaan','leadtime_perbaikan','achievement_site','remark']} empty="Belum ada data resume." captureAll />
    </section>}

    <div className={adminCanSeeAll(context) ? 'dashboard-analysis-grid' : 'dashboard-analysis-grid site-role-dashboard-grid'}>
      <section className="dashboard-card analysis-table">
        <div className="dashboard-card-head"><div><h3>Score Food Index per Inspeksi</h3><p>Dihitung dari jumlah parameter sesuai dibanding total parameter inspeksi. Score ≥ 95% = Tercapai.</p></div><button onClick={()=>downloadXlsx('score-food-index.xlsx', scoreRows)}>Export</button></div>
        <Table rows={scoreRows} columns={['no','site','vendor','minggu','parameter','sesuai','temuan','score','remark','status']} empty="Belum ada inspeksi pada filter ini." />
      </section>
      {adminCanSeeAll(context) && <section className="dashboard-card analysis-table">
        <div className="dashboard-card-head"><div><h3>Skor Pelaksanaan Inspeksi</h3><p>Jumlah vendor/task inspeksi yang sudah dilakukan dibanding total vendor/task yang terbentuk pada scope filter.</p><p className="period-note">{dashboardPeriodLabel}</p></div><button onClick={()=>downloadXlsx('skor-pelaksanaan-inspeksi.xlsx', executionRows)}>Export</button></div>
        <Table rows={executionRows} />
      </section>}
      {adminCanSeeAll(context) && <section className="dashboard-card analysis-table">
        <div className="dashboard-card-head"><div><h3>Leadtime Perbaikan</h3><p>Dihitung berdasarkan minggu saat outstanding dibuat. Remark Tercapai jika close outstanding tidak melewati due date.</p></div><button onClick={()=>downloadXlsx('leadtime-perbaikan-food-index.xlsx', leadtimeRows)}>Export</button></div>
        <Table rows={leadtimeRows} empty="Belum ada due date perbaikan." />
      </section>}
    </div>

    <div className="dashboard-main-grid">
      <section className="dashboard-card report-preview-card">
        <div className="dashboard-card-head">
          <div><h3>Report Temuan Food Index</h3><p>Hanya menampilkan parameter yang memiliki temuan. Siap untuk view dan export Excel.</p></div>
          <button onClick={()=>downloadXlsx('food-index-report-temuan-preview.xlsx', reportRows)}><Download size={16}/> Export Excel</button>
        </div>
        <div className="dashboard-table-wrap">
          <table>
            <thead><tr><th>No</th><th>Parameter Checklist</th><th>Caption Temuan</th><th>Tgl Inspeksi</th><th>Corrective Action</th><th>Preventive Action</th><th>Due Date</th><th>Foto Corrective</th><th>Foto Preventive</th><th>Status</th></tr></thead>
            <tbody>{reportRows.map(r => <tr key={r.no}><td>{r.no}</td><td>{r.parameter}</td><td>{r.caption}</td><td>{r.tanggal}</td><td>{r.corrective}</td><td>{r.preventive}</td><td>{r.due_date}</td><td>{r.foto_corrective}</td><td>{r.foto_preventive}</td><td><StatusPill value={r.status} /></td></tr>)}</tbody>
          </table>
          {!reportRows.length && <p className="muted table-empty">Belum ada temuan untuk ditampilkan.</p>}
        </div>
      </section>

      <aside className="dashboard-card alert-card-v43">
        <h3>Alert Mingguan</h3>
        <p>Mulai Kamis, site yang belum inspeksi ditampilkan sebagai prioritas.</p>
        <div className="alert-stack-v43">
          {loading ? <div className="mini-alert">Memuat alert...</div> : alertRows.length ? alertRows.slice(0, 6).map((t, i) => <div className="mini-alert" key={t.id || i}><div className="mini-alert-icon"><AlertTriangle size={18}/></div><div><b>{t.sites?.site_code || '-'} belum inspeksi</b><span>{t.food_vendors?.vendor_name || 'Vendor'} masih status {t.status}</span></div></div>) : <div className="mini-alert"><div className="mini-alert-icon"><CheckCircle2 size={18}/></div><div><b>Tidak ada alert</b><span>Semua task minggu ini aman sesuai filter.</span></div></div>}
          {expired > 0 && <div className="mini-alert"><div className="mini-alert-icon">⏳</div><div><b>{expired} task expired</b><span>Task expired menurunkan achievement site.</span></div></div>}
        </div>
      </aside>
    </div>
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
    const siteList = await fetchSites()
    const siteMap = new Map(siteList.map(s => [String(s.site_code).toUpperCase(), s]))
    const { data: existingData, error: existingError } = await supabase
      .from('food_vendors')
      .select('id, site_id, vendor_name, status, sites(site_code)')
      .eq('status', 'Aktif')
    if (existingError) throw existingError

    const normalizeKey = (siteId, name) => `${siteId || ''}::${cleanText(name).toUpperCase().replace(/\s+/g, ' ')}`
    const existingActive = new Set((existingData || []).map(v => normalizeKey(v.site_id, v.vendor_name)))
    const seenInFile = new Map()

    const mapped = rows.map((r, idx) => {
      const siteCode = cleanText(getVal(r, ['site_code','SITE_CODE','Site Code'])).toUpperCase()
      const site = siteMap.get(siteCode)
      const vendor_name = cleanText(getVal(r, ['vendor_name','VENDOR_NAME','nama_vendor','Nama Vendor']))
      const status = cleanText(getVal(r, ['status','STATUS'])) || 'Aktif'
      const isActive = String(status).toLowerCase() === 'aktif'
      const key = normalizeKey(site?.id, vendor_name)
      let err = !site ? 'site_code tidak ditemukan' : !vendor_name ? 'vendor_name wajib diisi' : ''
      if (!err && isActive) {
        if (seenInFile.has(key)) err = `Duplicate di Excel dengan baris ${seenInFile.get(key)} untuk site ${siteCode}`
        else if (existingActive.has(key)) err = `Vendor aktif sudah ada di database untuk site ${siteCode}`
      }
      if (!seenInFile.has(key)) seenInFile.set(key, idx+2)
      return { row: idx+2, site_code: siteCode, site_id: site?.id, vendor_name, status, error: err }
    })
    const errCount = mapped.filter(r => r.error).length
    if (errCount) setError(`Preview menemukan ${errCount} baris bermasalah. Perbaiki baris yang tidak valid sebelum submit.`)
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
  const emptyForm = { id:null, parameter_code:'', category:'General', parameter_text:'', standard_parameter:'', hazard_code:'', behavior:'', sort_order:1, status:'Aktif' }
  const [form, setForm] = useState(emptyForm)
  const [preview, setPreview] = useState([])
  const [msg, setMsg] = useState('')
  const [error, setError] = useState('')
  const editing = Boolean(form.id)

  useEffect(() => { load() }, [])
  async function load(){
    const { data, error } = await supabase.from('food_parameters').select('*').order('sort_order')
    if (error) setError(error.message)
    setRows(data || [])
  }

  function resetForm(){
    setForm({ ...emptyForm, sort_order: (rows.length + 1) || 1 })
  }

  async function save(e){
    e.preventDefault(); setMsg(''); setError('')
    try {
      if (!cleanText(form.parameter_text)) throw new Error('Parameter wajib diisi.')
      const payload = {
        category: cleanText(form.category) || 'General',
        parameter_text: cleanText(form.parameter_text),
        standard_parameter: cleanText(form.standard_parameter) || null,
        hazard_code: cleanText(form.hazard_code) || null,
        behavior: cleanText(form.behavior) || null,
        sort_order: Number(form.sort_order) || 1,
        status: form.status || 'Aktif'
      }
      if (editing) {
        const { error } = await supabase.from('food_parameters').update(payload).eq('id', form.id)
        if (error) throw error
        setMsg('Parameter berhasil diupdate.')
      } else {
        const { error } = await supabase.from('food_parameters').insert(payload)
        if (error) throw error
        setMsg('Parameter berhasil disimpan.')
      }
      resetForm(); load()
    } catch(e){ setError(e.message) }
  }

  function editRow(r){
    setForm({
      id:r.id,
      parameter_code:r.parameter_code || '',
      category:r.category || 'General',
      parameter_text:r.parameter_text || '',
      standard_parameter:r.standard_parameter || '',
      hazard_code:r.hazard_code || '',
      behavior:r.behavior || '',
      sort_order:r.sort_order || 1,
      status:r.status || 'Aktif'
    })
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }

  async function deleteRow(r){
    setMsg(''); setError('')
    try {
      if (!confirm(`Nonaktifkan parameter "${r.parameter_text}"?`)) return
      const { error } = await supabase.from('food_parameters').update({ status:'Nonaktif' }).eq('id', r.id)
      if (error) throw error
      setMsg('Parameter berhasil dinonaktifkan. Data histori inspeksi tetap aman.'); load()
    } catch(e){ setError(e.message) }
  }

  function downloadTemplate(){
    downloadXlsx('template-master-parameter-food-index.xlsx', [{
      id:'', parameter_code:'', category:'Kebersihan', parameter_text:'Area penyajian bersih dan rapi',
      standard_parameter:'Area bebas sampah, lantai kering, tidak licin', hazard_code:'FIP-HAZ-01',
      behavior:'Housekeeping dan hygiene vendor', sort_order:1, status:'Aktif'
    }])
  }

  function downloadExisting(){
    const existing = rows.map(r => ({
      id:r.id,
      parameter_code:r.parameter_code || '',
      category:r.category || '',
      parameter_text:r.parameter_text || '',
      standard_parameter:r.standard_parameter || '',
      hazard_code:r.hazard_code || '',
      behavior:r.behavior || '',
      sort_order:r.sort_order || 1,
      status:r.status || 'Aktif'
    }))
    downloadXlsx('food-index-parameters-existing-update.xlsx', existing)
  }

  async function parseExcel(file){
    setMsg(''); setError('')
    const buf = await file.arrayBuffer()
    const wb = XLSX.read(buf)
    const seen = new Map()
    const items = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]).map(normalizeRow).map((r, idx) => {
      const id = cleanText(getVal(r, ['id','ID']))
      const parameter_code = cleanText(getVal(r, ['parameter_code','PARAMETER_CODE','kode_parameter','Kode Parameter']))
      const parameter_text = cleanText(getVal(r, ['parameter_text','PARAMETER_TEXT','parameter','Parameter','soal','Soal']))
      const key = (id || parameter_code || parameter_text).toUpperCase()
      let error = parameter_text ? '' : 'parameter_text wajib diisi'
      if (!error && seen.has(key)) error = `Duplicate di Excel dengan baris ${seen.get(key)}`
      if (!seen.has(key)) seen.set(key, idx+2)
      return {
        row:idx+2,
        id,
        parameter_code,
        category:cleanText(getVal(r, ['category','CATEGORY','kategori','Kategori'])) || 'General',
        parameter_text,
        standard_parameter: cleanText(getVal(r, ['standard_parameter','STANDARD_PARAMETER','standar_parameter','Standar Parameter','standar','Standar'])),
        hazard_code: cleanText(getVal(r, ['hazard_code','HAZARD_CODE','kode_bahaya','Kode Bahaya','kode bahaya'])),
        behavior: cleanText(getVal(r, ['behavior','BEHAVIOR','perilaku','Perilaku'])),
        sort_order:Number(getVal(r, ['sort_order','SORT_ORDER','urutan','Urutan'])) || idx+1,
        status:cleanText(getVal(r, ['status','STATUS'])) || 'Aktif',
        error
      }
    })
    const errCount = items.filter(r => r.error).length
    if (errCount) setError(`Preview menemukan ${errCount} baris bermasalah. Perbaiki dulu sebelum submit.`)
    setPreview(items)
  }

  async function submitImport(){
    setMsg(''); setError('')
    try {
      const valid = preview.filter(r => !r.error).map(({row,error,...r}) => ({
        ...r,
        standard_parameter: cleanText(r.standard_parameter) || null,
        hazard_code: cleanText(r.hazard_code) || null,
        behavior: cleanText(r.behavior) || null
      }))
      if (!valid.length) throw new Error('Tidak ada parameter valid.')
      const idMap = new Map(rows.map(r => [r.id, r]))
      const codeMap = new Map(rows.filter(r => r.parameter_code).map(r => [String(r.parameter_code).toUpperCase(), r]))
      let updated = 0
      let inserted = 0
      for (const r of valid) {
        const payload = {
          category:r.category || 'General',
          parameter_text:r.parameter_text,
          standard_parameter:r.standard_parameter,
          hazard_code:r.hazard_code,
          behavior:r.behavior,
          sort_order:Number(r.sort_order) || 1,
          status:r.status || 'Aktif'
        }
        const targetById = r.id && idMap.get(r.id)
        const targetByCode = !targetById && r.parameter_code && codeMap.get(String(r.parameter_code).toUpperCase())
        if (targetById) {
          const { error } = await supabase.from('food_parameters').update(payload).eq('id', r.id)
          if (error) throw error
          updated += 1
        } else if (targetByCode) {
          const { error } = await supabase.from('food_parameters').update(payload).eq('id', targetByCode.id)
          if (error) throw error
          updated += 1
        } else {
          const { error } = await supabase.from('food_parameters').insert(payload)
          if (error) throw error
          inserted += 1
        }
      }
      setMsg(`Import selesai: ${updated} parameter diupdate, ${inserted} parameter baru tersimpan.`); setPreview([]); load()
    } catch(e){ setError(e.message) }
  }

  const tableRows = rows.map(r => ({ id:r.id, category:r.category, parameter_code:r.parameter_code, parameter_text:r.parameter_text, standard_parameter:r.standard_parameter || '-', hazard_code:r.hazard_code || '-', behavior:r.behavior || '-', sort_order:r.sort_order, status:r.status }))

  return <div className="stack">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title={editing ? 'Edit Parameter' : 'Tambah Parameter Manual'} desc="Jawaban inspeksi tetap hanya 1 atau 0. Field standar parameter, kode bahaya, dan behavior dipakai sebagai referensi saat inspeksi.">
      <form className="form-grid" onSubmit={save}>
        {editing && <label>Kode Parameter<input value={form.parameter_code} disabled /></label>}
        <label>Kategori<input value={form.category} onChange={e=>setForm({...form, category:e.target.value})} /></label>
        <label>Parameter Checklist<input value={form.parameter_text} onChange={e=>setForm({...form, parameter_text:e.target.value})} placeholder="Isi parameter inspeksi" /></label>
        <label>Standar Parameter<input value={form.standard_parameter} onChange={e=>setForm({...form, standard_parameter:e.target.value})} placeholder="Contoh: Area bersih, kering, bebas kontaminasi" /></label>
        <label>Kode Bahaya<input value={form.hazard_code} onChange={e=>setForm({...form, hazard_code:e.target.value})} placeholder="Contoh: HZ-FOOD-001" /></label>
        <label>Behavior<input value={form.behavior} onChange={e=>setForm({...form, behavior:e.target.value})} placeholder="Contoh: Hygiene personil / housekeeping" /></label>
        <label>Urutan<input type="number" value={form.sort_order} onChange={e=>setForm({...form, sort_order:e.target.value})} /></label>
        <label>Status<select value={form.status} onChange={e=>setForm({...form, status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label>
        <button>{editing ? 'Update Parameter' : 'Simpan Parameter'}</button>
        {editing && <button type="button" className="secondary" onClick={resetForm}>Batal Edit</button>}
      </form>
    </Panel>
    <Panel title="Upload / Update Parameter Excel" desc="Download data existing, lengkapi kolom kosong, lalu upload ulang untuk update. Kolom: id, parameter_code, category, parameter_text, standard_parameter, hazard_code, behavior, sort_order, status." action={<div className="panel-actions"><button className="secondary" onClick={downloadTemplate}><Download size={16}/> Download Template</button><button className="secondary" onClick={downloadExisting}><Download size={16}/> Download Data Existing</button></div>}>
      <div className="import-row"><label className="upload-line"><Upload size={20}/><span>Upload Excel Update</span><input type="file" accept=".xlsx,.xls" hidden onChange={e=>e.target.files?.[0] && parseExcel(e.target.files[0])}/></label>{preview.length > 0 && <button onClick={submitImport}>Submit Import / Update Valid</button>}</div>
      {preview.length > 0 && <Table rows={preview.map(r=>({ row:r.row, id:r.id || '-', parameter_code:r.parameter_code || '-', category:r.category, parameter_text:r.parameter_text, standard_parameter:r.standard_parameter || '-', hazard_code:r.hazard_code || '-', behavior:r.behavior || '-', sort_order:r.sort_order, status:r.status, valid:r.error || 'Valid' }))} />}
    </Panel>
    <Panel title="Row Data Parameter Food Index" action={<button onClick={()=>downloadXlsx('food-index-parameters.xlsx', tableRows)}><Download size={16}/> Export</button>}>
      <div className="table-toolbar"><div className="searchbox"><Search size={18}/><input placeholder="Search row data..." onChange={e=>{ const q=e.target.value.toLowerCase(); document.querySelectorAll('.parameter-row').forEach(tr=>{ tr.style.display = !q || tr.innerText.toLowerCase().includes(q) ? '' : 'none' }) }} /></div></div>
      <div className="table-wrap" style={{maxHeight: 460, overflow:'auto'}}>
        <table>
          <thead><tr><th>Parameter Code</th><th>Category</th><th>Parameter Text</th><th>Standard Parameter</th><th>Hazard Code</th><th>Behavior</th><th>Sort Order</th><th>Status</th><th>Aksi</th></tr></thead>
          <tbody>{rows.map(r => <tr className="parameter-row" key={r.id}><td>{r.parameter_code || '-'}</td><td>{r.category}</td><td>{r.parameter_text}</td><td>{r.standard_parameter || '-'}</td><td>{r.hazard_code || '-'}</td><td>{r.behavior || '-'}</td><td>{r.sort_order}</td><td><StatusPill value={r.status} /></td><td className="action-cell"><button className="secondary small" onClick={()=>editRow(r)}>Edit</button><button className="danger small" onClick={()=>deleteRow(r)}>Delete</button></td></tr>)}</tbody>
        </table>
        {!rows.length && <p className="muted table-empty">Belum ada parameter.</p>}
      </div>
    </Panel>
  </div>
}
const FOOD_IMAGE_MAX_DIMENSION = 1280
const FOOD_IMAGE_QUALITY = 0.76
const FOOD_IMAGE_COMPRESS_MIN_SIZE = 350 * 1024

function loadImageForCompression(file){
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file)
    const img = new Image()
    img.onload = () => {
      URL.revokeObjectURL(url)
      resolve(img)
    }
    img.onerror = () => {
      URL.revokeObjectURL(url)
      reject(new Error('Gagal membaca file foto untuk kompresi.'))
    }
    img.src = url
  })
}

async function compressFoodImage(file){
  if (!file || !file.type?.startsWith('image/')) return file
  if (file.type === 'image/gif' || file.type === 'image/svg+xml') return file
  if (file.size > 0 && file.size < FOOD_IMAGE_COMPRESS_MIN_SIZE) return file

  try {
    const img = await loadImageForCompression(file)
    const originalWidth = img.naturalWidth || img.width
    const originalHeight = img.naturalHeight || img.height
    if (!originalWidth || !originalHeight) return file

    const scale = Math.min(1, FOOD_IMAGE_MAX_DIMENSION / Math.max(originalWidth, originalHeight))
    const width = Math.max(1, Math.round(originalWidth * scale))
    const height = Math.max(1, Math.round(originalHeight * scale))

    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const ctx = canvas.getContext('2d')
    if (!ctx) return file

    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, width, height)
    ctx.drawImage(img, 0, 0, width, height)

    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', FOOD_IMAGE_QUALITY))
    if (!blob || blob.size <= 0) return file

    if (file.size && blob.size >= file.size) return file
    const baseName = (file.name || 'food-photo').replace(/\.[^/.]+$/, '')
    return new File([blob], `${baseName}-compressed.jpg`, { type:'image/jpeg', lastModified: Date.now() })
  } catch (err) {
    console.warn('Food image compression skipped:', err)
    return file
  }
}

async function uploadFoodImage(file, folder='evidence'){
  if (!file) return ''
  const shouldCompress = folder !== 'signatures'
  const uploadFile = shouldCompress ? await compressFoodImage(file) : file
  const ext = (uploadFile.name?.split('.').pop() || 'jpg').toLowerCase()
  const safeExt = ['jpg','jpeg','png','webp'].includes(ext) ? ext : 'jpg'
  const path = `${folder}/${new Date().toISOString().slice(0,10)}/${crypto.randomUUID()}.${safeExt}`
  const { error } = await supabase.storage.from('food-index-assets').upload(path, uploadFile, { upsert:false, contentType:uploadFile.type || 'image/jpeg' })
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
  const [signatures, setSignatures] = useState([])
  const [signerForm, setSignerForm] = useState({ signer_user_id:'', site_id:'', signer_name:'', signer_nrp:'', signer_role:'', signer_company:'', signatureDataUrl:'' })
  const [signerSearch, setSignerSearch] = useState('')
  const [eligibleSigners, setEligibleSigners] = useState([])
  const [signaturePadKey, setSignaturePadKey] = useState(0)
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
      const [{ data: existing, error: ae }, { data: existingSignatures, error: se }] = await Promise.all([
        supabase.from('food_inspection_answers')
          .select('*, food_findings(corrective_action, preventive_action)')
          .eq('task_id', currentTask.id),
        supabase.from('food_inspection_signatures')
          .select('*')
          .eq('task_id', currentTask.id)
          .order('signed_at', { ascending:true })
      ])
      if (ae) throw ae
      if (se && !String(se.message || '').includes('food_inspection_signatures')) throw se
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
      setSignatures((existingSignatures || []).map(row => ({
        id: row.id,
        signer_user_id: row.signer_user_id || '',
        site_id: row.site_id || '',
        signer_name: row.signer_name || '',
        signer_nrp: row.signer_nrp || '',
        signer_role: row.signer_role || '',
        signer_company: row.signer_company || '',
        signatureUrl: row.signature_url || '',
        existing: true
      })))
      setSignerForm({ signer_user_id:'', site_id:'', signer_name:'', signer_nrp:'', signer_role:'', signer_company:'', signatureDataUrl:'' })
      setSignerSearch('')
      setEligibleSigners(await loadEligibleSigners(currentTask.site_id))
      setSignaturePadKey(prev => prev + 1)
      setActiveTask(currentTask)
      setTasks(await fetchTasks())
    } catch(e){ setError(e.message) }
  }
  function setAnswer(parameterId, patch){
    setAnswers(prev => ({ ...prev, [parameterId]: { ...(prev[parameterId] || {}), ...patch } }))
  }
  async function loadEligibleSigners(siteId){
    try {
      const foodApp = await getFoodIndexApplication()
      let q = supabase.from('user_app_access')
        .select('user_id,role,food_role_level,site_id,status,users_profile(id,nama,email,nrp,status),sites(site_code,site_name)')
        .eq('app_id', foodApp.id)
        .eq('status', 'Aktif')
        .in('role', FOOD_DB_ACCESS_ROLES)
        .limit(1000)
      if (!adminCanSeeAll(context) && siteId) q = q.eq('site_id', siteId)
      const { data, error } = await q
      if (error) throw error
      return (data || []).filter(r => {
        const roleLevel = foodRoleFromAccess(r)
        if (roleLevel === 'Administrator') return true
        return !siteId || r.site_id === siteId
      })
    } catch(e) {
      setError(`Gagal memuat daftar user tanda tangan: ${e.message}`)
      return []
    }
  }
  function selectUserSigner(row){
    const profile = row?.users_profile || {}
    setSignerForm({
      signer_user_id: row?.user_id || profile?.id || '',
      site_id: row?.site_id || activeTask?.site_id || '',
      signer_name: profile?.nama || profile?.email || '',
      signer_nrp: profile?.nrp || '',
      signer_role: foodRoleFromAccess(row),
      signer_company: row?.sites?.site_code || activeTask?.sites?.site_code || contextSiteName(context) || '',
      signatureDataUrl: ''
    })
    setSignaturePadKey(prev => prev + 1)
  }
  const filteredEligibleSigners = eligibleSigners.filter(r => {
    const term = signerSearch.toLowerCase().trim()
    if (!term) return false
    const p = r.users_profile || {}
    return String(p.nama || '').toLowerCase().includes(term) || String(p.nrp || '').toLowerCase().includes(term) || String(p.email || '').toLowerCase().includes(term)
  }).slice(0, 12)
  function addSignature(){
    const name = cleanText(signerForm.signer_name)
    if (!name) { setError('Nama petugas tanda tangan wajib diisi.'); return }
    if (!cleanText(signerForm.signatureDataUrl)) { setError('Tanda tangan canvas wajib diisi sebelum ditambahkan.'); return }
    setSignatures(prev => [...prev, {
      tempId: crypto.randomUUID(),
      signer_user_id: cleanText(signerForm.signer_user_id),
      site_id: cleanText(signerForm.site_id),
      signer_name: name,
      signer_nrp: cleanText(signerForm.signer_nrp),
      signer_role: normalizeRole(signerForm.signer_role),
      signer_company: cleanText(signerForm.signer_company),
      signatureDataUrl: signerForm.signatureDataUrl,
      existing: false
    }])
    setSignerForm({ signer_user_id:'', site_id:'', signer_name:'', signer_nrp:'', signer_role:'', signer_company:'', signatureDataUrl:'' })
    setSignerSearch('')
    setSignaturePadKey(prev => prev + 1)
    setError('')
  }
  function removeSignature(index){
    setSignatures(prev => prev.filter((_, i) => i !== index))
  }
  async function saveNewSignatures(taskId){
    const newSignatures = signatures.filter(s => !s.existing)
    if (!newSignatures.length) return
    const rows = []
    for (const sig of newSignatures) {
      const file = dataUrlToFile(sig.signatureDataUrl, `signature-${sig.signer_nrp || sig.signer_name || 'petugas'}.png`)
      const signatureUrl = await uploadFoodImage(file, 'signatures')
      rows.push({
        task_id: taskId,
        signer_user_id: sig.signer_user_id || null,
        site_id: sig.site_id || activeTask?.site_id || null,
        signer_name: sig.signer_name,
        signer_nrp: sig.signer_nrp || null,
        signer_role: sig.signer_role || null,
        signer_company: sig.signer_company || null,
        signature_url: signatureUrl,
        signed_at: new Date().toISOString(),
        created_by: profile?.id || null
      })
    }
    const { error } = await supabase.from('food_inspection_signatures').insert(rows)
    if (error) throw error
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
      if (!signatures.length) throw new Error('Minimal 1 petugas/peserta inspeksi wajib tanda tangan sebelum submit.')
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
      await saveNewSignatures(activeTask.id)
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
  const answeredCount = parameters.filter(p => ['0','1'].includes(String(answers[p.id]?.score ?? ''))).length
  const findingCount = parameters.filter(p => String(answers[p.id]?.score ?? '') === '0').length
  return <div className="stack">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title="Tasklist Mingguan Food Index" desc="Task otomatis dibuat berdasarkan vendor catering aktif di site. Klik Start untuk mulai inspeksi." action={<button className="secondary" onClick={load} disabled={loading}>{loading ? 'Memuat...' : 'Refresh / Generate Task'}</button>}>
      {loading ? <p>Memuat task...</p> : <div className="table-wrap" style={{maxHeight:500, overflow:'auto'}}><table><thead><tr><th>Site</th><th>Vendor</th><th>Minggu</th><th>Status</th><th>Mulai</th><th>Submit</th><th>Approval</th><th>Aksi</th></tr></thead><tbody>{tasks.map(t => <tr key={t.id}><td>{t.sites?.site_code || '-'}</td><td>{t.food_vendors?.vendor_name || '-'}</td><td>{t.week_start_date} s/d {t.week_end_date}</td><td>{renderCell(t.status)}</td><td>{t.started_at?.slice(0,16)?.replace('T',' ') || '-'}</td><td>{t.submitted_at?.slice(0,16)?.replace('T',' ') || '-'}</td><td>{t.approved_at?.slice(0,16)?.replace('T',' ') || '-'}</td><td>{canOpen(t) ? <button onClick={()=>openTask(t)}>{t.status === 'Open' ? 'Start Inspeksi' : 'Lanjut / Edit'}</button> : <span className="muted">-</span>}</td></tr>)}</tbody></table>{!tasks.length && <p className="muted table-empty">Belum ada task minggu ini. Pastikan Master Vendor Catering aktif sudah tersedia.</p>}</div>}
    </Panel>
    <Panel title="Export Tasklist" action={<button onClick={()=>downloadXlsx('food-index-tasklist.xlsx', rows)}><Download size={16}/> Export</button>}>
      <Table rows={rows} empty="Belum ada task untuk diexport." />
    </Panel>
    {activeTask && <div className="modal-backdrop food-modal-backdrop"><div className="modal-card wide-modal food-inspection-modal inspection-entry-modal">
      <div className="modal-head food-modal-head">
        <div>
          <span className="modal-eyebrow">Task Mingguan</span>
          <h3>Inspeksi Food Index</h3>
          <p>{activeTask.sites?.site_code} · {activeTask.food_vendors?.vendor_name} · {activeTask.week_start_date} s/d {activeTask.week_end_date}</p>
        </div>
        <button className="icon close-btn" onClick={()=>setActiveTask(null)} aria-label="Tutup modal"><X size={18}/></button>
      </div>
      <div className="inspection-modal-body">
      <div className="inspection-summary">
        <div><small>Parameter terisi</small><b>{answeredCount}/{parameters.length}</b></div>
        <div><small>Temuan</small><b>{findingCount}</b></div>
        <div><small>Ketentuan</small><span>Nilai 0 wajib catatan + foto evidence.</span></div>
      </div>
      <div className="info-note inspection-note"><b>Catatan:</b> Corrective & preventive action diisi setelah submit melalui menu Outstanding, bukan di sesi inspeksi ini.</div>
      <div className="inspection-scroll">
        {parameters.map((p, idx) => {
          const a = answers[p.id] || {}
          const isFinding = a.score === '0'
          const isOk = a.score === '1'
          return <div key={p.id} className={`inspection-item ${isFinding ? 'finding' : ''} ${isOk ? 'passed' : ''}`}>
            <div className="inspection-param">
              <span className="param-index">{idx + 1}</span>
              <div>
                <small>{p.category || 'General'}</small>
                <strong>{p.parameter_text}</strong>
                <div className="standard-meta">
                  {p.hazard_code && <span>Kode Bahaya: {p.hazard_code}</span>}
                  {p.behavior && <span>Behavior: {p.behavior}</span>}
                </div>
                {p.standard_parameter && <div className="standard-box"><b>Standar Parameter:</b> {p.standard_parameter}</div>}
              </div>
            </div>
            <div className="score-options" aria-label="Pilih nilai">
              <button type="button" className={isOk ? 'score-option active good' : 'score-option'} onClick={()=>setAnswer(p.id,{score:'1'})}>
                <b>1</b><span>Sesuai</span>
              </button>
              <button type="button" className={isFinding ? 'score-option active bad' : 'score-option'} onClick={()=>setAnswer(p.id,{score:'0'})}>
                <b>0</b><span>Temuan</span>
              </button>
            </div>
            {isFinding ? <div className="finding-fields">
              <label className="field-block"><span>Catatan Temuan <em>*</em></span><textarea rows={3} placeholder="Jelaskan kondisi temuan secara singkat" value={a.note || ''} onChange={e=>setAnswer(p.id,{note:e.target.value})} /></label>
              <div className="upload-card"><Upload size={20}/><b>{a.evidenceFile?.name || (a.evidenceUrl ? 'Evidence sudah ada' : 'Upload foto evidence')}</b><span>Pilih sumber foto evidence. Wajib untuk nilai 0.</span>
                <div className="camera-options">
                  <label className="camera-option"><Camera size={15}/> Kamera<input type="file" accept="image/*" capture="environment" hidden onChange={e=>setAnswer(p.id,{evidenceFile:e.target.files?.[0] || null})}/></label>
                  <label className="camera-option"><ImageIcon size={15}/> Galeri<input type="file" accept="image/*" hidden onChange={e=>setAnswer(p.id,{evidenceFile:e.target.files?.[0] || null})}/></label>
                </div>
              </div>
              {a.evidenceUrl && <a className="evidence-link" href={a.evidenceUrl} target="_blank" rel="noreferrer">Lihat evidence tersimpan</a>}
            </div> : <div className="no-evidence-note">Pilih <b>0</b> jika ada temuan, lalu isi catatan dan upload foto evidence.</div>}
          </div>
        })}
      <section className="signature-section">
        <div className="signature-section-head">
          <div><h4>Tanda Tangan Petugas / Peserta Inspeksi</h4><p>Minimal 1 tanda tangan wajib sebelum submit. Cari user berdasarkan NRP/nama dari akses Food Index site terkait.</p></div>
          <StatusPill value={`${signatures.length} TTD`} />
        </div>
        <div className="signature-search-box signature-live-search">
          <label>Live Search NRP / Nama User Food Index<input value={signerSearch} onChange={e=>setSignerSearch(e.target.value)} placeholder="Ketik nama, NRP, atau email" /></label>
          <button type="button" className="secondary" onClick={()=>setEligibleSigners(prev=>prev)} disabled>Live</button>
          {!!filteredEligibleSigners.length && <div className="driver-results">
            {filteredEligibleSigners.map(r => {
              const p = r.users_profile || {}
              return <button type="button" key={`${r.user_id}-${r.role}-${r.site_id || 'all'}`} onClick={()=>selectUserSigner(r)}>
                <b>{p.nama || p.email}</b><span>{p.nrp || '-'} · {foodRoleFromAccess(r)} · {r.sites?.site_code || 'All Site'}</span>
              </button>
            })}
          </div>}
          {signerSearch.trim().length >= 2 && !filteredEligibleSigners.length && <div className="signature-empty-hint">User tidak ditemukan. Pastikan user sudah dibuat di Admin Panel Food Index dan site-nya sesuai.</div>}
        </div>
        <div className="signature-form-grid readonly-signature-grid">
          <label>Nama Petugas<input value={signerForm.signer_name} readOnly placeholder="Pilih dari live search" /></label>
          <label>NRP<input value={signerForm.signer_nrp} readOnly placeholder="NRP" /></label>
          <label>Jabatan / Role<input value={signerForm.signer_role} readOnly placeholder="Role" /></label>
          <label>Instansi / Site<input value={signerForm.signer_company} readOnly placeholder="Site" /></label>
        </div>
        {cleanText(signerForm.signer_name) ? <SignaturePad key={`signature-pad-${signaturePadKey}`} value={signerForm.signatureDataUrl} onChange={dataUrl=>setSignerForm(prev=>({...prev, signatureDataUrl:dataUrl}))} /> : <div className="signature-empty-hint">Cari dan pilih user terlebih dahulu. Input manual tanda tangan dinonaktifkan agar role dan site ikut tercatat.</div>}
        <div className="signature-add-row"><button type="button" onClick={addSignature}>+ Tambah Tanda Tangan</button></div>
        <div className="signature-list">
          {signatures.map((sig, idx) => <div className="signature-list-item" key={sig.id || sig.tempId || idx}>
            <div><b>{sig.signer_name}</b><span>{sig.signer_nrp || '-'} · {sig.signer_role || '-'} · {sig.signer_company || '-'}</span></div>
            {sig.signatureUrl ? <a href={sig.signatureUrl} target="_blank" rel="noreferrer">Lihat TTD</a> : <img src={sig.signatureDataUrl} alt={`Tanda tangan ${sig.signer_name}`} />}
            {!sig.existing && <button type="button" className="secondary mini-btn" onClick={()=>removeSignature(idx)}>Hapus</button>}
          </div>)}
          {!signatures.length && <div className="no-evidence-note">Belum ada tanda tangan. Tambahkan minimal 1 petugas sebelum submit.</div>}
        </div>
      </section>
      </div>
      </div>
      <div className="modal-actions sticky-actions"><button className="secondary" onClick={()=>setActiveTask(null)} disabled={submitting}>Batal</button><button onClick={submitInspection} disabled={submitting}>{submitting ? 'Menyimpan...' : 'Submit Inspeksi'}</button></div>
    </div></div>}
  </div>
}
function FoodOutstanding({ context, profile }){
  const [loading, setLoading] = useState(true)
  const [rows, setRows] = useState([])
  const [active, setActive] = useState(null)
  const [form, setForm] = useState({ corrective_action:'', preventive_action:'', due_date:'' })
  const [closeActive, setCloseActive] = useState(null)
  const [closeForm, setCloseForm] = useState({ correctiveFile:null, preventiveFile:null, correctiveUrl:'', preventiveUrl:'' })
  const [msg, setMsg] = useState('')
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)
  const [photoPreview, setPhotoPreview] = useState(null)

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
            due_date,
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
      preventive_action: row.food_findings?.preventive_action || '',
      due_date: row.food_findings?.due_date || ''
    })
    setMsg(''); setError('')
  }

  function openCloseOutstanding(row){
    setCloseActive(row)
    setCloseForm({
      correctiveFile:null,
      preventiveFile:null,
      correctiveUrl: row.corrective_photo_url || '',
      preventiveUrl: row.preventive_photo_url || ''
    })
    setMsg(''); setError('')
  }

  function openOutstandingPhoto(url, title, row){
    if (!url) return
    const task = row?.food_weekly_tasks || {}
    const answer = row?.food_findings?.food_inspection_answers || {}
    setPhotoPreview({
      url,
      title: title || 'Foto Temuan',
      meta: [task?.sites?.site_code, task?.food_vendors?.vendor_name, answer?.food_parameters?.parameter_text].filter(Boolean).join(' · ')
    })
  }

  async function saveCloseOutstanding(){
    if (!closeActive) return
    setSaving(true); setMsg(''); setError('')
    try {
      const hasPlan = cleanText(closeActive.food_findings?.corrective_action) && cleanText(closeActive.food_findings?.preventive_action) && cleanText(closeActive.food_findings?.due_date)
      if (!hasPlan) throw new Error('Corrective, Preventive Action, dan Due Date harus diisi dulu sebelum close outstanding.')
      if (!closeForm.correctiveFile && !closeForm.correctiveUrl) throw new Error('Foto corrective wajib diupload.')
      if (!closeForm.preventiveFile && !closeForm.preventiveUrl) throw new Error('Foto preventive wajib diupload.')
      let correctiveUrl = closeForm.correctiveUrl || ''
      let preventiveUrl = closeForm.preventiveUrl || ''
      if (closeForm.correctiveFile) correctiveUrl = await uploadFoodImage(closeForm.correctiveFile, 'corrective')
      if (closeForm.preventiveFile) preventiveUrl = await uploadFoodImage(closeForm.preventiveFile, 'preventive')
      const now = new Date().toISOString()
      const { error: oe } = await supabase.from('food_outstandings')
        .update({
          corrective_photo_url: correctiveUrl,
          preventive_photo_url: preventiveUrl,
          closed_by: profile?.id,
          closed_at: now,
          status: 'Waiting Close Approval',
          updated_at: now
        })
        .eq('id', closeActive.id)
      if (oe) throw oe
      setCloseActive(null)
      setMsg('Foto corrective & preventive tersimpan. Outstanding menunggu approval Atasan Site.')
      await load()
    } catch(e){ setError(e.message) }
    setSaving(false)
  }

  async function saveActionPlan(){
    if (!active) return
    setSaving(true); setMsg(''); setError('')
    try {
      if (!cleanText(form.corrective_action)) throw new Error('Corrective Action wajib diisi.')
      if (!cleanText(form.preventive_action)) throw new Error('Preventive Action wajib diisi.')
      if (!cleanText(form.due_date)) throw new Error('Due Date wajib diisi.')
      const now = new Date().toISOString()
      const { error: fe } = await supabase.from('food_findings')
        .update({
          corrective_action: cleanText(form.corrective_action),
          preventive_action: cleanText(form.preventive_action),
          due_date: cleanText(form.due_date),
          validated_by: profile?.id,
          validated_at: now,
          status: 'Action Plan Submitted',
          updated_at: now
        })
        .eq('id', active.finding_id)
      if (fe) throw fe

      const taskId = active.task_id
      const { data: findings, error: checkError } = await supabase.from('food_findings')
        .select('id, corrective_action, preventive_action, due_date')
        .eq('task_id', taskId)
      if (checkError) throw checkError
      const allReady = (findings || []).length > 0 && (findings || []).every(f => cleanText(f.corrective_action) && cleanText(f.preventive_action) && cleanText(f.due_date))
      if (allReady) {
        const { error: te } = await supabase.from('food_weekly_tasks')
          .update({ status:'Waiting Approval', updated_at:now })
          .eq('id', taskId)
        if (te) throw te
      }
      setActive(null)
      setMsg(allReady ? 'Corrective, preventive, dan due date tersimpan. Inspeksi masuk Waiting Approval Atasan Site.' : 'Corrective, preventive, dan due date tersimpan.')
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
    due_date: r.food_findings?.due_date || '-',
    status_task: r.food_weekly_tasks?.status || '-',
    status_outstanding: r.status || '-'
  }))
  const outstandingExportColumns = ['site','vendor','minggu','parameter','catatan_temuan','corrective_action','preventive_action','due_date','status_task','status_outstanding']

  return <div className="stack">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title="Outstanding / Temuan Food Index" desc="Isi corrective action, preventive action, dan due date dari temuan inspeksi. Setelah semua temuan pada task terisi, inspeksi masuk approval Atasan Site." action={<button className="secondary" onClick={load} disabled={loading}>{loading ? 'Memuat...' : 'Refresh'}</button>}>
      {loading ? <p>Memuat outstanding...</p> : <div className="table-wrap outstanding-table-wrap" style={{maxHeight:520, overflow:'auto'}}><table><thead><tr><th>Site</th><th>Vendor</th><th>Minggu</th><th>Parameter</th><th>Catatan</th><th>Evidence</th><th>Corrective</th><th>Preventive</th><th>Due Date</th><th>Status Task</th><th>Aksi</th></tr></thead><tbody>{rows.map(r => {
        const answer = r.food_findings?.food_inspection_answers
        const hasPlan = cleanText(r.food_findings?.corrective_action) && cleanText(r.food_findings?.preventive_action) && cleanText(r.food_findings?.due_date)
        return <tr key={r.id}><td>{r.food_weekly_tasks?.sites?.site_code || '-'}</td><td>{r.food_weekly_tasks?.food_vendors?.vendor_name || '-'}</td><td>{r.food_weekly_tasks?.week_start_date || '-'}</td><td>{answer?.food_parameters?.parameter_text || '-'}</td><td>{answer?.finding_note || '-'}</td><td>{answer?.evidence_photo_url ? <button type="button" className="photo-view-btn" onClick={()=>openOutstandingPhoto(answer.evidence_photo_url, 'Foto Evidence', r)}><ImageIcon size={14}/> View</button> : '-'}</td><td>{r.food_findings?.corrective_action || '-'}</td><td>{r.food_findings?.preventive_action || '-'}</td><td>{r.food_findings?.due_date || '-'}</td><td className="status-task-cell"><TaskStatusPill value={r.food_weekly_tasks?.status} /></td><td>{!hasPlan ? <button onClick={()=>openActionPlan(r)}>Isi Tindakan</button> : r.status === 'Closed' ? <StatusPill value="Closed" /> : r.status === 'Waiting Close Approval' ? <StatusPill value="Waiting Close Approval" /> : <button onClick={()=>openCloseOutstanding(r)}>Close Outstanding</button>}</td></tr>
      })}</tbody></table>{!rows.length && <p className="muted table-empty">Belum ada outstanding / temuan.</p>}</div>}
    </Panel>
    <Panel title="Export Outstanding" action={<button onClick={()=>downloadXlsx('food-index-outstanding.xlsx', tableRows)}><Download size={16}/> Export</button>}>
      <div className="outstanding-export-table"><Table rows={tableRows} columns={outstandingExportColumns} empty="Belum ada data outstanding." /></div>
    </Panel>
    {active && <div className="modal-backdrop"><div className="modal-card">
      <div className="modal-head"><div><h3>Isi Action Plan Temuan</h3><p>{active.food_weekly_tasks?.sites?.site_code} · {active.food_weekly_tasks?.food_vendors?.vendor_name}</p></div><button className="icon" onClick={()=>setActive(null)}><X size={18}/></button></div>
      <div className="stack compact">
        <label>Catatan Temuan<input value={active.food_findings?.food_inspection_answers?.finding_note || '-'} disabled /></label>
        <label>Corrective Action<textarea rows={4} value={form.corrective_action} onChange={e=>setForm({...form, corrective_action:e.target.value})} placeholder="Isi tindakan corrective" /></label>
        <label>Preventive Action<textarea rows={4} value={form.preventive_action} onChange={e=>setForm({...form, preventive_action:e.target.value})} placeholder="Isi tindakan preventive" /></label>
        <label>Due Date<input type="date" value={form.due_date} onChange={e=>setForm({...form, due_date:e.target.value})} /></label>
      </div>
      <div className="modal-actions"><button className="secondary" onClick={()=>setActive(null)} disabled={saving}>Batal</button><button onClick={saveActionPlan} disabled={saving}>{saving ? 'Menyimpan...' : 'Simpan Action Plan'}</button></div>
    </div></div>}
    {closeActive && <div className="modal-backdrop"><div className="modal-card">
      <div className="modal-head"><div><h3>Close Outstanding</h3><p>{closeActive.food_weekly_tasks?.sites?.site_code} · {closeActive.food_weekly_tasks?.food_vendors?.vendor_name}</p></div><button className="icon" onClick={()=>setCloseActive(null)}><X size={18}/></button></div>
      <div className="stack compact">
        <div className="info-note"><b>Action Plan:</b><br/>Corrective: {closeActive.food_findings?.corrective_action || '-'}<br/>Preventive: {closeActive.food_findings?.preventive_action || '-'}<br/>Due Date: {closeActive.food_findings?.due_date || '-'}</div>
        <label className="upload-line"><Upload size={18}/><span>{closeForm.correctiveFile?.name || (closeForm.correctiveUrl ? 'Foto corrective sudah ada' : 'Upload Foto Corrective')}</span><input type="file" accept="image/*" hidden onChange={e=>setCloseForm({...closeForm, correctiveFile:e.target.files?.[0] || null})}/></label>
        {closeForm.correctiveUrl && <button type="button" className="photo-view-btn inline-photo-view-btn" onClick={()=>openOutstandingPhoto(closeForm.correctiveUrl, 'Foto Corrective Tersimpan', closeActive)}><ImageIcon size={14}/> View foto corrective tersimpan</button>}
        <label className="upload-line"><Upload size={18}/><span>{closeForm.preventiveFile?.name || (closeForm.preventiveUrl ? 'Foto preventive sudah ada' : 'Upload Foto Preventive')}</span><input type="file" accept="image/*" hidden onChange={e=>setCloseForm({...closeForm, preventiveFile:e.target.files?.[0] || null})}/></label>
        {closeForm.preventiveUrl && <button type="button" className="photo-view-btn inline-photo-view-btn" onClick={()=>openOutstandingPhoto(closeForm.preventiveUrl, 'Foto Preventive Tersimpan', closeActive)}><ImageIcon size={14}/> View foto preventive tersimpan</button>}
      </div>
      <div className="modal-actions"><button className="secondary" onClick={()=>setCloseActive(null)} disabled={saving}>Batal</button><button onClick={saveCloseOutstanding} disabled={saving}>{saving ? 'Menyimpan...' : 'Submit Close ke Approval'}</button></div>
    </div></div>}
    {photoPreview && <FoodPhotoModal photo={photoPreview} onClose={()=>setPhotoPreview(null)} />}
  </div>
}
function FoodApproval({ context, profile }){
  const [loading, setLoading] = useState(true)
  const [tasks, setTasks] = useState([])
  const [closeRows, setCloseRows] = useState([])
  const [active, setActive] = useState(null)
  const [rejectNote, setRejectNote] = useState('')
  const [msg, setMsg] = useState('')
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)

  useEffect(() => { load() }, [context?.id])

  async function load(){
    setLoading(true); setError('')
    try {
      let q = supabase.from('food_weekly_tasks')
        .select(`
          *,
          sites(site_code, site_name),
          food_vendors(vendor_name),
          food_inspection_signatures(*),
          food_inspection_answers(
            id,
            score,
            finding_note,
            evidence_photo_url,
            food_parameters(category, parameter_text, sort_order)
          ),
          food_findings(
            id,
            corrective_action,
            preventive_action,
            due_date,
            status,
            validated_at,
            food_inspection_answers(
              finding_note,
              evidence_photo_url,
              food_parameters(category, parameter_text, sort_order)
            )
          )
        `)
        .eq('status', 'Waiting Approval')
        .order('submitted_at', { ascending:true })
      if (!adminCanSeeAll(context)) q = q.eq('site_id', context.site_id)
      const { data, error } = await q
      if (error) throw error
      setTasks(data || [])

      let closeQ = supabase.from('food_outstandings')
        .select(`
          *,
          food_findings(
            corrective_action,
            preventive_action,
            due_date,
            food_inspection_answers(
              finding_note,
              evidence_photo_url,
              food_parameters(category, parameter_text)
            )
          ),
          food_weekly_tasks!inner(
            id,
            week_start_date,
            site_id,
            vendor_id,
            sites(site_code, site_name),
            food_vendors(vendor_name)
          )
        `)
        .eq('status', 'Waiting Close Approval')
        .order('closed_at', { ascending:true })
      if (!adminCanSeeAll(context)) closeQ = closeQ.eq('food_weekly_tasks.site_id', context.site_id)
      const { data: closeData, error: closeError } = await closeQ
      if (closeError) throw closeError
      setCloseRows(closeData || [])
    } catch(e){ setError(e.message) }
    setLoading(false)
  }

  function openReview(task){
    setActive(task)
    setRejectNote('')
    setMsg('')
    setError('')
  }

  function getTaskFindings(task){
    return (task?.food_findings || []).filter(f => cleanText(f.corrective_action) && cleanText(f.preventive_action) && cleanText(f.due_date))
  }

  function isTaskReady(task){
    const answers = task?.food_inspection_answers || []
    const findingAnswers = answers.filter(a => Number(a.score) === 0)
    if (!findingAnswers.length) return true
    const readyFindings = getTaskFindings(task)
    return readyFindings.length >= findingAnswers.length
  }

  async function approveTask(task){
    if (!task) return
    setSaving(true); setMsg(''); setError('')
    try {
      if (!isTaskReady(task)) throw new Error('Masih ada temuan yang belum punya corrective, preventive action, dan due date.')
      const now = new Date().toISOString()
      const { error: te } = await supabase.from('food_weekly_tasks')
        .update({ status:'Approved', approved_by:profile?.id, approved_at:now, updated_at:now })
        .eq('id', task.id)
      if (te) throw te
      const { error: fe } = await supabase.from('food_findings')
        .update({ status:'Approved', updated_at:now })
        .eq('task_id', task.id)
      if (fe) throw fe
      setActive(null)
      setMsg('Inspeksi Food Index berhasil di-approve. Achievement site sudah terhitung.')
      await load()
    } catch(e){ setError(e.message) }
    setSaving(false)
  }

  async function rejectTask(task){
    if (!task) return
    setSaving(true); setMsg(''); setError('')
    try {
      if (!cleanText(rejectNote)) throw new Error('Catatan reject wajib diisi.')
      const now = new Date().toISOString()
      const { error: te } = await supabase.from('food_weekly_tasks')
        .update({ status:'Rejected', rejected_by:profile?.id, rejected_at:now, rejection_note:cleanText(rejectNote), updated_at:now })
        .eq('id', task.id)
      if (te) throw te
      setActive(null)
      setMsg('Inspeksi ditolak. GL dapat memperbaiki dan submit ulang dari Tasklist Inspeksi.')
      await load()
    } catch(e){ setError(e.message) }
    setSaving(false)
  }

  async function approveCloseOutstanding(row){
    if (!row) return
    setSaving(true); setMsg(''); setError('')
    try {
      const now = new Date().toISOString()
      const { error } = await supabase.from('food_outstandings')
        .update({ status:'Closed', approved_by:profile?.id, approved_at:now, updated_at:now })
        .eq('id', row.id)
      if (error) throw error
      setMsg('Close outstanding berhasil di-approve.')
      await load()
    } catch(e){ setError(e.message) }
    setSaving(false)
  }

  async function rejectCloseOutstanding(row){
    if (!row) return
    const note = window.prompt('Catatan reject close outstanding:')
    if (!cleanText(note)) return
    setSaving(true); setMsg(''); setError('')
    try {
      const now = new Date().toISOString()
      const { error } = await supabase.from('food_outstandings')
        .update({ status:'Rejected', rejected_by:profile?.id, rejected_at:now, rejection_note:cleanText(note), updated_at:now })
        .eq('id', row.id)
      if (error) throw error
      setMsg('Close outstanding ditolak. GL dapat upload ulang foto close dari menu Outstanding.')
      await load()
    } catch(e){ setError(e.message) }
    setSaving(false)
  }

  const tableRows = tasks.map(t => ({
    id: t.id,
    site: t.sites?.site_code || '-',
    vendor: t.food_vendors?.vendor_name || '-',
    minggu: `${t.week_start_date} s/d ${t.week_end_date}`,
    status: t.status,
    submitted_at: t.submitted_at?.slice(0,16)?.replace('T',' ') || '-',
    jumlah_temuan: (t.food_inspection_answers || []).filter(a => Number(a.score) === 0).length,
    action_plan_ready: isTaskReady(t) ? 'Siap Approval' : 'Belum Lengkap'
  }))

  const answers = active?.food_inspection_answers || []
  const sortedAnswers = [...answers].sort((a,b)=>(a.food_parameters?.sort_order || 0) - (b.food_parameters?.sort_order || 0))
  const findingCount = answers.filter(a => Number(a.score) === 0).length

  return <div className="stack">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title="Approval Inspeksi Food Index" desc="Approve hanya untuk inspeksi yang sudah submit. Jika ada temuan, corrective, preventive, dan due date harus sudah diisi GL dari menu Outstanding." action={<button className="secondary" onClick={load} disabled={loading}>{loading ? 'Memuat...' : 'Refresh'}</button>}>
      {loading ? <p>Memuat approval...</p> : <div className="table-wrap" style={{maxHeight:520, overflow:'auto'}}><table><thead><tr><th>Site</th><th>Vendor</th><th>Minggu</th><th>Submit</th><th>Temuan</th><th>Action Plan</th><th>Aksi</th></tr></thead><tbody>{tasks.map(t => <tr key={t.id}><td>{t.sites?.site_code || '-'}</td><td>{t.food_vendors?.vendor_name || '-'}</td><td>{t.week_start_date} s/d {t.week_end_date}</td><td>{t.submitted_at?.slice(0,16)?.replace('T',' ') || '-'}</td><td>{(t.food_inspection_answers || []).filter(a => Number(a.score) === 0).length}</td><td>{isTaskReady(t) ? <StatusPill value="Approved" /> : <StatusPill value="Need Action Plan" />}</td><td><button onClick={()=>openReview(t)}>Review</button></td></tr>)}</tbody></table>{!tasks.length && <p className="muted table-empty">Belum ada inspeksi yang menunggu approval.</p>}</div>}
    </Panel>
    <Panel title="Approval Close Outstanding" desc="Approve close outstanding setelah GL upload foto corrective dan preventive." action={<button className="secondary" onClick={load} disabled={loading}>{loading ? 'Memuat...' : 'Refresh'}</button>}>
      {loading ? <p>Memuat approval close outstanding...</p> : <div className="table-wrap" style={{maxHeight:420, overflow:'auto'}}><table><thead><tr><th>Site</th><th>Vendor</th><th>Minggu</th><th>Parameter</th><th>Corrective</th><th>Preventive</th><th>Due Date</th><th>Foto Corrective</th><th>Foto Preventive</th><th>Aksi</th></tr></thead><tbody>{closeRows.map(r => {
        const ans = r.food_findings?.food_inspection_answers
        return <tr key={r.id}><td>{r.food_weekly_tasks?.sites?.site_code || '-'}</td><td>{r.food_weekly_tasks?.food_vendors?.vendor_name || '-'}</td><td>{r.food_weekly_tasks?.week_start_date || '-'}</td><td>{ans?.food_parameters?.parameter_text || '-'}</td><td>{r.food_findings?.corrective_action || '-'}</td><td>{r.food_findings?.preventive_action || '-'}</td><td>{r.food_findings?.due_date || '-'}</td><td>{r.corrective_photo_url ? <a href={r.corrective_photo_url} target="_blank" rel="noreferrer">Lihat foto</a> : '-'}</td><td>{r.preventive_photo_url ? <a href={r.preventive_photo_url} target="_blank" rel="noreferrer">Lihat foto</a> : '-'}</td><td className="action-cell"><button className="danger" disabled={saving} onClick={()=>rejectCloseOutstanding(r)}>Reject</button><button disabled={saving} onClick={()=>approveCloseOutstanding(r)}>Approve Close</button></td></tr>
      })}</tbody></table>{!closeRows.length && <p className="muted table-empty">Belum ada close outstanding yang menunggu approval.</p>}</div>}
    </Panel>
    <Panel title="Export Approval Queue" action={<button onClick={()=>downloadXlsx('food-index-approval-queue.xlsx', tableRows)}><Download size={16}/> Export</button>}>
      <Table rows={tableRows} empty="Tidak ada data approval." />
    </Panel>
    {active && <div className="modal-backdrop food-modal-backdrop"><div className="modal-card wide-modal food-inspection-modal">
      <div className="modal-head food-modal-head">
        <div>
          <span className="modal-eyebrow">Approval Atasan Site</span>
          <h3>Review Inspeksi Food Index</h3>
          <p>{active.sites?.site_code} · {active.food_vendors?.vendor_name} · {active.week_start_date} s/d {active.week_end_date}</p>
        </div>
        <button className="icon close-btn" onClick={()=>setActive(null)} aria-label="Tutup modal"><X size={18}/></button>
      </div>
      <div className="inspection-summary">
        <div><small>Total Parameter</small><b>{answers.length}</b></div>
        <div><small>Total Temuan</small><b>{findingCount}</b></div>
        <div><small>Status</small><span>{isTaskReady(active) ? 'Siap di-approve' : 'Belum siap, action plan belum lengkap'}</span></div>
      </div>
      <div className="inspection-scroll">
        {sortedAnswers.map((a, idx) => {
          const finding = (active.food_findings || []).find(f => f.answer_id === a.id)
          const isFinding = Number(a.score) === 0
          return <div key={a.id} className={`inspection-item ${isFinding ? 'finding approval-review-finding' : 'passed'}`}>
            <div className="inspection-param">
              <span className="param-index">{idx + 1}</span>
              <div><small>{a.food_parameters?.category || 'General'}</small><strong>{a.food_parameters?.parameter_text || '-'}</strong></div>
            </div>
            <div className="approval-detail-grid">
              <div><small>Nilai</small><b>{Number(a.score) === 1 ? '1 - Sesuai' : '0 - Temuan'}</b></div>
              <div><small>Catatan Temuan</small><span>{a.finding_note || '-'}</span></div>
              <div><small>Evidence</small>{a.evidence_photo_url ? <a href={a.evidence_photo_url} target="_blank" rel="noreferrer">Lihat foto</a> : <span>-</span>}</div>
              <div><small>Corrective Action</small><span>{finding?.corrective_action || '-'}</span></div>
              <div><small>Preventive Action</small><span>{finding?.preventive_action || '-'}</span></div>
              <div><small>Due Date</small><span>{finding?.due_date || '-'}</span></div>
            </div>
          </div>
        })}
        <section className="approval-signature-box">
          <div className="signature-section-head"><div><h4>Tanda Tangan Petugas / Peserta Inspeksi</h4><p>Daftar petugas yang menandatangani hasil inspeksi sebelum submit.</p></div><StatusPill value={`${(active.food_inspection_signatures || []).length} TTD`} /></div>
          <div className="signature-list">
            {(active.food_inspection_signatures || []).map(sig => <div className="signature-list-item" key={sig.id}>
              <div><b>{sig.signer_name}</b><span>{sig.signer_nrp || '-'} · {sig.signer_role || '-'} · {sig.signer_company || '-'}</span></div>
              {sig.signature_url ? <a href={sig.signature_url} target="_blank" rel="noreferrer">Lihat TTD</a> : <span>-</span>}
            </div>)}
            {!(active.food_inspection_signatures || []).length && <div className="no-evidence-note">Belum ada tanda tangan tersimpan pada inspeksi ini.</div>}
          </div>
        </section>
        <label className="field-block"><span>Catatan Reject</span><textarea rows={3} placeholder="Isi jika approval ditolak" value={rejectNote} onChange={e=>setRejectNote(e.target.value)} /></label>
      </div>
      <div className="modal-actions sticky-actions">
        <button className="secondary" onClick={()=>setActive(null)} disabled={saving}>Tutup</button>
        <button className="danger" onClick={()=>rejectTask(active)} disabled={saving}>{saving ? 'Menyimpan...' : 'Reject'}</button>
        <button onClick={()=>approveTask(active)} disabled={saving || !isTaskReady(active)}>{saving ? 'Menyimpan...' : 'Approve Inspeksi'}</button>
      </div>
    </div></div>}
  </div>
}
function FoodReport({ context }){
  const [loading, setLoading] = useState(true)
  const [rows, setRows] = useState([])
  const [error, setError] = useState('')
  const [photoPreview, setPhotoPreview] = useState(null)

  useEffect(() => { load() }, [context?.id])

  async function load(){
    setLoading(true); setError('')
    try {
      let q = supabase.from('food_outstandings')
        .select(`
          id,
          status,
          corrective_photo_url,
          preventive_photo_url,
          created_at,
          food_findings(
            id,
            corrective_action,
            preventive_action,
            due_date,
            food_inspection_answers(
              finding_note,
              evidence_photo_url,
              food_parameters(category, parameter_text)
            )
          ),
          food_weekly_tasks!inner(
            id,
            week_start_date,
            submitted_at,
            status,
            site_id,
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

  const reportRows = rows.map((r, idx) => {
    const answer = r.food_findings?.food_inspection_answers
    const task = r.food_weekly_tasks
    return {
      Nomor: idx + 1,
      Site: task?.sites?.site_code || '-',
      Vendor: task?.food_vendors?.vendor_name || '-',
      'Parameter checklist': answer?.food_parameters?.parameter_text || '-',
      'Caption Temuannya': answer?.finding_note || '-',
      'Tgl inspeksi': task?.submitted_at ? task.submitted_at.slice(0, 10) : (task?.week_start_date || '-'),
      'Corrective Action': r.food_findings?.corrective_action || '-',
      'Preventive Action': r.food_findings?.preventive_action || '-',
      'Due Date': r.food_findings?.due_date || '-',
      'Foto Corrective': r.corrective_photo_url || '-',
      'Foto Preventive': r.preventive_photo_url || '-',
      Status: r.status === 'Closed' ? 'Close' : 'Open'
    }
  })

  const totalOpen = rows.filter(r => r.status !== 'Closed').length
  const totalClosed = rows.filter(r => r.status === 'Closed').length
  const dueSoon = rows.filter(r => r.status !== 'Closed' && r.food_findings?.due_date && r.food_findings.due_date <= today()).length

  function openReportPhoto(url, title, r){
    if (!url) return
    const task = r?.food_weekly_tasks || {}
    const answer = r?.food_findings?.food_inspection_answers || {}
    setPhotoPreview({
      url,
      title: title || 'Foto Temuan',
      meta: [task?.sites?.site_code, task?.food_vendors?.vendor_name, answer?.food_parameters?.parameter_text].filter(Boolean).join(' · ')
    })
  }

  return <div className="stack">
    {error && <div className="error">{error}</div>}
    <div className="kpi-grid three">
      <Kpi title="Total Temuan" value={rows.length} icon={<AlertTriangle size={22}/>} />
      <Kpi title="Open" value={totalOpen} icon={<CalendarCheck size={22}/>} />
      <Kpi title="Closed" value={totalClosed} icon={<CheckCircle2 size={22}/>} />
      <Kpi title="Due / Overdue" value={dueSoon} icon={<ShieldCheck size={22}/>} />
    </div>
    <Panel title="Report Temuan Food Index" desc="Khusus item yang memiliki temuan. Data ini bisa dilihat langsung dan diexport ke Excel." action={<div className="action-cell"><button className="secondary" onClick={load} disabled={loading}>{loading ? 'Memuat...' : 'Refresh'}</button><button onClick={()=>downloadXlsx('food-index-report-temuan.xlsx', reportRows)}><Download size={16}/> Export Excel</button></div>}>
      {loading ? <p>Memuat report...</p> : <div>
        <div className="table-wrap report-table-wrap" style={{maxHeight:620, overflow:'auto'}}>
          <table>
            <thead><tr><th>Nomor</th><th>Site</th><th>Vendor</th><th>Parameter checklist</th><th>Caption Temuannya</th><th>Tgl inspeksi</th><th>Corrective Action</th><th>Preventive Action</th><th>Due Date</th><th>Foto Corrective</th><th>Foto Preventive</th><th>Status</th></tr></thead>
            <tbody>{rows.map((r, idx) => {
              const answer = r.food_findings?.food_inspection_answers
              const task = r.food_weekly_tasks
              return <tr key={r.id}>
                <td>{idx + 1}</td>
                <td>{task?.sites?.site_code || '-'}</td>
                <td>{task?.food_vendors?.vendor_name || '-'}</td>
                <td>{answer?.food_parameters?.parameter_text || '-'}</td>
                <td>{answer?.finding_note || '-'}</td>
                <td>{task?.submitted_at ? task.submitted_at.slice(0, 10) : (task?.week_start_date || '-')}</td>
                <td>{r.food_findings?.corrective_action || '-'}</td>
                <td>{r.food_findings?.preventive_action || '-'}</td>
                <td>{r.food_findings?.due_date || '-'}</td>
                <td>{r.corrective_photo_url ? <button type="button" className="photo-view-btn" onClick={()=>openReportPhoto(r.corrective_photo_url, 'Foto Corrective', r)}><ImageIcon size={14}/> View</button> : '-'}</td>
                <td>{r.preventive_photo_url ? <button type="button" className="photo-view-btn" onClick={()=>openReportPhoto(r.preventive_photo_url, 'Foto Preventive', r)}><ImageIcon size={14}/> View</button> : '-'}</td>
                <td><StatusPill value={r.status === 'Closed' ? 'Close' : 'Open'} /></td>
              </tr>
            })}</tbody>
          </table>
          {!rows.length && <p className="muted table-empty">Belum ada temuan Food Index.</p>}
        </div>
      </div>}
    </Panel>
    {photoPreview && <FoodPhotoModal photo={photoPreview} onClose={()=>setPhotoPreview(null)} />}
  </div>
}

function FoodPhotoModal({ photo, onClose }){
  const [failed, setFailed] = useState(false)
  if (!photo?.url) return null
  return <div className="modal-backdrop food-modal-backdrop photo-view-backdrop" role="dialog" aria-modal="true" onClick={onClose}>
    <div className="modal-card photo-view-modal" onClick={e=>e.stopPropagation()}>
      <div className="modal-head photo-view-head">
        <div>
          <h3>{photo.title || 'Preview Foto'}</h3>
          <p>{photo.meta || 'Preview foto dari report Food Index'}</p>
        </div>
        <button className="icon" onClick={onClose} aria-label="Tutup preview foto"><X size={18}/></button>
      </div>
      <div className="photo-view-body">
        {failed ? <div className="photo-view-failed"><AlertTriangle size={22}/><b>Foto gagal dimuat</b><span>Pastikan file masih tersedia di Supabase Storage.</span></div> : <img src={photo.url} alt={photo.title || 'Foto Food Index'} onError={()=>setFailed(true)} />}
      </div>
      <div className="photo-view-actions"><button className="secondary" onClick={onClose}>Tutup</button></div>
    </div>
  </div>
}
function Placeholder({ title, desc }){
  return <Panel title={title} desc={desc}><div className="card"><ClipboardCheck size={42}/><h3>{title} sedang disiapkan</h3><p>{desc}</p></div></Panel>
}
