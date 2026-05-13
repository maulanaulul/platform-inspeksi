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

function FoodIndexScopedStyles(){
  return <style>{`
    .food-index-app, .food-index-app * { font-family: Arial, Helvetica, sans-serif !important; }
    .food-index-app h2, .food-index-app h3, .food-index-app b, .food-index-app strong { letter-spacing: -0.02em; }
    .food-index-app h2 { font-weight: 700; }
    .food-index-app h3 { font-weight: 700; }
    .food-index-app button { font-weight: 600; }
    .food-index-app input, .food-index-app select, .food-index-app textarea { font-family: Arial, Helvetica, sans-serif !important; font-weight: 500; }
    .food-index-app .panel, .food-index-app .kpi, .food-index-app .modal-card { border-radius: 22px; }
    .food-index-app .table-wrap { border-radius: 18px; }
    .food-index-app .food-modal-backdrop { align-items: center; padding: 22px; }
    .food-index-app .food-inspection-modal { width: min(1060px, calc(100vw - 34px)); max-height: calc(100vh - 42px); display: flex; flex-direction: column; overflow: hidden; padding: 0; }
    .food-index-app .food-modal-head { padding: 26px 30px 18px; border-bottom: 1px solid #e8eef8; }
    .food-index-app .food-modal-head h3 { margin: 4px 0 6px; font-size: 30px; line-height: 1.08; }
    .food-index-app .food-modal-head p { margin: 0; color: #667694; font-weight: 500; }
    .food-index-app .modal-eyebrow { color: #2563eb; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .08em; }
    .food-index-app .close-btn { width: 46px; height: 46px; border-radius: 14px; }
    .food-index-app .inspection-summary { display: grid; grid-template-columns: 1fr 1fr 2.2fr; gap: 12px; padding: 16px 30px 0; }
    .food-index-app .inspection-summary > div { background: #f8fbff; border: 1px solid #dfe9f8; border-radius: 16px; padding: 12px 14px; min-height: 58px; }
    .food-index-app .inspection-summary small { display:block; color:#64748b; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing:.04em; }
    .food-index-app .inspection-summary b { color:#0f172a; font-size: 22px; }
    .food-index-app .inspection-summary span { display:block; margin-top:4px; color:#334155; font-weight: 600; }
    .food-index-app .inspection-note { margin: 14px 30px 0; border-radius: 14px; line-height: 1.45; }
    .food-index-app .inspection-scroll { flex: 1; overflow: auto; padding: 16px 30px 22px; display: grid; gap: 14px; }
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
    .food-index-app .table-wrap { background:#fff; border: 1px solid #dbe7f7; box-shadow: inset 0 1px 0 rgba(255,255,255,.8); }
    .food-index-app table th { background: linear-gradient(180deg, #f8fbff, #eef6ff); color:#334155; font-size: 12px; letter-spacing:.08em; }
    .food-index-app table td { vertical-align: top; }
    .food-index-app .status-pill { border-radius: 999px; padding: 7px 11px; font-weight: 700; }
    .food-index-app .status-pill.ok { background:#dcfce7; color:#166534; }
    .food-index-app .status-pill.warn { background:#fef3c7; color:#92400e; }
    .food-index-app .status-pill.bad { background:#fee2e2; color:#991b1b; }
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


  `}</style>
}

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
  const weekEnd = endOfWeekSunday(monday)
  const thursdayAlert = new Date().getDay() >= 4 || new Date().getDay() === 0

  useEffect(() => { load() }, [context?.id, siteFilter])
  async function load(){
    setLoading(true); setError('')
    try {
      if (adminCanSeeAll(context)) setSites(await fetchSites())
      let tq = supabase
        .from('food_weekly_tasks')
        .select('*, sites(site_code,site_name), food_vendors(vendor_name)')
        .eq('week_start_date', monday)
        .order('generated_at', { ascending:false })
      let oq = supabase
        .from('food_outstandings')
        .select('*, food_findings(corrective_action, preventive_action, due_date, food_inspection_answers(finding_note, evidence_photo_url, food_parameters(parameter_text)), food_weekly_tasks(week_start_date, submitted_at, approved_at, sites(site_code,site_name), food_vendors(vendor_name)))')
        .order('created_at', { ascending:false })
        .limit(500)
      tq = applySiteFilter(tq, context)
      if (!adminCanSeeAll(context)) oq = oq.eq('task_id', '00000000-0000-0000-0000-000000000000')
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
  const alertRows = tasks.filter(t => thursdayAlert && ['Open','In Progress','Need Action Plan'].includes(t.status))
  const reportRows = outs.slice(0, 8).map((o, idx) => {
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

  return <div className="food-dashboard-v43">
    {error && <div className="error">{error}</div>}
    <div className="dashboard-filter-card">
      <div>
        <small>Scope dashboard</small>
        <h3>{siteFilter ? sites.find(s => s.id === siteFilter)?.site_code || 'Site terpilih' : 'Semua site'}</h3>
        <p>Dashboard mengikuti minggu berjalan. Achievement hanya bertambah setelah inspeksi full approved.</p>
      </div>
      <div className="dashboard-filter-actions">
        {adminCanSeeAll(context) && <label>Filter Site<select value={siteFilter} onChange={e=>setSiteFilter(e.target.value)}><option value="">Semua Site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>}
        <button onClick={load}>⟳ Refresh</button>
      </div>
      <div className="dashboard-tags"><span>Site: {siteFilter ? sites.find(s => s.id === siteFilter)?.site_code || 'Terpilih' : 'Semua site'}</span><span>Minggu: {monday} s/d {weekEnd}</span>{thursdayAlert && <span>Alert Kamis aktif</span>}</div>
    </div>

    <div className="dashboard-kpi-grid">
      <div className="dash-kpi"><small>Task Minggu Ini</small><strong>{total}</strong><span>↗ {sites.length || 1} site aktif</span></div>
      <div className="dash-kpi"><small>Approved</small><strong>{approved}</strong><span>Achievement {achievement}%</span></div>
      <div className="dash-kpi"><small>Outstanding Open</small><strong>{outs.filter(o => o.status !== 'Closed').length}</strong><span className="orange">Butuh follow up</span></div>
      <div className="dash-kpi"><small>Expired</small><strong>{expired}</strong><span className="red">Impact achievement</span></div>
    </div>

    <div className="dashboard-main-grid">
      <section className="dashboard-card report-preview-card">
        <div className="dashboard-card-head">
          <div><h3>Report Temuan Food Index</h3><p>Hanya menampilkan parameter yang memiliki temuan. Siap untuk view dan export Excel.</p></div>
          <button onClick={()=>downloadXlsx('food-index-report-temuan-preview.xlsx', reportRows)}><Download size={16}/> Export Excel</button>
        </div>
        <div className="report-filter-row"><input placeholder="Search row data..." readOnly /><select value={siteFilter} onChange={e=>setSiteFilter(e.target.value)} disabled={!adminCanSeeAll(context)}><option value="">Semua Site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code}</option>)}</select><select disabled><option>Semua Status</option></select></div>
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

    <section className="dashboard-card task-row-card">
      <div className="dashboard-card-head"><div><h3>Row Data Task Minggu Ini</h3><p>Task otomatis berdasarkan vendor catering aktif per site.</p></div><button onClick={()=>downloadXlsx('food-index-task-minggu-ini.xlsx', tasks.map(t=>({ site:t.sites?.site_code, vendor:t.food_vendors?.vendor_name, minggu:t.week_start_date, status:t.status, mulai:t.started_at?.slice(0,10), submit:t.submitted_at?.slice(0,10), approved:t.approved_at?.slice(0,10) })))}><Download size={16}/> Export</button></div>
      {loading ? <p>Memuat...</p> : <Table rows={tasks.map(t => ({ site:t.sites?.site_code, vendor:t.food_vendors?.vendor_name, minggu:t.week_start_date, status:t.status, mulai:t.started_at?.slice(0,10), submit:t.submitted_at?.slice(0,10), approved:t.approved_at?.slice(0,10) }))} />}
    </section>
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
    {activeTask && <div className="modal-backdrop food-modal-backdrop"><div className="modal-card wide-modal food-inspection-modal">
      <div className="modal-head food-modal-head">
        <div>
          <span className="modal-eyebrow">Task Mingguan</span>
          <h3>Inspeksi Food Index</h3>
          <p>{activeTask.sites?.site_code} · {activeTask.food_vendors?.vendor_name} · {activeTask.week_start_date} s/d {activeTask.week_end_date}</p>
        </div>
        <button className="icon close-btn" onClick={()=>setActiveTask(null)} aria-label="Tutup modal"><X size={18}/></button>
      </div>
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
              <label className="upload-card"><Upload size={20}/><b>{a.evidenceFile?.name || (a.evidenceUrl ? 'Evidence sudah ada' : 'Upload foto evidence')}</b><span>Format JPG/PNG/WEBP. Wajib untuk nilai 0.</span><input type="file" accept="image/*" hidden onChange={e=>setAnswer(p.id,{evidenceFile:e.target.files?.[0] || null})}/></label>
              {a.evidenceUrl && <a className="evidence-link" href={a.evidenceUrl} target="_blank" rel="noreferrer">Lihat evidence tersimpan</a>}
            </div> : <div className="no-evidence-note">Pilih <b>0</b> jika ada temuan, lalu isi catatan dan upload foto evidence.</div>}
          </div>
        })}
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

  return <div className="stack">
    {msg && <div className="success">{msg}</div>}{error && <div className="error">{error}</div>}
    <Panel title="Outstanding / Temuan Food Index" desc="Isi corrective action, preventive action, dan due date dari temuan inspeksi. Setelah semua temuan pada task terisi, inspeksi masuk approval Atasan Site." action={<button className="secondary" onClick={load} disabled={loading}>{loading ? 'Memuat...' : 'Refresh'}</button>}>
      {loading ? <p>Memuat outstanding...</p> : <div className="table-wrap" style={{maxHeight:520, overflow:'auto'}}><table><thead><tr><th>Site</th><th>Vendor</th><th>Minggu</th><th>Parameter</th><th>Catatan</th><th>Evidence</th><th>Corrective</th><th>Preventive</th><th>Due Date</th><th>Status Task</th><th>Aksi</th></tr></thead><tbody>{rows.map(r => {
        const answer = r.food_findings?.food_inspection_answers
        const hasPlan = cleanText(r.food_findings?.corrective_action) && cleanText(r.food_findings?.preventive_action) && cleanText(r.food_findings?.due_date)
        return <tr key={r.id}><td>{r.food_weekly_tasks?.sites?.site_code || '-'}</td><td>{r.food_weekly_tasks?.food_vendors?.vendor_name || '-'}</td><td>{r.food_weekly_tasks?.week_start_date || '-'}</td><td>{answer?.food_parameters?.parameter_text || '-'}</td><td>{answer?.finding_note || '-'}</td><td>{answer?.evidence_photo_url ? <a href={answer.evidence_photo_url} target="_blank" rel="noreferrer">Lihat foto</a> : '-'}</td><td>{r.food_findings?.corrective_action || '-'}</td><td>{r.food_findings?.preventive_action || '-'}</td><td>{r.food_findings?.due_date || '-'}</td><td>{renderCell(r.food_weekly_tasks?.status)}</td><td>{!hasPlan ? <button onClick={()=>openActionPlan(r)}>Isi Tindakan</button> : r.status === 'Closed' ? <StatusPill value="Closed" /> : r.status === 'Waiting Close Approval' ? <StatusPill value="Waiting Close Approval" /> : <button onClick={()=>openCloseOutstanding(r)}>Close Outstanding</button>}</td></tr>
      })}</tbody></table>{!rows.length && <p className="muted table-empty">Belum ada outstanding / temuan.</p>}</div>}
    </Panel>
    <Panel title="Export Outstanding" action={<button onClick={()=>downloadXlsx('food-index-outstanding.xlsx', tableRows)}><Download size={16}/> Export</button>}>
      <Table rows={tableRows} empty="Belum ada data outstanding." />
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
        {closeForm.correctiveUrl && <a href={closeForm.correctiveUrl} target="_blank" rel="noreferrer">Lihat foto corrective tersimpan</a>}
        <label className="upload-line"><Upload size={18}/><span>{closeForm.preventiveFile?.name || (closeForm.preventiveUrl ? 'Foto preventive sudah ada' : 'Upload Foto Preventive')}</span><input type="file" accept="image/*" hidden onChange={e=>setCloseForm({...closeForm, preventiveFile:e.target.files?.[0] || null})}/></label>
        {closeForm.preventiveUrl && <a href={closeForm.preventiveUrl} target="_blank" rel="noreferrer">Lihat foto preventive tersimpan</a>}
      </div>
      <div className="modal-actions"><button className="secondary" onClick={()=>setCloseActive(null)} disabled={saving}>Batal</button><button onClick={saveCloseOutstanding} disabled={saving}>{saving ? 'Menyimpan...' : 'Submit Close ke Approval'}</button></div>
    </div></div>}
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
          return <div key={a.id} className={`inspection-item ${isFinding ? 'finding' : 'passed'}`}>
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
                <td>{r.corrective_photo_url ? <a href={r.corrective_photo_url} target="_blank" rel="noreferrer">Lihat foto</a> : '-'}</td>
                <td>{r.preventive_photo_url ? <a href={r.preventive_photo_url} target="_blank" rel="noreferrer">Lihat foto</a> : '-'}</td>
                <td><StatusPill value={r.status === 'Closed' ? 'Close' : 'Open'} /></td>
              </tr>
            })}</tbody>
          </table>
          {!rows.length && <p className="muted table-empty">Belum ada temuan Food Index.</p>}
        </div>
      </div>}
    </Panel>
  </div>
}
function Placeholder({ title, desc }){
  return <Panel title={title} desc={desc}><div className="card"><ClipboardCheck size={42}/><h3>{title} sedang disiapkan</h3><p>{desc}</p></div></Panel>
}
