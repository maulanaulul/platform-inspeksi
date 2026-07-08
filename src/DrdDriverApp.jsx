import React, { useEffect, useMemo, useState } from 'react'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import {
  AlertTriangle, Award, BarChart3, CalendarCheck, CheckCircle2, ClipboardCheck,
  Download, FileQuestion, FileText, LogOut, Menu, PlayCircle, ShieldCheck,
  Truck, Upload, Users, Video, XCircle, Search
} from 'lucide-react'
import logoSrgs from './assets/logo-icon.png'
import SharedAdminPanel from './SharedAdminPanel.jsx'
import './styles.css'

const supabase = createClient(import.meta.env.VITE_SUPABASE_URL, import.meta.env.VITE_SUPABASE_ANON_KEY)
const ROLE_ADMIN = ['Platform Admin', 'App Admin']
const QUESTION_CATEGORIES = ['DRD', 'Induksi Driver']

function clean(v){ return String(v ?? '').trim() }
function makeVendorCode(idx=0){ return `VEN-${Date.now().toString(36).toUpperCase()}-${String(idx+1).padStart(3,'0')}` }
function pad2(v){ return String(v).padStart(2,'0') }
function expandYear(v){
  const y = Number(v)
  if(!Number.isFinite(y)) return NaN
  if(y < 100) return y >= 70 ? 1900 + y : 2000 + y
  return y
}
function datePartsToIso(year, month, day){
  const y = expandYear(year), m = Number(month), d = Number(day)
  if(!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return ''
  if(y < 1900 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return ''
  const dt = new Date(Date.UTC(y, m - 1, d))
  if(dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return ''
  return `${y}-${pad2(m)}-${pad2(d)}`
}
function dateToIso(v){
  if(!(v instanceof Date) || Number.isNaN(v.getTime())) return ''
  return datePartsToIso(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate())
}
function excelDateToIso(v){
  if(v === null || v === undefined || v === '') return ''
  if(v instanceof Date) return dateToIso(v)
  if(typeof v === 'number'){
    const d = XLSX.SSF.parse_date_code(v)
    return d ? datePartsToIso(d.y, d.m, d.d) : ''
  }
  const raw = clean(v).replace(/[​-‍﻿]/g,'').replace(/\s+/g,' ')
  if(!raw) return ''
  if(/^\d+(\.\d+)?$/.test(raw)){
    const n = Number(raw)
    if(n > 20000 && n < 90000){
      const d = XLSX.SSF.parse_date_code(n)
      const iso = d ? datePartsToIso(d.y, d.m, d.d) : ''
      if(iso) return iso
    }
  }
  const ymd = raw.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})(?:[ T].*)?$/)
  if(ymd) return datePartsToIso(ymd[1], ymd[2], ymd[3]) || raw
  const dmy = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:[ T].*)?$/)
  if(dmy) return datePartsToIso(dmy[3], dmy[2], dmy[1]) || raw
  const short = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2})(?:[ T].*)?$/)
  if(short){
    const a = Number(short[1]), b = Number(short[2])
    // Excel sering menampilkan date cell sebagai m/d/yy, misalnya 6/1/26 untuk 2026-06-01.
    if(a > 12 && b <= 12) return datePartsToIso(short[3], b, a) || raw
    if(b > 12 && a <= 12) return datePartsToIso(short[3], a, b) || raw
    return datePartsToIso(short[3], a, b) || raw
  }
  const parsed = new Date(raw)
  if(!Number.isNaN(parsed.getTime())) return dateToIso(parsed)
  return raw
}
function isoToExcelDate(v){
  const iso = excelDateToIso(v)
  if(!isIsoDateString(iso)) return v || ''
  const [y,m,d] = iso.split('-').map(Number)
  return new Date(Date.UTC(y, m - 1, d))
}
function isIsoDateString(v){
  const raw = clean(v)
  if(!raw) return true
  const m = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  return !!m && datePartsToIso(m[1], m[2], m[3]) === raw
}
function normEmail(v){ return clean(v).toLowerCase() }
function isValidEmail(v){ const e = normEmail(v); return !e || /^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(e) }
function today(){ return new Date().toISOString().slice(0,10) }
function months(n){ const d = new Date(); d.setMonth(d.getMonth()+n); return d.toISOString().slice(0,10) }
function isAdmin(w){ return ROLE_ADMIN.includes(w?.role) }
function isGL(w){ return w?.role === 'GL' }
function isDriver(w){ return w?.role === 'Driver' }
function isExpired(date){ return !!date && date < today() }
function isOnsiteDue(p){ return !!p?.onsite_date && p.onsite_date <= today() && p.status !== 'Closed' }
function exportXlsx(name, rows){ const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows?.length ? rows : [{}]), 'Data'); XLSX.writeFile(wb, name) }
function templateXlsx(name, rows, dateColumns=[]){
  const prepared = (rows || []).map(row => {
    const copy = {...row}
    dateColumns.forEach(col => { if(copy[col]) copy[col] = isoToExcelDate(copy[col]) })
    return copy
  })
  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet(prepared)
  const headers = Object.keys(prepared[0] || {})
  dateColumns.forEach(col => {
    const c = headers.indexOf(col)
    if(c < 0) return
    for(let r=1; r<=prepared.length; r++){
      const addr = XLSX.utils.encode_cell({r, c})
      if(ws[addr]) ws[addr].z = 'yyyy-mm-dd'
    }
  })
  ws['!cols'] = headers.map(h => ({wch: ['nama_driver','question_text'].includes(h) ? 28 : ['mulai_dinas','end_masa_dinas'].includes(h) ? 14 : 16}))
  XLSX.utils.book_append_sheet(wb, ws, 'Template')
  XLSX.writeFile(wb, name, {cellDates:true})
}
async function readExcel(file){
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf,{type:'array', cellDates:true, dateNF:'yyyy-mm-dd'})
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:'', raw:true})
}
const QUESTION_IMAGE_BUCKET = 'drd-assets'
function questionImageStyle(){ return {maxWidth:'100%',maxHeight:260,objectFit:'contain',borderRadius:14,border:'1px solid #dbeafe',background:'#f8fbff',margin:'10px 0 16px'} }
async function uploadQuestionImage(file, prefix='question'){
  if(!file) return ''
  const allowed = ['image/jpeg','image/png','image/webp']
  if(!allowed.includes(file.type)) throw new Error('Format foto harus JPG, PNG, atau WEBP.')
  if(file.size > 5 * 1024 * 1024) throw new Error('Ukuran foto maksimal 5 MB.')
  const ext = (file.name.split('.').pop() || 'jpg').toLowerCase().replace(/[^a-z0-9]/g,'') || 'jpg'
  const safePrefix = String(prefix || 'question').toLowerCase().replace(/[^a-z0-9-]/g,'-').slice(0,40)
  const path = `question-images/${safePrefix}-${Date.now()}-${Math.random().toString(36).slice(2,8)}.${ext}`
  const {error} = await supabase.storage.from(QUESTION_IMAGE_BUCKET).upload(path, file, {cacheControl:'3600',upsert:false,contentType:file.type})
  if(error) throw new Error(error.message || 'Gagal upload foto soal.')
  const {data} = supabase.storage.from(QUESTION_IMAGE_BUCKET).getPublicUrl(path)
  return data?.publicUrl || ''
}
function badge(s){ const v = String(s || '-'); const cls = v.includes('Lulus') || v.includes('Closed') || v.includes('Aktif') || v.includes('Sudah') ? 'b-green' : v.includes('Open') || v.includes('Habis') || v.includes('Expired') || v.includes('Wajib') ? 'b-red' : v.includes('Cuti') || v.includes('Belum') ? 'b-amber' : 'b-blue'; return <span className={`badge ${cls}`}>{v}</span> }
async function createAuthUser({ email, password, nama, nrp, app_id, site_id, role='Driver' }){
  if(!email || !password) return {ok:true,skipped:true}
  const cleanEmail = normEmail(email)
  if(!isValidEmail(cleanEmail)) throw new Error('Format email login driver tidak valid. Gunakan format nama@domain.com.')
  const {data,error}=await supabase.functions.invoke('admin-create-user',{ body:{ email:cleanEmail, password, nama, nrp, app_id, role, site_id } })
  if(error){
    let detail = error.message || 'Gagal membuat akun Auth'
    try{
      if(error.context && typeof error.context.json === 'function'){
        const body = await error.context.json()
        detail = body?.error || body?.message || detail
      }
    }catch(_){ }
    throw new Error(detail)
  }
  if(data && data.ok===false) throw new Error(data.error||'Gagal membuat akun Auth')
  return data||{ok:true}
}

function App({ embeddedProfile=null, embeddedWork=null, onChangeApp=null } = {}){
  const [session,setSession]=useState(null), [profile,setProfile]=useState(null), [access,setAccess]=useState([]), [work,setWork]=useState(null), [loading,setLoading]=useState(true)
  useEffect(()=>{ if(embeddedProfile && embeddedWork){ setLoading(false); return } init(); const {data:{subscription}}=supabase.auth.onAuthStateChange(()=>init()); return ()=>subscription.unsubscribe() },[embeddedProfile,embeddedWork])
  async function init(){
    try{
      setLoading(true)
      const {data:{session}} = await supabase.auth.getSession(); setSession(session)
      if(!session){ setProfile(null); setAccess([]); return }
      let {data:prof} = await supabase.from('users_profile').select('*').eq('auth_user_id',session.user.id).maybeSingle()
      if(!prof && session.user.email){ const {data:byEmail}=await supabase.from('users_profile').select('*').eq('email',normEmail(session.user.email)).maybeSingle(); if(byEmail){ await supabase.from('users_profile').update({auth_user_id:session.user.id}).eq('id',byEmail.id); prof={...byEmail,auth_user_id:session.user.id} } }
      setProfile(prof)
      if(prof){ const {data:a}=await supabase.from('user_app_access').select('*,applications(*),sites(*)').eq('user_id',prof.id).eq('status','Aktif'); const rows=(a||[]).filter(r => (r.applications?.app_code || '').includes('drd')); setAccess(rows); if(rows.length===1) setWork(rows[0]) }
    }finally{ setLoading(false) }
  }
  if(embeddedProfile && embeddedWork) return <Shell profile={embeddedProfile} work={embeddedWork} setWork={()=>onChangeApp?.()} />
  if(loading) return <FullCenter text="Memuat DRD Driver..." />
  if(!session) return <Login />
  if(!profile) return <Blocked text="Akun belum dibuat/mapping di Admin Panel." />
  if(!work) return <SessionPicker profile={profile} access={access} onSelect={setWork} />
  return <Shell profile={profile} work={work} setWork={setWork} />
}

function Login(){
  const [email,setEmail]=useState(''), [password,setPassword]=useState(''), [msg,setMsg]=useState('')
  async function submit(e){ e.preventDefault(); const {error}=await supabase.auth.signInWithPassword({email:normEmail(email),password}); if(error) setMsg(error.message) }
  return <section className="login-page"><div className="login-hero"><div className="brand"><img src={logoSrgs} className="brand-logo-img"/><div><b>DRD Driver</b><span>Validasi DRD dan Induksi Driver</span></div></div><h1>DRD dan induksi driver dalam satu alur monitoring.</h1><p>Driver mengerjakan DRD otomatis saat belum valid, dan mengerjakan induksi saat kembali onsite setelah cuti.</p><div className="chips"><span>Bank Soal Langsung</span><span>Induksi Video</span><span>Dashboard Site</span></div></div><form className="login-card" onSubmit={submit}><ShieldCheck color="#2563eb" size={42}/><h2>Masuk</h2><p>Gunakan akun yang sudah dibuat dan dimapping administrator.</p><label>Email<input required value={email} onChange={e=>setEmail(e.target.value)} /></label><label>Password<input required type="password" value={password} onChange={e=>setPassword(e.target.value)} /></label>{msg&&<p className="message error">{msg}</p>}<button>Masuk</button></form></section>
}

function SelfPasswordModal({profile,onClose}){
  const [form,setForm]=useState({oldPassword:'',newPassword:'',confirmPassword:''}), [saving,setSaving]=useState(false), [msg,setMsg]=useState(''), [err,setErr]=useState('')
  async function submit(e){
    e.preventDefault(); setErr(''); setMsg('')
    try{
      const email=normEmail(profile?.email)
      const oldPassword=clean(form.oldPassword), newPassword=clean(form.newPassword), confirmPassword=clean(form.confirmPassword)
      if(!email) throw new Error('Email akun tidak ditemukan. Silakan logout lalu login ulang.')
      if(!oldPassword) throw new Error('Password lama wajib diisi.')
      if(!newPassword||newPassword.length<6) throw new Error('Password baru minimal 6 karakter.')
      if(newPassword!==confirmPassword) throw new Error('Konfirmasi password baru tidak sama.')
      if(oldPassword===newPassword) throw new Error('Password baru tidak boleh sama dengan password lama.')
      setSaving(true)
      const {error:verifyError}=await supabase.auth.signInWithPassword({email,password:oldPassword})
      if(verifyError) throw new Error('Password lama tidak sesuai.')
      const {error:updateError}=await supabase.auth.updateUser({password:newPassword})
      if(updateError) throw updateError
      setMsg('Password berhasil diganti. Gunakan password baru saat login berikutnya.')
      setForm({oldPassword:'',newPassword:'',confirmPassword:''})
      setTimeout(()=>onClose?.(),1200)
    }catch(e){ setErr(e.message||'Gagal mengganti password.') }
    finally{ setSaving(false) }
  }
  return <div className="modal-backdrop" role="dialog" aria-modal="true" onMouseDown={onClose}>
    <div className="modal-card" onMouseDown={e=>e.stopPropagation()}>
      <div className="modal-head"><div><h3>Ganti Password Saya</h3><p>Password akan diperbarui untuk akun yang sedang login.</p></div><button className="secondary" onClick={onClose}>×</button></div>
      <form className="form-grid" onSubmit={submit} autoComplete="off">
        {err&&<p className="message error">{err}</p>}
        {msg&&<p className="message">{msg}</p>}
        <label>Email<input value={profile?.email||''} disabled/></label>
        <label>Password Lama<input type="password" autoComplete="current-password" value={form.oldPassword} onChange={e=>setForm({...form,oldPassword:e.target.value})}/></label>
        <label>Password Baru<input type="password" autoComplete="new-password" value={form.newPassword} onChange={e=>setForm({...form,newPassword:e.target.value})} placeholder="Minimal 6 karakter"/></label>
        <label>Konfirmasi Password Baru<input type="password" autoComplete="new-password" value={form.confirmPassword} onChange={e=>setForm({...form,confirmPassword:e.target.value})}/></label>
        <div className="row-actions"><button type="button" className="secondary" onClick={onClose}>Batal</button><button disabled={saving}>{saving?'Menyimpan...':'Simpan Password'}</button></div>
      </form>
    </div>
  </div>
}
function Blocked({text}){ return <div className="full-center"><div className="panel"><h2>{text}</h2><button onClick={()=>supabase.auth.signOut()}>Logout</button></div></div> }
function FullCenter({text}){ return <div className="full-center"><div className="panel"><h2>{text}</h2></div></div> }
function SessionPicker({profile,access,onSelect}){ if(!access.length) return <Blocked text="Akses DRD Driver belum dimapping."/>; return <section className="context-page"><div className="context-card"><h1>Pilih Sesi DRD Driver</h1><p>{profile.nama} · {profile.email}</p><div className="cards">{access.map(a=><div className="card" key={a.id}><h3>{a.role}</h3><p>{a.sites?.site_name || 'All Site'}</p><button onClick={()=>onSelect(a)}>Masuk</button></div>)}</div><button className="secondary" onClick={()=>supabase.auth.signOut()}>Logout</button></div></section> }

function Shell({profile,work,setWork}){
  const [page,setPage]=useState(()=>isDriver(work)?'test':'dashboard'), [sidebar,setSidebar]=useState(false), [passwordOpen,setPasswordOpen]=useState(false)
  let nav=isDriver(work) ? [] : [['dashboard','Dashboard',BarChart3]]
  if(isAdmin(work)) nav.push(['admin','Admin Panel',Users],['questions','Bank Soal',FileQuestion],['videos','Video Induksi',Video],['drivers','Master Driver',Truck],['cuti','Periode Cuti',CalendarCheck],['results','Hasil DRD & Induksi',Award])
  if(isGL(work)) nav.push(['drivers','Master Driver',Truck],['cuti','Periode Cuti',CalendarCheck],['results','Hasil DRD & Induksi',Award])
  if(isDriver(work)) nav.push(['test','Tes DRD',PlayCircle],['induction','Induksi Driver',Video],['results','Hasil Saya',Award])
  const safePage=nav.some(n=>n[0]===page)?page:(nav[0]?.[0]||page)
  useEffect(()=>{ if(safePage!==page) setPage(safePage) },[safePage,page])
  const title=nav.find(n=>n[0]===safePage)?.[1]||'DRD Driver'
  return <div className="app real-ui-shell">
    <button className="icon mobile mobile-menu-trigger" onClick={()=>setSidebar(true)}><Menu size={20}/></button>{sidebar&&<button className="mobile-nav-backdrop" aria-label="Tutup navigasi" onClick={()=>setSidebar(false)}/>}<aside className={sidebar?'sidebar open':'sidebar'}><button className="sidebar-close" onClick={()=>setSidebar(false)}>×</button><div className="brand dark sidebar-brand"><img src={logoSrgs} alt="DRD" className="brand-logo-img"/><div><b>DRD Driver</b><span>{work.role} · {work.sites?.site_name}</span></div></div><div className="nav">{nav.map(([k,l,I])=><button key={k} className={page===k?'active':''} onClick={()=>{setPage(k);setSidebar(false)}}><I size={18}/> {l}</button>)}</div><div className="sidebar-info-cards"><div className="sidebar-info-card"><b>{work.sites?.site_name||'All Site'}</b><span>Lokasi kerja aktif</span></div><div className="sidebar-info-card"><b>DRD + Induksi</b><span>Validasi driver, cuti, onsite, dan induksi</span></div></div><div className="sidebar-card"><b>Monitoring Driver</b><p>DRD otomatis muncul saat driver belum valid. Induksi muncul saat driver kembali onsite setelah cuti.</p></div><div className="side-bottom"><button className="secondary" onClick={()=>setWork(null)}>Ganti Aplikasi</button></div></aside>
    <main className="main"><div className="top app-header"><div className="header-copy"><h1>{title}</h1><p>Kelola DRD, induksi driver, masa dinas, cuti, hasil tes, dan pencapaian site.</p></div><div className="header-actions redesign-actions"><div className="user-chip"><Users size={16}/><span>{profile?.nama} · {work.sites?.site_name||'All Site'}</span></div><button className="secondary" onClick={()=>setPasswordOpen(true)}>Ganti Password</button><button className="secondary" onClick={()=>setWork(null)}>Ganti Aplikasi</button><button className="secondary" onClick={()=>supabase.auth.signOut()}><LogOut size={16}/>Logout</button></div></div>{passwordOpen&&<SelfPasswordModal profile={profile} onClose={()=>setPasswordOpen(false)}/>}<DRD page={safePage} profile={profile} work={work}/></main>
  </div>
}
function DRD({page,profile,work}){ if(page==='dashboard')return <Dashboard work={work}/>; if(page==='admin')return <SharedAdminPanel profile={profile} context={work}/>; if(page==='questions')return <Questions/>; if(page==='videos')return <InductionVideos/>; if(page==='drivers')return <MasterDriver profile={profile} work={work}/>; if(page==='cuti')return <CutiPeriods profile={profile} work={work}/>; if(page==='test')return <DriverTest profile={profile}/>; if(page==='induction')return <DriverInduction profile={profile}/>; return <Results profile={profile} work={work}/> }

function Panel({title,desc,action,children}){ return <div className="panel"><div className="panel-head"><div><h2>{title}</h2>{desc&&<p className="muted">{desc}</p>}</div>{action}</div>{children}</div> }
function DataTable({rows, actions}){ const [q,setQ]=useState(''); if(!rows?.length) return <p className="muted">Belum ada data.</p>; const cols=Object.keys(rows[0]); const filtered=rows.filter(r=>!q||Object.values(r).some(v=>String(v??'').toLowerCase().includes(q.toLowerCase()))); return <div className="data-table-block"><div className="table-search"><Search size={16}/><input placeholder="Search row data..." value={q} onChange={e=>setQ(e.target.value)}/></div><div className="table-wrap"><table className="table"><thead><tr>{cols.map(c=><th key={c}>{c}</th>)}{actions&&<th>AKSI</th>}</tr></thead><tbody>{filtered.map((r,i)=>{const originalIdx=rows.indexOf(r);return <tr key={r.id || i}>{cols.map(c=><td key={c}>{String(r[c]??'')}</td>)}{actions&&<td>{actions(r,originalIdx)}</td>}</tr>})}</tbody></table></div>{filtered.length===0&&<p className="muted">Tidak ada data sesuai pencarian.</p>}</div> }

async function fetchAllPages(buildQuery, pageSize=1000){
  const rows=[]
  for(let from=0;;from+=pageSize){
    const {data,error}=await buildQuery().range(from, from+pageSize-1)
    if(error) throw error
    const batch=data||[]
    rows.push(...batch)
    if(batch.length<pageSize) break
  }
  return rows
}
function ScrollTable({rows, actions, height=420}){
  return <div style={{maxHeight:height,overflow:'auto',border:'1px solid #dbeafe',borderRadius:18,padding:12,background:'#fff'}}><DataTable rows={rows} actions={actions}/></div>
}


function DashboardDateFilter({ dateFrom, dateTo, setDateFrom, setDateTo, onClear, sites=[], siteFilter='', setSiteFilter, showSiteFilter=false }) {
  return <Panel title="Filter Dashboard" desc="KPI, achievement, chart, dan row data dashboard mengikuti filter tanggal dan site ini.">
    <div className="form-grid">
      <label>Dari Tanggal<input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} /></label>
      <label>Sampai Tanggal<input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} /></label>
      {showSiteFilter&&<label>Site<select value={siteFilter} onChange={e=>setSiteFilter(e.target.value)}><option value="">Semua Site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>}
      <button type="button" className="secondary" onClick={onClear}>Reset Filter</button>
    </div>
  </Panel>
}

function Dashboard({work}){
  const [drivers,setDrivers]=useState([]), [attempts,setAttempts]=useState([]), [periods,setPeriods]=useState([]), [sites,setSites]=useState([])
  const [dateFrom,setDateFrom]=useState(''), [dateTo,setDateTo]=useState(''), [siteFilter,setSiteFilter]=useState('')
  useEffect(()=>{load()},[work.id,dateFrom,dateTo,siteFilter])
  async function load(){
    const admin=isAdmin(work)
    const buildDrivers=()=>{ let q=supabase.from('drivers').select('*,sites(site_code,site_name),vendors(vendor_name)').eq('status','Aktif').order('created_at',{ascending:false}); if(admin&&siteFilter) q=q.eq('site_id',siteFilter); if(!admin) q=q.eq('site_id',work.site_id); return q }
    const buildAttempts=()=>{ let q=supabase.from('drd_attempts').select('*,drivers(site_id,nama_driver,nrp_driver,email,sites(site_code,site_name))').order('created_at',{ascending:false}); if(dateFrom) q=q.gte('created_at',dateFrom); if(dateTo) q=q.lte('created_at',dateTo+'T23:59:59'); return q }
    const buildPeriods=()=>{ let q=supabase.from('drd_induction_periods').select('*,drivers(nama_driver,nrp_driver,site_id,sites(site_code,site_name))').order('created_at',{ascending:false}); if(dateFrom) q=q.gte('created_at',dateFrom); if(dateTo) q=q.lte('created_at',dateTo+'T23:59:59'); return q }
    const [d,a,p,s]=await Promise.all([
      fetchAllPages(buildDrivers),
      fetchAllPages(buildAttempts),
      fetchAllPages(buildPeriods),
      admin ? fetchAllPages(()=>supabase.from('sites').select('id,site_code,site_name').order('site_code')) : Promise.resolve([])
    ])
    const scopedAttempts=(a||[]).filter(r=>admin ? (!siteFilter || r.drivers?.site_id===siteFilter) : r.drivers?.site_id===work.site_id)
    const scopedPeriods=(p||[]).filter(r=>admin ? (!siteFilter || r.drivers?.site_id===siteFilter) : r.drivers?.site_id===work.site_id)
    setDrivers(d||[]); setAttempts(scopedAttempts); setPeriods(scopedPeriods); if(admin) setSites(s||[])
  }
  const latestDrd=new Map()
  attempts.filter(a=>(a.test_type||'DRD')==='DRD').forEach(a=>{ if(!latestDrd.has(a.driver_id)) latestDrd.set(a.driver_id,a) })
  function currentInductionPeriod(d){
    return periods.find(p=>String(p.driver_id)===String(d.id) && String(p.masa_dinas_end_date||'')===String(d.end_masa_dinas||''))
      || periods.find(p=>String(p.driver_id)===String(d.id) && !p.masa_dinas_end_date)
      || null
  }
  const drdOk=drivers.filter(d=>{ const a=latestDrd.get(d.id); return a?.status==='Lulus' && (!a.valid_until || a.valid_until>=today()) }).length
  const total=drivers.length, drdBelum=total-drdOk, drdAch=total?Math.round(drdOk/total*100):0
  const expiredDrivers=drivers.filter(d=>isExpired(d.end_masa_dinas))
  const duePeriods=drivers.map(currentInductionPeriod).filter(Boolean).filter(isOnsiteDue)
  const closedInduksi=drivers.filter(d=>currentInductionPeriod(d)?.status==='Closed').length
  const openPeriodInput=expiredDrivers.filter(d=>!currentInductionPeriod(d)).length
  const openInduksi=duePeriods.length
  // Achievement induksi berbasis masa dinas driver aktif saat ini per site.
  // Driver yang masa dinasnya habis tetapi belum diinput periode tetap masuk denominator agar ACH tidak terlihat 100% palsu.
  const inductionTarget=Math.max(expiredDrivers.length, openInduksi+closedInduksi+openPeriodInput)
  const inductionAch=inductionTarget?Math.round(closedInduksi/inductionTarget*100):100

  const bySite={}
  const bySiteInduksi={}
  drivers.forEach(d=>{
    const code=d.sites?.site_code||'-'
    const siteName=d.sites?.site_name||'-'
    bySite[code]??={site:code,total:0,drd_ok:0,drd_belum:0,achievement:0}
    bySite[code].total++
    const a=latestDrd.get(d.id)
    if(a?.status==='Lulus'&&(!a.valid_until||a.valid_until>=today())) bySite[code].drd_ok++

    bySiteInduksi[code]??={site:code,site_name:siteName,total_driver:0,masa_dinas_habis:0,butuh_input_periode:0,wajib_induksi:0,open_induksi:0,closed_induksi:0,belum_selesai:0,achievement_induksi:100}
    bySiteInduksi[code].total_driver++
    const period=currentInductionPeriod(d)
    const expired=isExpired(d.end_masa_dinas)
    const due=period ? isOnsiteDue(period) : false
    const closed=period?.status==='Closed'
    if(expired) bySiteInduksi[code].masa_dinas_habis++
    if(expired && !period) bySiteInduksi[code].butuh_input_periode++
    if(closed || due){
      bySiteInduksi[code].wajib_induksi++
      if(due) bySiteInduksi[code].open_induksi++
      if(closed) bySiteInduksi[code].closed_induksi++
    }
  })
  Object.values(bySite).forEach(x=>{ x.drd_belum=x.total-x.drd_ok; x.achievement=x.total?Math.round(x.drd_ok/x.total*100):0 })
  Object.values(bySiteInduksi).forEach(x=>{
    const target=Math.max(x.masa_dinas_habis, x.wajib_induksi + x.butuh_input_periode)
    x.target_induksi=target
    x.belum_selesai=Math.max(target-x.closed_induksi,0)
    x.achievement_induksi=target?Math.round(x.closed_induksi/target*100):100
  })
  const siteRows=Object.values(bySite).sort((a,b)=>b.achievement-a.achievement)
  const inductionSiteRows=Object.values(bySiteInduksi).sort((a,b)=>b.achievement_induksi-a.achievement_induksi || String(a.site).localeCompare(String(b.site)))
  const driverDrdRows=drivers.map(d=>{ const a=latestDrd.get(d.id); const valid=a?.status==='Lulus' && (!a.valid_until || a.valid_until>=today()); return {site:d.sites?.site_code||'-', nama_driver:d.nama_driver, nrp_driver:d.nrp_driver, email:d.email||'-', vendor:d.vendors?.vendor_name||'-', status_drd:valid?'Sudah DRD':'Belum DRD', nilai:a?.score ?? '-', tanggal_test:a?.submitted_at ? String(a.submitted_at).slice(0,10) : '-', valid_until:a?.valid_until||'-'} })
  const sudahDrdRows=driverDrdRows.filter(r=>r.status_drd==='Sudah DRD').sort((a,b)=>String(a.site).localeCompare(String(b.site))||String(a.nama_driver).localeCompare(String(b.nama_driver)))
  const belumDrdRows=driverDrdRows.filter(r=>r.status_drd==='Belum DRD').sort((a,b)=>String(a.site).localeCompare(String(b.site))||String(a.nama_driver).localeCompare(String(b.nama_driver)))
  const inductionRows=drivers.map(d=>{
    const p=currentInductionPeriod(d)
    const expired=isExpired(d.end_masa_dinas)
    const due=p ? isOnsiteDue(p) : false
    const status=!expired ? 'Belum Wajib' : !p ? 'Butuh Input Periode' : p.status==='Closed' ? 'Closed' : due ? 'Open - Wajib Induksi' : 'Menunggu Onsite'
    return {
      site:d.sites?.site_code||'-',
      nama_driver:d.nama_driver,
      nrp_driver:d.nrp_driver,
      end_masa_dinas:d.end_masa_dinas||'-',
      cuti_mulai:p?.cuti_start_date||'-',
      onsite:p?.onsite_date||'-',
      status_induksi:status,
      status_periode:p?.status||'-',
      completed_at:p?.completed_at ? String(p.completed_at).slice(0,10) : '-'
    }
  }).sort((a,b)=>String(a.site).localeCompare(String(b.site))||String(a.status_induksi).localeCompare(String(b.status_induksi))||String(a.nama_driver).localeCompare(String(b.nama_driver)))
  const detailDriverRows=[...driverDrdRows].sort((a,b)=>String(a.site).localeCompare(String(b.site))||String(a.nama_driver).localeCompare(String(b.nama_driver)))
  const detailDriverFileName=`detail-driver-dashboard-drd-${siteFilter ? (sites.find(s=>String(s.id)===String(siteFilter))?.site_code || 'site-terpilih') : 'semua-site'}.xlsx`
  const inductionDetailFileName=`detail-induksi-driver-${siteFilter ? (sites.find(s=>String(s.id)===String(siteFilter))?.site_code || 'site-terpilih') : 'semua-site'}.xlsx`
  return <div className="stack">
    <DashboardDateFilter dateFrom={dateFrom} dateTo={dateTo} setDateFrom={setDateFrom} setDateTo={setDateTo} sites={sites} siteFilter={siteFilter} setSiteFilter={setSiteFilter} showSiteFilter={isAdmin(work)} onClear={()=>{setDateFrom('');setDateTo('');setSiteFilter('')}}/>
    <div className="kpi-grid"><Kpi title="Total Driver" value={total} icon={<Truck/>}/><Kpi title="Sudah DRD" value={drdOk} icon={<CheckCircle2/>}/><Kpi title="Belum DRD" value={drdBelum} icon={<AlertTriangle/>}/><Kpi title="Achievement DRD" value={`${drdAch}%`} icon={<BarChart3/>}/><Kpi title="Open Induksi" value={openInduksi} icon={<Video/>}/><Kpi title="Induksi Closed" value={closedInduksi} icon={<ShieldCheck/>}/><Kpi title="Achievement Induksi" value={`${inductionAch}%`} icon={<Award/>}/><Kpi title="Masa Dinas Habis" value={expiredDrivers.length} icon={<CalendarCheck/>}/></div>
    <Panel title="Dashboard DRD per Site" desc="Pencapaian DRD dihitung dari driver aktif yang sudah lulus dan belum expired." action={<div className="row-actions"><button onClick={()=>exportXlsx('achievement-drd-site.xlsx',siteRows)}><Download size={16}/> Export Summary</button><button className="secondary" onClick={()=>exportXlsx(detailDriverFileName,detailDriverRows)}><Download size={16}/> Export Detail Driver</button></div>}>
      <div className="site-chart">{siteRows.map(r=><div className="site-bar" key={r.site}><div className="site-meta"><b>{r.site}</b><span>{r.drd_ok}/{r.total} · {r.achievement}%</span></div><div className="bar"><span style={{width:`${Math.min(r.achievement,100)}%`}}/></div></div>)}</div><ScrollTable rows={siteRows} height={320}/>
    </Panel>
    <Panel title="Dashboard Induksi Driver ACH per Site" desc="Achievement induksi dihitung per site dari driver aktif yang masa dinasnya habis, open induksi, closed induksi, dan driver yang masih butuh input periode cuti." action={<div className="row-actions"><button onClick={()=>exportXlsx('achievement-induksi-driver-persite.xlsx',inductionSiteRows)}><Download size={16}/> Export Summary</button><button className="secondary" onClick={()=>exportXlsx(inductionDetailFileName,inductionRows)}><Download size={16}/> Export Detail Driver</button></div>}>
      <div className="site-chart">{inductionSiteRows.map(r=><div className="site-bar" key={r.site}><div className="site-meta"><b>{r.site}</b><span>{r.closed_induksi}/{r.target_induksi} · {r.achievement_induksi}%</span></div><div className="bar"><span style={{width:`${Math.min(r.achievement_induksi,100)}%`}}/></div></div>)}</div>
      <ScrollTable rows={inductionSiteRows} height={340}/>
    </Panel>
    <Panel title="Detail Driver Dashboard DRD" desc="Detail seluruh driver aktif yang menjadi denominator dashboard. Jika filter site BRCB dipilih, jumlah row export harus sama dengan total dashboard BRCB." action={<button onClick={()=>exportXlsx(detailDriverFileName,detailDriverRows)}><Download size={16}/> Export Detail Driver</button>}><ScrollTable rows={detailDriverRows} height={420}/></Panel>
    <Panel title="Driver Sudah DRD" desc="Driver aktif yang sudah lulus DRD dan valid_until belum expired." action={<button onClick={()=>exportXlsx('driver-sudah-drd.xlsx',sudahDrdRows)}><Download size={16}/> Export</button>}><ScrollTable rows={sudahDrdRows} height={360}/></Panel>
    <Panel title="Driver Belum DRD" desc="Driver aktif yang belum pernah lulus DRD atau masa berlaku DRD-nya sudah expired." action={<button onClick={()=>exportXlsx('driver-belum-drd.xlsx',belumDrdRows)}><Download size={16}/> Export</button>}><ScrollTable rows={belumDrdRows} height={420}/></Panel>
    <Panel title="Detail Dashboard Induksi Driver" desc="Status induksi driver per nama: belum wajib, butuh input periode, menunggu onsite, open wajib induksi, atau closed." action={<button onClick={()=>exportXlsx(inductionDetailFileName,inductionRows)}><Download size={16}/> Export</button>}><div className="summary-strip"><span><b>{inductionTarget}</b> Target Induksi</span><span><b>{openPeriodInput}</b> Butuh Input Periode</span><span><b>{openInduksi}</b> Open Induksi</span><span><b>{closedInduksi}</b> Closed</span><span><b>{inductionAch}%</b> Achievement</span></div><ScrollTable rows={inductionRows} height={420}/></Panel>
  </div>
}
function Kpi({title,value,icon}){ return <div className="kpi"><div><span>{title}</span><strong>{value}</strong></div><div className="kpi-icon">{icon}</div></div> }

function Questions(){
  const emptyQuestion={category:'DRD',question_text:'',option_a:'',option_b:'',option_c:'',option_d:'',correct_answer:'A',image_url:''}
  const [questions,setQuestions]=useState([]), [preview,setPreview]=useState([]), [msg,setMsg]=useState(''), [form,setForm]=useState(emptyQuestion), [editOpen,setEditOpen]=useState(false), [editId,setEditId]=useState(null), [editForm,setEditForm]=useState({...emptyQuestion,status:'Aktif'}), [formImageFile,setFormImageFile]=useState(null), [editImageFile,setEditImageFile]=useState(null), [saving,setSaving]=useState(false)
  useEffect(()=>{load()},[])
  async function load(){ const {data:q}=await supabase.from('drd_questions').select('*').order('category').order('created_at',{ascending:false}); setQuestions(q||[]) }
  async function file(e){ const f=e.target.files?.[0]; if(!f)return; const rows=await readExcel(f); setPreview(rows.map((r,i)=>({row:i+2,category:clean(r.category)||'DRD',question_text:clean(r.question_text),option_a:clean(r.option_a),option_b:clean(r.option_b),option_c:clean(r.option_c),option_d:clean(r.option_d),correct_answer:clean(r.correct_answer).toUpperCase()}))) }
  async function importRows(){ const ok=preview.filter(r=>QUESTION_CATEGORIES.includes(r.category)&&r.question_text&&r.option_a&&r.option_b&&r.option_c&&r.option_d&&['A','B','C','D'].includes(r.correct_answer)); if(ok.length!==preview.length)return setMsg('Masih ada baris invalid. Category harus DRD atau Induksi Driver.'); const {error}=await supabase.from('drd_questions').insert(ok.map(({row,...r})=>({...r,package_id:null,status:'Aktif',image_url:null}))); setMsg(error?error.message:`Import berhasil ${ok.length} soal.`); setPreview([]); load() }
  async function save(e){
    e.preventDefault(); setMsg(''); setSaving(true)
    try{
      if(!QUESTION_CATEGORIES.includes(form.category)) throw new Error('Kategori harus DRD atau Induksi Driver.')
      if(!form.question_text || !form.option_a || !form.option_b || !form.option_c || !form.option_d) throw new Error('Soal dan semua pilihan jawaban wajib diisi.')
      if(!['A','B','C','D'].includes(form.correct_answer)) throw new Error('Jawaban harus A, B, C, atau D.')
      const image_url = formImageFile ? await uploadQuestionImage(formImageFile, form.category) : (form.image_url || null)
      const {error}=await supabase.from('drd_questions').insert({...form,image_url,package_id:null,status:'Aktif'})
      if(error) throw error
      setMsg('Soal tersimpan.'); setForm({...emptyQuestion}); setFormImageFile(null); load()
    }catch(err){ setMsg(err.message || 'Gagal menyimpan soal.') }
    finally{ setSaving(false) }
  }
  function openEdit(q){ setEditId(q.id); setEditForm({category:q.category||'DRD',question_text:q.question_text||'',option_a:q.option_a||'',option_b:q.option_b||'',option_c:q.option_c||'',option_d:q.option_d||'',correct_answer:q.correct_answer||'A',status:q.status||'Aktif',image_url:q.image_url||''}); setEditImageFile(null); setEditOpen(true); setMsg('') }
  function closeEdit(){ setEditOpen(false); setEditId(null); setEditForm({...emptyQuestion,status:'Aktif'}); setEditImageFile(null) }
  async function updateQuestion(e){
    e.preventDefault(); setSaving(true)
    try{
      if(!QUESTION_CATEGORIES.includes(editForm.category)) throw new Error('Kategori harus DRD atau Induksi Driver.')
      if(!editForm.question_text || !editForm.option_a || !editForm.option_b || !editForm.option_c || !editForm.option_d) throw new Error('Soal dan semua pilihan jawaban wajib diisi.')
      if(!['A','B','C','D'].includes(editForm.correct_answer)) throw new Error('Correct answer harus A, B, C, atau D.')
      const image_url = editImageFile ? await uploadQuestionImage(editImageFile, editForm.category) : (editForm.image_url || null)
      const {error}=await supabase.from('drd_questions').update({...editForm,image_url}).eq('id',editId)
      if(error) throw error
      setMsg('Soal berhasil diupdate.'); closeEdit(); load()
    }catch(err){ setMsg(err.message || 'Gagal update soal.') }
    finally{ setSaving(false) }
  }
  async function deleteQuestion(q){
    if(!confirm('Delete soal ini? Soal akan dinonaktifkan agar histori test tetap aman.')) return
    const {error}=await supabase.from('drd_questions').update({status:'Nonaktif'}).eq('id',q.id)
    setMsg(error?error.message:'Soal berhasil di-delete/nonaktifkan.')
    load()
  }
  const rows=questions.map(q=>({id:q.id, kategori:q.category, soal:q.question_text, foto:q.image_url?'Ada':'-', jawaban:q.correct_answer, status:q.status}))
  return <div className="stack"><Panel title="Bank Soal DRD & Induksi" desc="Tidak ada paket soal. Semua soal aktif pada kategori DRD otomatis muncul untuk tes DRD; kategori Induksi Driver otomatis muncul untuk induksi."><div className="import-actions"><button className="secondary" onClick={()=>templateXlsx('template-bank-soal-drd-induksi.xlsx',[{category:'DRD',question_text:'Apa tindakan saat mengantuk?',option_a:'Tetap jalan',option_b:'Berhenti dan istirahat',option_c:'Tambah kecepatan',option_d:'Abaikan',correct_answer:'B'},{category:'Induksi Driver',question_text:'Apa yang wajib dilakukan setelah kembali onsite?',option_a:'Langsung kerja',option_b:'Mengikuti induksi',option_c:'Abaikan arahan',option_d:'Tidak perlu lapor',correct_answer:'B'}])}><Download size={16}/> Download Template</button><label className="upload-line"><Upload size={16}/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={file}/></label>{preview.length>0&&<button onClick={importRows}>Submit Import</button>}</div>{preview.length>0&&<div className="upload-preview"><DataTable rows={preview}/></div>}</Panel><Panel title="Tambah Soal Manual"><form className="form-grid" onSubmit={save}><label>Kategori<select value={form.category} onChange={e=>setForm({...form,category:e.target.value})}>{QUESTION_CATEGORIES.map(c=><option key={c}>{c}</option>)}</select></label><label>Soal<input value={form.question_text} onChange={e=>setForm({...form,question_text:e.target.value})}/></label><label>Jawaban<select value={form.correct_answer} onChange={e=>setForm({...form,correct_answer:e.target.value})}>{['A','B','C','D'].map(x=><option key={x}>{x}</option>)}</select></label><label>Opsi A<input value={form.option_a} onChange={e=>setForm({...form,option_a:e.target.value})}/></label><label>Opsi B<input value={form.option_b} onChange={e=>setForm({...form,option_b:e.target.value})}/></label><label>Opsi C<input value={form.option_c} onChange={e=>setForm({...form,option_c:e.target.value})}/></label><label>Opsi D<input value={form.option_d} onChange={e=>setForm({...form,option_d:e.target.value})}/></label><label className="full">Foto Soal (opsional)<input type="file" accept="image/png,image/jpeg,image/webp" onChange={e=>setFormImageFile(e.target.files?.[0]||null)}/></label>{formImageFile&&<div className="full"><img src={URL.createObjectURL(formImageFile)} alt="Preview foto soal" style={questionImageStyle()}/></div>}<button disabled={saving}>{saving?'Menyimpan...':'Simpan Soal'}</button></form>{msg&&<p className="message">{msg}</p>}</Panel><Panel title="Row Data Bank Soal" action={<button onClick={()=>exportXlsx('bank-soal-drd-induksi.xlsx',rows)}><Download size={16}/> Export</button>}><DataTable rows={rows} actions={(r,i)=><div className="row-actions"><button className="secondary small" onClick={()=>openEdit(questions[i])}>Edit</button><button className="danger small" onClick={()=>deleteQuestion(questions[i])}>Delete</button></div>}/></Panel>{editOpen&&<div className="modal-backdrop" onClick={closeEdit}><div className="modal-card" onClick={e=>e.stopPropagation()}><div className="modal-head"><h3>Edit Soal</h3><button type="button" className="secondary small" onClick={closeEdit}>Tutup</button></div><form className="form-grid" onSubmit={updateQuestion}><label>Kategori<select value={editForm.category} onChange={e=>setEditForm({...editForm,category:e.target.value})}>{QUESTION_CATEGORIES.map(c=><option key={c}>{c}</option>)}</select></label><label>Status<select value={editForm.status} onChange={e=>setEditForm({...editForm,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label><label className="full">Soal<textarea rows="3" value={editForm.question_text} onChange={e=>setEditForm({...editForm,question_text:e.target.value})}/></label><label>Opsi A<input value={editForm.option_a} onChange={e=>setEditForm({...editForm,option_a:e.target.value})}/></label><label>Opsi B<input value={editForm.option_b} onChange={e=>setEditForm({...editForm,option_b:e.target.value})}/></label><label>Opsi C<input value={editForm.option_c} onChange={e=>setEditForm({...editForm,option_c:e.target.value})}/></label><label>Opsi D<input value={editForm.option_d} onChange={e=>setEditForm({...editForm,option_d:e.target.value})}/></label><label>Correct Answer<select value={editForm.correct_answer} onChange={e=>setEditForm({...editForm,correct_answer:e.target.value})}>{['A','B','C','D'].map(x=><option key={x}>{x}</option>)}</select></label><label className="full">Ganti Foto Soal (opsional)<input type="file" accept="image/png,image/jpeg,image/webp" onChange={e=>setEditImageFile(e.target.files?.[0]||null)}/></label>{editImageFile&&<div className="full"><img src={URL.createObjectURL(editImageFile)} alt="Preview foto soal baru" style={questionImageStyle()}/></div>}{!editImageFile&&editForm.image_url&&<div className="full"><img src={editForm.image_url} alt="Foto soal" style={questionImageStyle()}/><br/><button type="button" className="danger small" onClick={()=>setEditForm({...editForm,image_url:''})}>Hapus Foto</button></div>}<div className="modal-actions"><button type="button" className="secondary" onClick={closeEdit}>Batal</button><button disabled={saving}>{saving?'Menyimpan...':'Simpan Perubahan'}</button></div></form></div></div>}</div>
}

function InductionVideos(){
  const [rows,setRows]=useState([]), [form,setForm]=useState({title:'Video Induksi Driver',video_url:'',status:'Aktif'}), [msg,setMsg]=useState('')
  useEffect(()=>{load()},[])
  async function load(){ const {data}=await supabase.from('drd_induction_videos').select('*').order('created_at',{ascending:false}); setRows(data||[]) }
  async function save(e){ e.preventDefault(); const {error}=await supabase.from('drd_induction_videos').insert(form); setMsg(error?error.message:'Video induksi tersimpan.'); setForm({title:'Video Induksi Driver',video_url:'',status:'Aktif'}); load() }
  async function toggle(r){ await supabase.from('drd_induction_videos').update({status:r.status==='Aktif'?'Nonaktif':'Aktif'}).eq('id',r.id); load() }
  return <div className="stack"><Panel title="Upload / Link Video Induksi" desc="Administrator mengisi link video induksi. Driver wajib menonton sampai selesai sebelum mengerjakan soal induksi."><form className="form-grid" onSubmit={save}><label>Judul Video<input value={form.title} onChange={e=>setForm({...form,title:e.target.value})}/></label><label>URL Video MP4 / Public URL<input required value={form.video_url} onChange={e=>setForm({...form,video_url:e.target.value})} placeholder="https://.../video.mp4"/></label><label>Status<select value={form.status} onChange={e=>setForm({...form,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label><button>Simpan Video</button></form>{msg&&<p className="message">{msg}</p>}</Panel><Panel title="Row Data Video Induksi"><DataTable rows={rows.map(r=>({judul:r.title,url:r.video_url,status:r.status,created:r.created_at}))} actions={(r,i)=><button className="secondary small" onClick={()=>toggle(rows[i])}>{rows[i].status==='Aktif'?'Nonaktifkan':'Aktifkan'}</button>}/></Panel></div>
}

function MasterDriver({profile,work}){
  const [drivers,setDrivers]=useState([]), [vendors,setVendors]=useState([]), [sites,setSites]=useState([]), [preview,setPreview]=useState([]), [msg,setMsg]=useState(''), [editing,setEditing]=useState(null), [modalOpen,setModalOpen]=useState(false), [importing,setImporting]=useState(false), [importProgress,setImportProgress]=useState({current:0,total:0}), [authGenerating,setAuthGenerating]=useState(false), [authProgress,setAuthProgress]=useState({processed:0,created:0,skipped:0,failed:0,remaining:null}), [masaDinasPreview,setMasaDinasPreview]=useState([]), [masaDinasUpdating,setMasaDinasUpdating]=useState(false), [masaDinasProgress,setMasaDinasProgress]=useState({current:0,total:0})
  const emptyForm={site_id:work.site_id||'',nama_driver:'',nrp_driver:'',email:'',password:'',vendor_id:'',status:'Aktif',mulai_dinas:'',end_masa_dinas:''}
  const [form,setForm]=useState(emptyForm)
  useEffect(()=>{load()},[work.id])
  async function load(){
    // WAJIB pakai fetchAllPages. Supabase .select() biasa hanya mengembalikan default ±1000 row,
    // sehingga saat data seluruh site besar, sebagian driver BRCB tidak ikut tertarik di Master Driver/export.
    const buildDrivers=()=>{
      let q=supabase.from('drivers').select('*,sites(site_code,site_name),vendors(vendor_name)').order('created_at',{ascending:false})
      if(!isAdmin(work)) q=q.eq('site_id',work.site_id)
      return q
    }
    const [d,v,s]=await Promise.all([
      fetchAllPages(buildDrivers),
      fetchAllPages(()=>supabase.from('vendors').select('*').order('vendor_name')),
      fetchAllPages(()=>supabase.from('sites').select('*').neq('site_code','JIEP').order('site_code'))
    ])
    setDrivers(d||[]); setVendors(v||[]); setSites(s||[])
  }
  function reset(){ setEditing(null); setModalOpen(false); setForm({...emptyForm,site_id:work.site_id||''}) }
  function edit(d){
    setEditing(d.id)
    setForm({site_id:d.site_id||work.site_id||'',nama_driver:d.nama_driver||'',nrp_driver:d.nrp_driver||'',email:d.email||'',password:'',vendor_id:d.vendor_id||'',status:d.status||'Aktif',mulai_dinas:d.mulai_dinas||'',end_masa_dinas:d.end_masa_dinas||''})
    setModalOpen(true)
  }
  async function save(e){
    e.preventDefault(); setMsg('')
    const siteId=isAdmin(work)?form.site_id:work.site_id
    const nrpKey=clean(form.nrp_driver).toLowerCase(), emailKey=normEmail(form.email)
    const dup=drivers.find(d=>String(d.id)!==String(editing||'') && (clean(d.nrp_driver).toLowerCase()===nrpKey || (emailKey && normEmail(d.email)===emailKey)))
    if(dup) return setMsg(dup.nrp_driver&&clean(dup.nrp_driver).toLowerCase()===nrpKey?'NRP driver sudah terdaftar.':'Email driver sudah terdaftar.')
    if(form.email && !isValidEmail(form.email)) return setMsg('Format email driver tidak valid. Gunakan format nama@domain.com.')
    if(form.email&&form.password){ try{ await createAuthUser({email:form.email,password:form.password,nama:form.nama_driver,nrp:form.nrp_driver,app_id:work.app_id||work.applications?.id,site_id:siteId,role:'Driver'}) }catch(err){ return setMsg(err.message||'Gagal membuat akun login driver') } }
    const payload={site_id:siteId,nama_driver:form.nama_driver,nrp_driver:form.nrp_driver,email:normEmail(form.email)||null,vendor_id:form.vendor_id||null,status:form.status,mulai_dinas:form.mulai_dinas||null,end_masa_dinas:form.end_masa_dinas||null,updated_at:new Date().toISOString()}
    const {error}=editing
      ? await supabase.from('drivers').update(payload).eq('id',editing)
      : await supabase.from('drivers').upsert(payload,{onConflict:'nrp_driver'})
    if(error) return setMsg(error.message)
    setMsg(editing?'Driver berhasil diupdate.':'Driver berhasil disimpan.')
    reset(); load()
  }
  async function remove(d){
    if(!confirm('Delete/nonaktifkan driver ini?'))return
    const {error}=await supabase.from('drivers').update({status:'Nonaktif',updated_at:new Date().toISOString()}).eq('id',d.id)
    setMsg(error?error.message:'Driver dinonaktifkan.'); load()
  }
  async function previewFile(file){
    if(!file)return
    const rows=await readExcel(file)
    const seenNrp=new Set(), seenEmail=new Set()
    const mapped=rows.map((r,i)=>{
      const siteCode=clean(r.site_code).toUpperCase()
      const site=isAdmin(work)?sites.find(s=>s.site_code===siteCode):sites.find(s=>s.id===work.site_id)
      const vendorName=clean(r.vendor_name)
      const vendor=vendors.find(v=>clean(v.vendor_name).toLowerCase()===vendorName.toLowerCase())
      let error=''
      const nrpKey=clean(r.nrp_driver).toLowerCase(), emailKey=normEmail(r.email)
      if(!clean(r.nama_driver))error='nama_driver wajib'
      else if(!clean(r.nrp_driver))error='nrp_driver wajib'
      else if(emailKey && !isValidEmail(emailKey))error='email tidak valid'
      else if(seenNrp.has(nrpKey))error='nrp_driver double di file import'
      else if(emailKey && seenEmail.has(emailKey))error='email double di file import'
      else if(drivers.some(d=>clean(d.nrp_driver).toLowerCase()===nrpKey))error='nrp_driver sudah terdaftar'
      else if(emailKey && drivers.some(d=>normEmail(d.email)===emailKey))error='email sudah terdaftar'
      else if(isAdmin(work)&&!site)error='site_code tidak ditemukan'
      seenNrp.add(nrpKey); if(emailKey) seenEmail.add(emailKey)
      return {row:i+2,site_code:isAdmin(work)?siteCode:work.sites?.site_code,nama_driver:clean(r.nama_driver),nrp_driver:clean(r.nrp_driver),email:normEmail(r.email),password:clean(r.password),vendor_name:vendorName,status:clean(r.status)||'Aktif',mulai_dinas:excelDateToIso(r.mulai_dinas),end_masa_dinas:excelDateToIso(r.end_masa_dinas),site_id:site?.id,vendor_id:vendor?.id,error}
    })
    setPreview(mapped)
  }
  async function importRows(){
    setMsg('')
    const valid=preview.filter(r=>!r.error)
    if(valid.length!==preview.length)return setMsg('Masih ada baris invalid.')
    const errors=[]
    let successCount=0
    setImporting(true)
    setImportProgress({current:0,total:valid.length})
    try{
      for(const [i,r] of valid.entries()){
        try{
          let vendorId=r.vendor_id
          if(r.vendor_name&&!vendorId){
            const {data:v,error:vendorError}=await supabase.from('vendors').insert({vendor_code:makeVendorCode(i),vendor_name:r.vendor_name,status:'Aktif'}).select().single()
            if(vendorError) throw vendorError
            vendorId=v?.id
          }
          const {error:driverError}=await supabase.from('drivers').upsert({site_id:r.site_id,nama_driver:r.nama_driver,nrp_driver:r.nrp_driver,email:r.email||null,vendor_id:vendorId||null,status:r.status,mulai_dinas:r.mulai_dinas||null,end_masa_dinas:r.end_masa_dinas||null},{onConflict:'nrp_driver'})
          if(driverError) throw driverError
          successCount += 1
          if(r.email&&r.password){
            try{
              await createAuthUser({email:r.email,password:r.password,nama:r.nama_driver,nrp:r.nrp_driver,app_id:work.app_id||work.applications?.id,site_id:r.site_id,role:'Driver'})
            }catch(authErr){
              errors.push(`Baris ${r.row}: data driver tersimpan, tapi akun login gagal dibuat (${authErr.message||String(authErr)})`)
            }
          }
        }catch(e){
          errors.push(`Baris ${r.row}: ${e.message||String(e)}`)
        }finally{
          setImportProgress({current:i+1,total:valid.length})
        }
      }
      if(errors.length){
        setMsg(`Import selesai: ${successCount} driver tersimpan. Catatan akun login: ${errors.slice(0,3).join(' | ')}`)
        setPreview([]); load(); return
      }
      setMsg(`Import berhasil ${successCount} driver.`); setPreview([]); load()
    } finally {
      setImporting(false)
      setImportProgress({current:0,total:0})
    }
  }
  function downloadMasaDinasTemplate(){
    const rows=[...drivers]
      .sort((a,b)=>String(a.sites?.site_code||'').localeCompare(String(b.sites?.site_code||''))||String(a.nama_driver||'').localeCompare(String(b.nama_driver||'')))
      .map(d=>({
        site_code:d.sites?.site_code||'',
        nrp_driver:d.nrp_driver||'',
        nama_driver:d.nama_driver||'',
        mulai_dinas:d.mulai_dinas||'',
        end_masa_dinas:d.end_masa_dinas||'',
        status:d.status||'Aktif'
      }))
    templateXlsx('template-update-masa-dinas-driver.xlsx', rows.length ? rows : [{site_code:work.sites?.site_code||'',nrp_driver:'',nama_driver:'',mulai_dinas:'2026-04-01',end_masa_dinas:'2026-05-01',status:'Aktif'}], ['mulai_dinas','end_masa_dinas'])
  }
  async function previewMasaDinasFile(file){
    if(!file) return
    const rows=await readExcel(file)
    const byNrp=new Map(drivers.map(d=>[clean(d.nrp_driver).toLowerCase(),d]))
    const seen=new Set()
    const mapped=rows.map((r,i)=>{
      const nrpKey=clean(r.nrp_driver).toLowerCase()
      const existing=byNrp.get(nrpKey)
      const mulai=excelDateToIso(r.mulai_dinas)
      const end=excelDateToIso(r.end_masa_dinas)
      const status=clean(r.status)
      const siteCode=clean(r.site_code).toUpperCase()
      let error=''
      if(!nrpKey) error='nrp_driver wajib'
      else if(seen.has(nrpKey)) error='nrp_driver double di file update'
      else if(!existing) error='nrp_driver tidak ditemukan di Master Driver aktif/scope site ini'
      else if(isAdmin(work) && siteCode && clean(existing.sites?.site_code).toUpperCase()!==siteCode) error='site_code tidak sesuai dengan Master Driver'
      else if(!isAdmin(work) && existing.site_id!==work.site_id) error='driver bukan milik site aktif'
      else if(mulai && !isIsoDateString(mulai)) error='format mulai_dinas harus yyyy-mm-dd atau dd/mm/yyyy'
      else if(end && !isIsoDateString(end)) error='format end_masa_dinas harus yyyy-mm-dd atau dd/mm/yyyy'
      else if(!mulai && !end && !status) error='isi minimal salah satu: mulai_dinas, end_masa_dinas, atau status'
      else if(status && !['Aktif','Nonaktif'].includes(status)) error='status hanya boleh Aktif/Nonaktif'
      seen.add(nrpKey)
      return {
        row:i+2,
        site_code:existing?.sites?.site_code||siteCode||'-',
        nama_driver:existing?.nama_driver||clean(r.nama_driver),
        nrp_driver:clean(r.nrp_driver),
        mulai_dinas_lama:existing?.mulai_dinas||'-',
        mulai_dinas_baru:mulai||existing?.mulai_dinas||'',
        end_masa_dinas_lama:existing?.end_masa_dinas||'-',
        end_masa_dinas_baru:end||existing?.end_masa_dinas||'',
        status_lama:existing?.status||'-',
        status_baru:status||existing?.status||'',
        driver_id:existing?.id||'',
        error
      }
    })
    setMasaDinasPreview(mapped)
  }
  async function updateMasaDinasRows(){
    setMsg('')
    const valid=masaDinasPreview.filter(r=>!r.error)
    if(!valid.length) return setMsg('Belum ada baris update masa dinas yang valid.')
    if(valid.length!==masaDinasPreview.length) return setMsg('Masih ada baris invalid pada preview update masa dinas.')
    setMasaDinasUpdating(true)
    setMasaDinasProgress({current:0,total:valid.length})
    const errors=[]
    let successCount=0
    try{
      for(const [i,r] of valid.entries()){
        try{
          const payload={updated_at:new Date().toISOString()}
          payload.mulai_dinas=r.mulai_dinas_baru||null
          payload.end_masa_dinas=r.end_masa_dinas_baru||null
          if(r.status_baru) payload.status=r.status_baru
          const {error}=await supabase.from('drivers').update(payload).eq('id',r.driver_id)
          if(error) throw error
          successCount += 1
        }catch(e){
          errors.push(`Baris ${r.row}: ${e.message||String(e)}`)
        }finally{
          setMasaDinasProgress({current:i+1,total:valid.length})
        }
      }
      if(errors.length) setMsg(`Update masa dinas selesai sebagian. Berhasil: ${successCount}. Error: ${errors.slice(0,3).join(' | ')}`)
      else setMsg(`Update masa dinas berhasil: ${successCount} driver.`)
      setMasaDinasPreview([])
      await load()
    }finally{
      setMasaDinasUpdating(false)
      setMasaDinasProgress({current:0,total:0})
    }
  }
  async function generateMissingDriverAuth(){
    if(authGenerating) return
    const appId=work.app_id||work.applications?.id
    if(!appId) return setMsg('App ID DRD tidak ditemukan. Coba refresh halaman atau login ulang.')
    const scopeText=isAdmin(work)?'semua site yang belum punya auth':'site ini'
    if(!confirm(`Buat akun login driver untuk ${scopeText}? Password awal akan memakai NRP masing-masing driver.`)) return
    setMsg('')
    setAuthGenerating(true)
    setAuthProgress({processed:0,created:0,skipped:0,failed:0,remaining:null})
    let totalProcessed=0, totalCreated=0, totalSkipped=0, totalFailed=0, remaining=null
    try{
      for(let batch=0; batch<100; batch++){
        const {data,error}=await supabase.functions.invoke('bulk-create-driver-auth',{ body:{ app_id:appId, site_id:isAdmin(work)?null:work.site_id, limit:20 } })
        if(error){
          let detail=error.message||'Gagal generate auth driver'
          try{
            if(error.context && typeof error.context.json === 'function'){
              const body=await error.context.json()
              detail=body?.error||body?.message||detail
            }
          }catch(_){ }
          throw new Error(detail)
        }
        if(data?.ok===false) throw new Error(data.error||'Gagal generate auth driver')
        const summary=data?.summary||{}
        totalProcessed += Number(summary.processed||0)
        totalCreated += Number(summary.created_auth||0)
        totalSkipped += Number(summary.skipped||0)
        totalFailed += Number(summary.failed||0)
        remaining = Number(summary.remaining_after||0)
        setAuthProgress({processed:totalProcessed,created:totalCreated,skipped:totalSkipped,failed:totalFailed,remaining})
        if(!remaining || Number(summary.processed||0)===0) break
        await new Promise(resolve=>setTimeout(resolve,300))
      }
      setMsg(`Generate auth selesai. Diproses: ${totalProcessed}. Auth baru: ${totalCreated}. Skip: ${totalSkipped}. Gagal: ${totalFailed}. Sisa belum punya auth: ${remaining ?? 0}.`)
      load()
    }catch(e){
      setMsg(e.message||String(e))
    }finally{
      setAuthGenerating(false)
    }
  }
  const table=drivers.map(d=>({nama:d.nama_driver,nrp:d.nrp_driver,email:d.email||'-',site:d.sites?.site_code||'-',vendor:d.vendors?.vendor_name||'-',mulai_dinas:d.mulai_dinas||'-',end_masa_dinas:d.end_masa_dinas||'-',status_masa_dinas:isExpired(d.end_masa_dinas)?'Masa Dinas Habis':'Aktif',status:d.status}))
  const renderDriverForm = (submitLabel, cancelLabel = null) => <form className="form-grid" onSubmit={save} autoComplete="off">
    {isAdmin(work)&&<label>Site<select required value={form.site_id} onChange={e=>setForm({...form,site_id:e.target.value})}><option value="">Pilih site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>}
    <label>Nama Driver<input required value={form.nama_driver} onChange={e=>setForm({...form,nama_driver:e.target.value})}/></label>
    <label>NRP<input required value={form.nrp_driver} onChange={e=>setForm({...form,nrp_driver:e.target.value})}/></label>
    <label>Email<input type="email" name="driver_email_no_autofill" autoComplete="new-email" value={form.email} onChange={e=>setForm({...form,email:e.target.value})}/></label>
    <label>Password<input type="password" name="driver_password_no_autofill" autoComplete="new-password" value={form.password} onChange={e=>setForm({...form,password:e.target.value})} placeholder={editing?'Kosongkan jika tidak ubah akun':'Password awal'}/></label>
    <label>Vendor<select value={form.vendor_id} onChange={e=>setForm({...form,vendor_id:e.target.value})}><option value="">Tanpa vendor</option>{vendors.map(v=><option key={v.id} value={v.id}>{v.vendor_name}</option>)}</select></label>
    <label>Mulai Dinas<input type="date" value={form.mulai_dinas} onChange={e=>setForm({...form,mulai_dinas:e.target.value})}/></label>
    <label>End Masa Dinas<input type="date" value={form.end_masa_dinas} onChange={e=>setForm({...form,end_masa_dinas:e.target.value})}/></label>
    <label>Status<select value={form.status} onChange={e=>setForm({...form,status:e.target.value})}><option>Aktif</option><option>Nonaktif</option></select></label>
    <button>{submitLabel}</button>{cancelLabel&&<button type="button" className="secondary" onClick={reset}>{cancelLabel}</button>}
  </form>
  return <div className="stack">
    <Panel title="Master Driver DRD" desc="Tambah driver baru dan masa dinas. Untuk mengubah data existing, klik Edit pada tabel agar form terbuka dalam modal.">
      {!editing && renderDriverForm("Simpan Driver + Akun")}
      {editing && <p className="message">Sedang mengedit driver di modal. Tutup modal untuk kembali tambah driver baru.</p>}
      {msg&&<p className="message">{msg}</p>}
    </Panel>
    {modalOpen && <div className="modal-backdrop" onClick={reset}>
      <div className="modal-card" onClick={e=>e.stopPropagation()}>
        <div className="modal-head">
          <div><h2>Edit Master Driver DRD</h2><p className="muted">Ubah data driver di sini agar tidak tercampur dengan form tambah driver.</p></div>
          <button type="button" className="secondary small" onClick={reset}>Tutup</button>
        </div>
        {renderDriverForm("Update Driver", "Batal Edit")}
        {msg&&<p className="message">{msg}</p>}
      </div>
    </div>}
    <Panel title="Upload Bulk Master Driver DRD" desc="Template tidak memakai kode unik. Cukup isi site_code, nama_driver, nrp_driver, email/password bila perlu akun, vendor_name bila ada, serta tanggal masa dinas.">
      <style>{'@keyframes srgsSpin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}'}</style>
      <div className="import-actions"><button className="secondary" disabled={importing||authGenerating} onClick={()=>templateXlsx('template-master-driver-drd.xlsx',[{site_code:'BAYA',nama_driver:'Budi',nrp_driver:'D-001',email:'budi@company.co.id',password:'password123',vendor_name:'PBM',mulai_dinas:'2026-04-01',end_masa_dinas:'2026-05-01',status:'Aktif'}], ['mulai_dinas','end_masa_dinas'])}><Download size={16}/> Download Template Excel</button><label className="upload-line"><Upload size={16}/> Upload Excel<input type="file" accept=".xlsx,.xls" disabled={importing||authGenerating} onChange={e=>previewFile(e.target.files?.[0])}/></label>{preview.length>0&&<button disabled={importing||authGenerating} onClick={importRows}>{importing?'Mengimpor...':'Submit Import Valid'}</button>}{isAdmin(work)&&<button className="secondary" disabled={importing||authGenerating} onClick={generateMissingDriverAuth}><Users size={16}/> {authGenerating?'Generate Auth...':'Generate Auth Driver Belum Ada'}</button>}</div>
      {importing&&<div className="message" style={{display:'flex',alignItems:'center',gap:10}}><span aria-hidden="true" style={{width:16,height:16,border:'2px solid #bfdbfe',borderTopColor:'#2563eb',borderRadius:'50%',display:'inline-block',animation:'srgsSpin 0.8s linear infinite'}}/><span>Import sedang diproses... {importProgress.current}/{importProgress.total} baris. Mohon tunggu untuk file besar.</span></div>}
      {authGenerating&&<div className="message" style={{display:'flex',alignItems:'center',gap:10}}><span aria-hidden="true" style={{width:16,height:16,border:'2px solid #bfdbfe',borderTopColor:'#2563eb',borderRadius:'50%',display:'inline-block',animation:'srgsSpin 0.8s linear infinite'}}/><span>Generate auth driver sedang diproses... diproses {authProgress.processed}, auth baru {authProgress.created}, skip {authProgress.skipped}, gagal {authProgress.failed}{authProgress.remaining!==null?`, sisa ${authProgress.remaining}`:''}. Password awal memakai NRP.</span></div>}
      {preview.length>0&&<div className="upload-preview"><DataTable rows={preview}/></div>}
    </Panel>
    <Panel title="Update Masa Dinas Driver Existing" desc="Download template berisi driver aktif pada scope site saat ini, ubah mulai_dinas/end_masa_dinas/status di Excel, lalu upload kembali untuk update massal tanpa membuat driver baru.">
      <div className="import-actions">
        <button className="secondary" disabled={masaDinasUpdating||importing||authGenerating} onClick={downloadMasaDinasTemplate}><Download size={16}/> Download Template Masa Dinas</button>
        <label className="upload-line"><Upload size={16}/> Upload Update Masa Dinas<input type="file" accept=".xlsx,.xls" disabled={masaDinasUpdating||importing||authGenerating} onChange={e=>previewMasaDinasFile(e.target.files?.[0])}/></label>
        {masaDinasPreview.length>0&&<button disabled={masaDinasUpdating||importing||authGenerating} onClick={updateMasaDinasRows}>{masaDinasUpdating?'Mengupdate...':'Submit Update Masa Dinas'}</button>}
        {masaDinasPreview.length>0&&<button className="secondary" disabled={masaDinasUpdating} onClick={()=>setMasaDinasPreview([])}>Batal Preview</button>}
      </div>
      {masaDinasUpdating&&<div className="message" style={{display:'flex',alignItems:'center',gap:10}}><span aria-hidden="true" style={{width:16,height:16,border:'2px solid #bfdbfe',borderTopColor:'#2563eb',borderRadius:'50%',display:'inline-block',animation:'srgsSpin 0.8s linear infinite'}}/><span>Update masa dinas sedang diproses... {masaDinasProgress.current}/{masaDinasProgress.total} baris.</span></div>}
      {masaDinasPreview.length>0&&<div className="upload-preview"><DataTable rows={masaDinasPreview.map(({driver_id,...r})=>r)}/></div>}
    </Panel>
    <Panel title="Row Data Driver" desc="Aksi edit/delete dibuat rapi dalam satu tabel. Delete akan mengubah status menjadi Nonaktif agar histori tetap aman." action={<button onClick={()=>exportXlsx('master-driver-drd.xlsx',table)}><Download size={16}/> Export</button>}>
      <DataTable rows={table} actions={(r,i)=><div className="row-actions"><button className="secondary small" onClick={()=>edit(drivers[i])}>Edit</button><button className="danger small" onClick={()=>remove(drivers[i])}>Delete</button></div>}/>
    </Panel>
  </div>
}
function CutiPeriods({profile,work}){
  const [drivers,setDrivers]=useState([]), [periods,setPeriods]=useState([]), [form,setForm]=useState({driver_id:'',cuti_start_date:'',onsite_date:''}), [msg,setMsg]=useState(''), [modalOpen,setModalOpen]=useState(false)
  useEffect(()=>{load()},[work.id])
  async function load(){
    let dq=supabase.from('drivers').select('*,sites(site_code,site_name)').eq('status','Aktif')
    if(!isAdmin(work))dq=dq.eq('site_id',work.site_id)
    const [{data:d},{data:p}]=await Promise.all([
      dq,
      supabase.from('drd_induction_periods').select('*,drivers(nama_driver,nrp_driver,site_id,end_masa_dinas,sites(site_code,site_name))').order('created_at',{ascending:false})
    ])
    const scopedDrivers=d||[]
    const scopedPeriods=(p||[]).filter(x=>isAdmin(work)||x.drivers?.site_id===work.site_id)
    setDrivers(scopedDrivers); setPeriods(scopedPeriods)
    setForm(f=>({...f,driver_id:f.driver_id||scopedDrivers?.[0]?.id||''}))
  }
  function hasFilledPeriodForCurrentDinas(d){
    return periods.some(p=>p.driver_id===d.id && String(p.masa_dinas_end_date||p.drivers?.end_masa_dinas||'')===String(d.end_masa_dinas||'') && p.cuti_start_date && p.onsite_date)
  }
  const outstanding=drivers.filter(d=>isExpired(d.end_masa_dinas)&&!hasFilledPeriodForCurrentDinas(d))
  const selectedDriver=drivers.find(x=>x.id===form.driver_id)
  function openPeriodModal(d=null){
    setMsg('')
    if(d) setForm({driver_id:d.id,cuti_start_date:'',onsite_date:''})
    else setForm(f=>({...f,driver_id:f.driver_id || drivers?.[0]?.id || '',cuti_start_date:f.cuti_start_date||'',onsite_date:f.onsite_date||''}))
    setModalOpen(true)
  }
  function closePeriodModal(){ setModalOpen(false); setMsg('') }
  async function save(e){
    e.preventDefault()
    const d=drivers.find(x=>x.id===form.driver_id)
    if(!d) return setMsg('Driver belum dipilih.')
    if(!form.cuti_start_date || !form.onsite_date) return setMsg('Tanggal mulai cuti dan tanggal onsite wajib diisi.')
    const payload={driver_id:form.driver_id,site_id:d?.site_id,masa_dinas_end_date:d?.end_masa_dinas||null,cuti_start_date:form.cuti_start_date,onsite_date:form.onsite_date,status:'Open',updated_at:new Date().toISOString()}

    // Hindari upsert ON CONFLICT karena database lama bisa belum punya unique constraint
    // pada kombinasi driver_id + masa_dinas_end_date. Flow ini lebih aman untuk project
    // yang sudah beberapa kali migrasi SQL.
    const existing = periods.find(p =>
      String(p.driver_id) === String(payload.driver_id) &&
      String(p.masa_dinas_end_date || '') === String(payload.masa_dinas_end_date || '')
    )

    let error = null
    if (existing?.id) {
      const res = await supabase.from('drd_induction_periods').update(payload).eq('id', existing.id)
      error = res.error
    } else {
      const res = await supabase.from('drd_induction_periods').insert(payload)
      error = res.error
    }

    if(error){
      setMsg(error.message)
    } else {
      setMsg('Periode cuti/onsite berhasil disimpan. Card outstanding akan hilang untuk periode masa dinas ini.')
      setForm({driver_id:'',cuti_start_date:'',onsite_date:''})
      setModalOpen(false)
    }
    load()
  }
  const rows=periods.map(p=>({driver:p.drivers?.nama_driver,nrp:p.drivers?.nrp_driver,site:p.drivers?.sites?.site_code,end_masa_dinas:p.masa_dinas_end_date||p.drivers?.end_masa_dinas||'-',cuti_mulai:p.cuti_start_date||'-',onsite:p.onsite_date||'-',status:p.status,alert:isOnsiteDue(p)?'Open - Driver wajib induksi':'-'}))
  return <div className="stack">
    <Panel title="Outstanding Input Periode Cuti" desc="Card akan muncul saat masa dinas driver habis dan belum ada input tanggal cuti + tanggal onsite untuk periode masa dinas tersebut. Setelah diisi, card akan hilang sampai ada periode masa dinas berikutnya.">
      {outstanding.length ? <div className="cards-grid">{outstanding.map(d=><div className="card" key={d.id}><h3>{d.nama_driver}</h3><p><b>NRP:</b> {d.nrp_driver}</p><p><b>Site:</b> {d.sites?.site_code || '-'}</p><p><b>End Masa Dinas:</b> {d.end_masa_dinas}</p><button onClick={()=>openPeriodModal(d)}>Isi Periode Cuti</button></div>)}</div> : <p className="message">Tidak ada outstanding input periode cuti saat ini.</p>}
    </Panel>
    <Panel title="Set Periode Cuti & Onsite" desc="Input tanggal cuti dan onsite sekarang memakai box modal supaya form tidak memenuhi halaman. Saat onsite sudah tiba dan induksi belum closed, status dihitung Open Induksi." action={<button onClick={()=>openPeriodModal()}>+ Tambah Periode</button>}>
      {msg&&<p className="message">{msg}</p>}
      <p className="muted">Klik kartu outstanding atau tombol tambah periode untuk membuka form input periode cuti.</p>
    </Panel>
    <Panel title="Row Data Periode Cuti & Induksi" action={<button onClick={()=>exportXlsx('periode-cuti-induksi.xlsx',rows)}><Download size={16}/> Export</button>}><DataTable rows={rows}/></Panel>

    {modalOpen && <div className="modal-backdrop" onClick={closePeriodModal}>
      <div className="modal-card period-modal" onClick={e=>e.stopPropagation()}>
        <div className="modal-head">
          <div><h2>Isi Periode Cuti & Onsite</h2><p className="muted">Pilih driver, tanggal mulai cuti, dan tanggal onsite. Periode ini menjadi dasar alert induksi saat driver kembali onsite.</p></div>
          <button type="button" className="secondary icon-btn" onClick={closePeriodModal}>×</button>
        </div>
        {selectedDriver && <div className="summary-strip period-driver-summary"><span><b>{selectedDriver.nama_driver}</b> Driver</span><span><b>{selectedDriver.nrp_driver}</b> NRP</span><span><b>{selectedDriver.sites?.site_code||'-'}</b> Site</span><span><b>{selectedDriver.end_masa_dinas||'-'}</b> End Masa Dinas</span></div>}
        <form className="form-grid" onSubmit={save}>
          <label>Driver<select value={form.driver_id} onChange={e=>setForm({...form,driver_id:e.target.value})}><option value="">Pilih driver</option>{drivers.map(d=><option key={d.id} value={d.id}>{d.nama_driver} · {d.nrp_driver} · end {d.end_masa_dinas||'-'}</option>)}</select></label>
          <label>Mulai Cuti<input type="date" value={form.cuti_start_date} onChange={e=>setForm({...form,cuti_start_date:e.target.value})}/></label>
          <label>Tanggal Onsite<input type="date" value={form.onsite_date} onChange={e=>setForm({...form,onsite_date:e.target.value})}/></label>
          <div className="modal-actions"><button type="button" className="secondary" onClick={closePeriodModal}>Batal</button><button>Simpan Periode</button></div>
        </form>
        {msg&&<p className="message">{msg}</p>}
      </div>
    </div>}
  </div>
}

function DriverTest({profile}){ return <DriverExam profile={profile} type="DRD" /> }
function DriverInduction({profile}){ return <DriverExam profile={profile} type="Induksi Driver" /> }
function DriverExam({profile,type}){
  const [driver,setDriver]=useState(null), [questions,setQuestions]=useState([]), [active,setActive]=useState(false), [answers,setAnswers]=useState({}), [idx,setIdx]=useState(0), [msg,setMsg]=useState(''), [video,setVideo]=useState(null), [videoDone,setVideoDone]=useState(false), [period,setPeriod]=useState(null), [latestDrd,setLatestDrd]=useState(null)
  useEffect(()=>{load()},[profile.email,type])
  async function load(){
    const {data:d}=await supabase.from('drivers').select('*').eq('email',profile.email).maybeSingle(); setDriver(d)
    if(!d)return
    const {data:a}=await supabase.from('drd_attempts').select('*').eq('driver_id',d.id).eq('test_type','DRD').order('created_at',{ascending:false}).limit(1); setLatestDrd(a?.[0]||null)
    if(type==='Induksi Driver'){ const {data:v}=await supabase.from('drd_induction_videos').select('*').eq('status','Aktif').order('created_at',{ascending:false}).limit(1); setVideo(v?.[0]||null); const {data:p}=await supabase.from('drd_induction_periods').select('*').eq('driver_id',d.id).eq('status','Open').lte('onsite_date',today()).order('onsite_date',{ascending:false}).limit(1); setPeriod(p?.[0]||null) }
  }
  const drdValid=latestDrd?.status==='Lulus' && (!latestDrd.valid_until || latestDrd.valid_until>=today())
  const canShow = type==='DRD' ? !drdValid : !!period
  async function start(){ if(type==='Induksi Driver' && !videoDone) return setMsg('Tonton video induksi sampai selesai terlebih dahulu.'); const {data:q,error}=await supabase.from('drd_questions').select('*').eq('category',type).eq('status','Aktif').order('created_at'); if(error)return setMsg(error.message); if(!q?.length)return setMsg(`Bank soal kategori ${type} masih kosong.`); setQuestions(q); setActive(true); setAnswers({}); setIdx(0); setMsg('') }
  async function submit(){ const total=questions.length, correct=questions.filter(q=>answers[q.id]===q.correct_answer).length, pass=correct===total, score=total?Math.round(correct/total*100):0; const {data:att,error}=await supabase.from('drd_attempts').insert({assignment_id:null,driver_id:driver.id,package_id:null,test_type:type,induction_period_id:period?.id||null,submitted_at:new Date().toISOString(),score,total_questions:total,correct_count:correct,status:pass?'Lulus':'Tidak Lulus',valid_until:type==='DRD'&&pass?months(6):null}).select().single(); if(error)return setMsg(error.message); await supabase.from('drd_answers').insert(questions.map(q=>({attempt_id:att.id,question_id:q.id,selected_answer:answers[q.id]||null,is_correct:answers[q.id]===q.correct_answer}))); if(type==='Induksi Driver'&&pass&&period) await supabase.from('drd_induction_periods').update({status:'Closed',completed_at:new Date().toISOString()}).eq('id',period.id); setMsg(pass ? (type==='DRD'?`Lulus DRD. Berlaku sampai ${months(6)}.`:'Induksi closed untuk periode ini.') : `Belum lulus. Benar ${correct}/${total}. Silakan ulangi ${type}.`); setActive(false); setQuestions([]); load() }
  if(!driver)return <Panel title={type}><p className="message error">Email login belum terhubung ke Master Driver.</p></Panel>
  if(active){ const q=questions[idx]; return <Panel title={`${type} - Soal ${idx+1}/${questions.length}`} desc="Passing grade 100%. Semua jawaban harus benar."><h2>{q.question_text}</h2>{q.image_url&&<img src={q.image_url} alt="Foto soal" style={questionImageStyle()}/>} { [['A',q.option_a],['B',q.option_b],['C',q.option_c],['D',q.option_d]].map(([k,v])=><div key={k} className={'option '+(answers[q.id]===k?'selected':'')} onClick={()=>setAnswers({...answers,[q.id]:k})}><b>{k}.</b> {v}</div>)}<br/><div className="actions"><button className="secondary" disabled={idx===0} onClick={()=>setIdx(idx-1)}>Sebelumnya</button><button className="secondary" disabled={idx===questions.length-1} onClick={()=>setIdx(idx+1)}>Selanjutnya</button><button disabled={Object.keys(answers).length<questions.length} onClick={submit}>Submit</button></div></Panel> }
  if(!canShow) return <Panel title={type}>{type==='DRD'?<p className="message">Status DRD masih valid. Tidak ada tes aktif.</p>:<p className="message">Tidak ada induksi aktif. Induksi muncul saat driver sudah onsite setelah cuti dan periode induksi masih Open.</p>}</Panel>
  return <Panel title={type} desc={type==='DRD'?'Status driver belum DRD / DRD expired, maka tombol mulai tes otomatis muncul.':'Tonton video sampai selesai, lalu kerjakan soal induksi. Jika ada jawaban salah, induksi harus diulang.'}>{msg&&<p className="message">{msg}</p>}{type==='Induksi Driver'&&<div className="stack"><div className="video-box">{video?.video_url?<video width="100%" controls onEnded={()=>setVideoDone(true)}><source src={video.video_url}/></video>:<p className="message error">Video induksi aktif belum tersedia. Hubungi Administrator.</p>}</div>{video&&badge(videoDone?'Video Selesai':'Wajib Tonton Video')}</div>}<br/><button disabled={type==='Induksi Driver' && (!videoDone || !video)} onClick={start}>Mulai {type}</button></Panel>
}

function Results({profile,work}){
  const [attempts,setAttempts]=useState([]), [periods,setPeriods]=useState([])
  useEffect(()=>{load()},[work.id])
  async function load(){
    let attemptRows=[]; let periodRows=[]
    if(isDriver(work)){
      const {data:d}=await supabase.from('drivers').select('*').eq('email',profile.email).maybeSingle()
      if(d){
        const [attemptRes, periodRes]=await Promise.all([
          supabase.from('drd_attempts').select('*,drivers(nama_driver,nrp_driver,site_id,sites(site_code,site_name))').eq('driver_id',d.id).order('created_at',{ascending:false}),
          supabase.from('drd_induction_periods').select('*,drivers(nama_driver,nrp_driver,site_id,sites(site_code,site_name))').eq('driver_id',d.id).order('created_at',{ascending:false})
        ])
        attemptRows=attemptRes.data||[]; periodRows=periodRes.data||[]
      }
    } else {
      const [attemptRes, periodRes]=await Promise.all([
        supabase.from('drd_attempts').select('*,drivers(nama_driver,nrp_driver,site_id,sites(site_code,site_name))').order('created_at',{ascending:false}),
        supabase.from('drd_induction_periods').select('*,drivers(nama_driver,nrp_driver,site_id,sites(site_code,site_name))').order('created_at',{ascending:false})
      ])
      attemptRows=(attemptRes.data||[]).filter(r=>isAdmin(work)||r.drivers?.site_id===work.site_id)
      periodRows=(periodRes.data||[]).filter(r=>isAdmin(work)||r.drivers?.site_id===work.site_id)
    }
    setAttempts(attemptRows); setPeriods(periodRows)
  }
  const drdRows=attempts.filter(r=>(r.test_type||'DRD')==='DRD').map(r=>({driver:r.drivers?.nama_driver,nrp:r.drivers?.nrp_driver,site:r.drivers?.sites?.site_code,score:r.score,benar:r.correct_count,total:r.total_questions,status:(r.valid_until&&r.valid_until<today())?'Expired':r.status,valid_until:r.valid_until||'-',tanggal:r.submitted_at?new Date(r.submitted_at).toLocaleString('id-ID'):'-'}))
  const inductionAttempts=attempts.filter(r=>r.test_type==='Induksi Driver')
  const latestByPeriod=new Map()
  inductionAttempts.forEach(a=>{ const key=a.induction_period_id||`attempt-${a.id}`; if(!latestByPeriod.has(key)) latestByPeriod.set(key,a) })
  const inductionRows=periods.map(p=>{ const a=latestByPeriod.get(p.id); return {driver:p.drivers?.nama_driver,nrp:p.drivers?.nrp_driver,site:p.drivers?.sites?.site_code,end_masa_dinas:p.masa_dinas_end_date||'-',cuti_mulai:p.cuti_start_date||'-',onsite:p.onsite_date||'-',status_induksi:p.status,score:a?.score??'-',benar:a?.correct_count??'-',total:a?.total_questions??'-',status_test:a?.status||'-',tanggal_induksi:a?.submitted_at?new Date(a.submitted_at).toLocaleString('id-ID'):'-',completed_at:p.completed_at?new Date(p.completed_at).toLocaleString('id-ID'):'-' } })
  return <div className="stack">
    <Panel title="Hasil DRD" desc="Riwayat tes DRD dipisahkan dari induksi agar monitoring validasi driver lebih jelas." action={<button onClick={()=>exportXlsx('hasil-drd.xlsx',drdRows)}><Download size={16}/> Export DRD</button>}><DataTable rows={drdRows}/></Panel>
    <Panel title="Hasil Induksi Driver" desc="Monitoring induksi berdasarkan periode cuti/onsite. Status Open berarti driver sudah perlu induksi tetapi periode belum closed." action={<button onClick={()=>exportXlsx('hasil-induksi-driver.xlsx',inductionRows)}><Download size={16}/> Export Induksi</button>}><DataTable rows={inductionRows}/></Panel>
  </div>
}


function AdminPanelDRD({profile,work}){
  const [apps,setApps]=useState([]), [sites,setSites]=useState([]), [profiles,setProfiles]=useState([]), [access,setAccess]=useState([])
  const [questions,setQuestions]=useState([]), [videos,setVideos]=useState([]), [drivers,setDrivers]=useState([]), [periods,setPeriods]=useState([])
  const [preview,setPreview]=useState([]), [msg,setMsg]=useState(''), [loading,setLoading]=useState(false)
  const empty={nama:'',nrp:'',email:'',password:'',role:'Driver',site_id:work.site_id||'',status:'Aktif'}
  const [form,setForm]=useState(empty)

  useEffect(()=>{load()},[work.id])
  async function load(){
    try{
      setMsg('')
      const [appRes, siteRes, profRes, accessRes, qRes, vRes, dRes, pRes]=await Promise.all([
        supabase.from('applications').select('*').eq('app_code','drd_driver').limit(1),
        supabase.from('sites').select('*').order('site_code'),
        supabase.from('users_profile').select('*').order('created_at',{ascending:false}).limit(500),
        supabase.from('user_app_access').select('*,users_profile(nama,nrp,email,status),applications(app_code,app_name),sites(site_code,site_name)').order('created_at',{ascending:false}),
        supabase.from('drd_questions').select('id,category,status').eq('status','Aktif'),
        supabase.from('drd_induction_videos').select('id,status').eq('status','Aktif'),
        supabase.from('drivers').select('id,site_id,status,end_masa_dinas,sites(site_code)').eq('status','Aktif'),
        supabase.from('drd_induction_periods').select('id,site_id,status,drivers(site_id)').order('created_at',{ascending:false})
      ])
      const appRows=appRes.data||[]
      const siteRows=siteRes.data||[]
      const accessRows=(accessRes.data||[]).filter(r=>(r.applications?.app_code||'')==='drd_driver' && (isAdmin(work)||r.site_id===work.site_id||!r.site_id))
      const driverRows=(dRes.data||[]).filter(r=>isAdmin(work)||r.site_id===work.site_id)
      const periodRows=(pRes.data||[]).filter(r=>isAdmin(work)||r.drivers?.site_id===work.site_id||r.site_id===work.site_id)
      setApps(appRows); setSites(siteRows); setProfiles(profRes.data||[]); setAccess(accessRows); setQuestions(qRes.data||[]); setVideos(vRes.data||[]); setDrivers(driverRows); setPeriods(periodRows)
      setForm(f=>({...f,site_id:f.site_id||work.site_id||''}))
    }catch(e){
      console.warn('[ADMIN PANEL DRD LOAD]', e)
      setMsg(e.message||String(e))
    }
  }

  async function upsertProfile(row){
    const email=normEmail(row.email)
    const existing=profiles.find(p=>normEmail(p.email)===email)
    const payload={nama:clean(row.nama)||email,nrp:clean(row.nrp)||null,email,status:row.status||'Aktif'}
    if(existing?.id){
      const {data,error}=await supabase.from('users_profile').update(payload).eq('id',existing.id).select('*').single()
      if(error) throw error
      return data
    }
    const {data,error}=await supabase.from('users_profile').insert(payload).select('*').single()
    if(error) throw error
    return data
  }

  async function ensureAccess(profileRow, row){
    const app=apps[0]
    if(!app?.id) throw new Error('Aplikasi DRD Driver belum ditemukan di tabel applications.')
    const role=clean(row.role)||'Driver'
    const siteId=row.site_id||null
    const exists=access.find(a=>String(a.user_id)===String(profileRow.id) && String(a.app_id)===String(app.id) && String(a.role)===String(role) && String(a.site_id||'')===String(siteId||'') && a.status==='Aktif')
    if(exists) return exists
    const {data,error}=await supabase.from('user_app_access').insert({user_id:profileRow.id,app_id:app.id,role,site_id:siteId,status:'Aktif'}).select('*').single()
    if(error) throw error
    return data
  }

  async function saveMapping(e){
    e.preventDefault(); setLoading(true); setMsg('')
    try{
      if(form.email && form.password) await createAuthUser({email:form.email,password:form.password,nama:form.nama,nrp:form.nrp,app_id:form.app_id||work.app_id||work.applications?.id,site_id:form.site_id,role:form.role})
      const profileRow=await upsertProfile(form)
      await ensureAccess(profileRow, form)
      setMsg('User dan mapping akses DRD berhasil disimpan.')
      setForm({...empty,site_id:work.site_id||''})
      await load()
    }catch(e){ setMsg(e.message||String(e)) }
    setLoading(false)
  }

  async function previewFile(file){
    if(!file) return
    const rows=await readExcel(file)
    const mapped=rows.map((r,i)=>{
      const siteCode=clean(r.site_code).toUpperCase()
      const site=siteCode ? sites.find(s=>String(s.site_code).toUpperCase()===siteCode) : null
      const role=clean(r.role)||'Driver'
      let error=''
      if(!normEmail(r.email)) error='email wajib'
      else if(!['Platform Admin','App Admin','GL','Driver','Atasan Site'].includes(role)) error='role tidak valid'
      else if(role!=='Platform Admin' && role!=='App Admin' && !site) error='site_code wajib/valid untuk role ini'
      return {row:i+2,nama:clean(r.nama),nrp:clean(r.nrp),email:normEmail(r.email),password:clean(r.password),role,site_code:site?.site_code||siteCode,status:clean(r.status)||'Aktif',site_id:site?.id||null,error}
    })
    setPreview(mapped)
  }

  async function importRows(){
    setLoading(true); setMsg('')
    try{
      const valid=preview.filter(r=>!r.error)
      if(valid.length!==preview.length) throw new Error('Masih ada baris invalid pada preview.')
      for(const [i,r] of valid.entries()){
        if(r.email && r.password) await createAuthUser({email:r.email,password:r.password,nama:r.nama,nrp:r.nrp,app_id:work.app_id||work.applications?.id,site_id:r.site_id,role:r.role})
        const profileRow=await upsertProfile(r)
        await ensureAccess(profileRow, r)
      }
      setPreview([])
      setMsg(`Import mapping berhasil: ${valid.length} baris.`)
      await load()
    }catch(e){ setMsg(e.message||String(e)) }
    setLoading(false)
  }

  async function nonaktif(row){
    if(!confirm('Nonaktifkan mapping akses ini?')) return
    const {error}=await supabase.from('user_app_access').update({status:'Nonaktif'}).eq('id',row.id)
    setMsg(error?error.message:'Mapping akses dinonaktifkan.')
    load()
  }

  const activeDrivers=drivers.filter(d=>d.status==='Aktif')
  const expiredDrivers=activeDrivers.filter(d=>isExpired(d.end_masa_dinas)).length
  const openPeriods=periods.filter(p=>p.status==='Open').length
  const qDrd=questions.filter(q=>q.category==='DRD').length
  const qInduksi=questions.filter(q=>q.category==='Induksi Driver').length
  const accessRows=access.map(a=>({nama:a.users_profile?.nama||'-',email:a.users_profile?.email||'-',nrp:a.users_profile?.nrp||'-',aplikasi:a.applications?.app_name||'DRD Driver',role:a.role,site:a.sites?.site_code||'All Site',status:a.status}))

  return <div className="stack admin-panel-aligned">
    <div className="kpi-grid">
      <Kpi title="Driver Aktif" value={activeDrivers.length} icon={<Truck/>}/>
      <Kpi title="Soal DRD" value={qDrd} icon={<FileQuestion/>}/>
      <Kpi title="Soal Induksi" value={qInduksi} icon={<ClipboardCheck/>}/>
      <Kpi title="Video Aktif" value={videos.length} icon={<Video/>}/>
      <Kpi title="Masa Dinas Habis" value={expiredDrivers} icon={<CalendarCheck/>}/>
      <Kpi title="Open Periode" value={openPeriods} icon={<AlertTriangle/>}/>
    </div>

    <Panel title="Buat User & Mapping Akses DRD" desc="Struktur Admin Panel disamakan dengan aplikasi lain: kelola akun, role, site, dan mapping akses langsung dari panel ini.">
      <form className="form-grid" onSubmit={saveMapping}>
        <label>Nama<input required value={form.nama} onChange={e=>setForm({...form,nama:e.target.value})}/></label>
        <label>NRP<input value={form.nrp} onChange={e=>setForm({...form,nrp:e.target.value})}/></label>
        <label>Email<input required type="email" value={form.email} onChange={e=>setForm({...form,email:e.target.value})}/></label>
        <label>Password<input type="password" name="driver_password_no_autofill" autoComplete="new-password" value={form.password} onChange={e=>setForm({...form,password:e.target.value})} placeholder="Isi jika ingin membuat/reset akun Auth"/></label>
        <label>Role<select value={form.role} onChange={e=>setForm({...form,role:e.target.value})}><option>Platform Admin</option><option>App Admin</option><option>GL</option><option>Atasan Site</option><option>Driver</option></select></label>
        <label>Site<select value={form.site_id} onChange={e=>setForm({...form,site_id:e.target.value})}><option value="">All Site / Head Office</option>{sites.map(s=><option key={s.id} value={s.id}>{s.site_code} - {s.site_name}</option>)}</select></label>
        <button disabled={loading}>{loading?'Memproses...':'Simpan User & Mapping'}</button>
      </form>
      {msg&&<p className="message">{msg}</p>}
    </Panel>

    <Panel title="Upload Bulk Mapping Akses DRD" desc="Gunakan template ini untuk membuat banyak user dan mapping akses DRD sekaligus. Kolom site_code wajib untuk role site/driver.">
      <div className="import-actions">
        <button className="secondary" onClick={()=>templateXlsx('template-mapping-akses-drd.xlsx',[{nama:'Driver Baya 1',nrp:'Baya001',email:'driverbaya1@gmail.com',password:'password123',role:'Driver',site_code:'BAYA',status:'Aktif'},{nama:'GL BAYA',nrp:'GL-001',email:'glbaya@company.co.id',password:'password123',role:'GL',site_code:'BAYA',status:'Aktif'}])}><Download size={16}/> Download Template</button>
        <label className="upload-line"><Upload size={16}/> Upload Excel<input type="file" accept=".xlsx,.xls" onChange={e=>previewFile(e.target.files?.[0])}/></label>
        {preview.length>0&&<button disabled={loading} onClick={importRows}>Submit Import Valid</button>}
      </div>
      {preview.length>0&&<div className="upload-preview"><DataTable rows={preview}/></div>}
    </Panel>

    <Panel title="Row Data Mapping Akses DRD" desc="Daftar akses user khusus aplikasi DRD Driver. Nonaktifkan mapping bila akses sudah tidak dipakai." action={<button onClick={()=>exportXlsx('mapping-akses-drd.xlsx',accessRows)}><Download size={16}/> Export</button>}>
      <DataTable rows={accessRows} actions={(r,i)=><button className="danger small" onClick={()=>nonaktif(access[i])}>Nonaktif</button>}/>
    </Panel>

    <Panel title="Shortcut Operasional DRD" desc="Panel ini mengikuti pola Admin Panel aplikasi lain: administrasi user/mapping tetap tersedia, sedangkan fitur operasional dibuka melalui menu sidebar.">
      <div className="cards-grid">
        <div className="card"><div className="kpi-icon"><FileQuestion/></div><h3>Bank Soal</h3><p className="muted">Soal kategori DRD dan Induksi Driver aktif langsung dipakai tanpa paket soal.</p></div>
        <div className="card"><div className="kpi-icon"><Video/></div><h3>Video Induksi</h3><p className="muted">Kelola video yang wajib ditonton driver sebelum mengerjakan soal induksi.</p></div>
        <div className="card"><div className="kpi-icon"><Truck/></div><h3>Master Driver</h3><p className="muted">Kelola driver, vendor, masa dinas, dan akun driver.</p></div>
        <div className="card"><div className="kpi-icon"><CalendarCheck/></div><h3>Periode Cuti</h3><p className="muted">Isi periode cuti dan tanggal onsite saat masa dinas habis.</p></div>
      </div>
    </Panel>
  </div>
}

export default App
