import { useState, useEffect, useCallback, useRef } from 'react'
import { supabase } from './supabase.js'
import * as XLSX from 'xlsx'

const TALLE_ORDER = ['2','4','6','8','10','12','14','Único','S','M','L','XL','XXL','XXXL']
const TALLES_ADULTO = ['S','M','L','XL','XXL','XXXL','Único']
const TALLES_NINO   = ['2','4','6','8','10','12','14']
const RECEPTORES = ['1° División','3° División','Juveniles','Captación','Femenino','Juveniles Femenino','Fútbol Sala Masculino','Fútbol Sala Femenino','Basket','Deportes Anexos','Funcionarios','Protocolo']
const CATEGORIAS = ['Entrenamiento','Juego','Casual']
const OCUPACIONES = ['3° División','Juveniles','Juveniles Femenino','Captacion']
const DIVISIONES = ['Sub 19','Sub 17','Sub 16','Sub 15','Sub 14','Captacion']
const CARGOS_REG = ['Coordinación','Director Técnico','Ayudante Técnico','Videoanalista','Preparador Físico','Entrenador de Arqueros','Doctor/a','Kinesiólogo/a','Utilero']
const ESTANTES = ['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20']
const ALTURAS = ['A','B','C','D','E','O']

const DEFAULT_USERS = [{ username:'compras', password:'peniarol1891', role:'admin', displayName:'Compras Peñarol', status:'aprobado' }]
const EMPTY_DB = { articles:[], deliveries:[], movimientos:[], nextId:1, nextDel:1, nextMov:1, users: DEFAULT_USERS, camisetasUtileria:[] }
const COMPETICIONES = ['CAMPEONATO URUGUAYO','CONMEBOL','COPA LIBERTADORES FEMENINA','COPA LIBERTADORES FÚTBOL SALA','COPA INTERCONTINENTAL SUB 20']

const USERS_KEY = 'dep_usuarios_v1'
const SESSION_KEY = 'dep_session'


async function loadFromSupabase() {
  const [{ data, error }, { data: usersRow }, { data: utiRow }] = await Promise.all([
    supabase.from('deposito_state').select('*').eq('id', 1).single(),
    supabase.from('deposito_state').select('deliveries').eq('id', 2).single(),
    supabase.from('deposito_state').select('articles').eq('id', 3).single(),
  ])
  if (error || !data) return null
  let users = (usersRow?.deliveries?.length > 0 && usersRow.deliveries[0]?.username)
    ? usersRow.deliveries
    : null
  if (!users) {
    try {
      const raw = JSON.parse(localStorage.getItem(USERS_KEY)) || []
      users = raw.length > 0
        ? raw.map(u => ({ displayName: u.username, ...u, role: u.role || 'admin' }))
        : DEFAULT_USERS
    } catch { users = DEFAULT_USERS }
  }
  return {
    articles: (data.articles || []).map(a => ({
      ...a, sizes: (a.sizes || []).map(s => ({ talle: s.talle, qty: Number(s.qty)||0, min: Number(s.min)||0 }))
    })),
    deliveries: data.deliveries || [],
    movimientos: data.movimientos || [],
    nextId: data.next_id || 1,
    nextDel: data.next_del || 1,
    nextMov: data.next_mov || 1,
    users,
    camisetasUtileria: utiRow?.articles || [],
  }
}

async function saveToSupabase(db) {
  const [r1, r2, r3] = await Promise.all([
    supabase.from('deposito_state').upsert({
      id: 1,
      articles: db.articles,
      deliveries: db.deliveries,
      movimientos: db.movimientos,
      next_id: db.nextId,
      next_del: db.nextDel,
      next_mov: db.nextMov,
      updated_at: new Date().toISOString(),
    }),
    supabase.from('deposito_state').upsert({
      id: 2,
      articles: [],
      deliveries: db.users,
      movimientos: [],
      next_id: 0,
      next_del: 0,
      next_mov: 0,
      updated_at: new Date().toISOString(),
    }),
    supabase.from('deposito_state').upsert({
      id: 3,
      articles: db.camisetasUtileria || [],
      deliveries: [],
      movimientos: [],
      next_id: 0,
      next_del: 0,
      next_mov: 0,
      updated_at: new Date().toISOString(),
    }),
  ])
  return !r1.error && !r2.error && !r3.error
}

function fmt(n) { return Number(n).toLocaleString('es-UY') }
function total(a) { return a.sizes.reduce((s,x) => s + x.qty, 0) }
function ini(name) {
  const p = (name||'').trim().split(/\s+/)
  return ((p[0]||'')[0]||'').toUpperCase() + ((p[1]||'')[0]||'').toUpperCase()
}
function sizesLabel(a) {
  const t = a.sizes.map(s => s.talle)
  if(t.length === 1 && t[0] === 'Único') return 'Único'
  const present = ['S','M','L','XL','XXL'].filter(o => t.includes(o))
  if(present.length === 0) return t.join(', ')
  if(present.length === 1) return present[0]
  return present[0] + '–' + present[present.length - 1]
}
function today() {
  const d = new Date()
  return String(d.getDate()).padStart(2,'0') + '/' + String(d.getMonth()+1).padStart(2,'0') + '/' + d.getFullYear()
}
function isLow(a) { return a.sizes.some(s => (s.min||0) > 0 && s.qty <= (s.min||0)) }


export default function App() {
  const [db, setDb] = useState(EMPTY_DB)
  const [loading, setLoading] = useState(true)
  const [view, setView] = useState('panel')
  const [selectedId, setSelectedId] = useState(null)
  const [selectedCode, setSelectedCode] = useState(null)
  const [search, setSearch] = useState('')
  const [cat, setCat] = useState('Todas')
  const [modal, setModal] = useState(null)
  const [confirm, setConfirm] = useState(null)
  const [editing, setEditing] = useState(null)
  const [movFilter, setMovFilter] = useState('Todos')
  const [delFilterReceptor, setDelFilterReceptor] = useState('')
  const [delFilterPersona, setDelFilterPersona] = useState('')
  const [selectedDeliveryId, setSelectedDeliveryId] = useState(null)
  const [selectedReceptor, setSelectedReceptor] = useState(null)
  const [utiFilter, setUtiFilter] = useState('')
  const [utiForm, setUtiForm] = useState({ competicion:'', numero:'', jugador:'', talle:'S', modelo:'', estampado:'', parches:'', detalle:'', id:null })
  const [utiModal, setUtiModal] = useState(false)
  const [toast, setToast] = useState('')
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [session, setSession] = useState(() => localStorage.getItem(SESSION_KEY) || null)
  const [loginView, setLoginView] = useState('login')
  const [loginForm, setLoginForm] = useState({ user:'', pass:'', err:'' })
  const [regForm, setRegForm] = useState({ displayName:'', email:'', telefono:'', cargo:'', categoria:'', division:'', pass:'', pass2:'', err:'' })
  const [forgotForm, setForgotForm] = useState({ email:'', newPass:'', newPass2:'', step:1, err:'' })
  const [changePassForm, setChangePassForm] = useState({ current:'', newPass:'', newPass2:'', err:'' })
  const [userMgmt, setUserMgmt] = useState({ list:[], newUser:'', newPass:'', err:'' })
  const toastTimer = useRef(null)
  const saveTimer = useRef(null)
  const saveEnabled = useRef(false)
  const dbRef = useRef(db)

  // delivery/devolución form
  const [nd, setNd] = useState({ mode:'entrega', persona:'', receptor:'', cCode:'', cSearch:'', cUbic:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' })
  // new article form
  const [na, setNa] = useState({ code:'', name:'', cat:'Entrenamiento', tipo:'adulto', precio:'', tallesArr:[], tallesMins:{}, tallesQty:{}, estante:'1', altura:'A' })
  // reponer form
  const [rep, setRep] = useState({ qtys:{} })
  // ajuste form
  const [aj, setAj] = useState({ talle:'', cantidad:'' })
  // mover form
  const [mv, setMv] = useState({ qtys:{}, estante:'1', altura:'A' })

  // Keep dbRef current so visibilitychange flush always sees latest state
  useEffect(() => { dbRef.current = db }, [db])

  // Load from Supabase on mount (filter out articles with no stock)
  useEffect(() => {
    loadFromSupabase().then(data => {
      if (data) {
        setDb({...data, articles: data.articles.filter(a => total(a) > 0)})
        saveEnabled.current = true
      }
      setLoading(false)
    })
  }, [])

  // Save to Supabase whenever data changes (debounced 800ms)
  useEffect(() => {
    if (loading || !saveEnabled.current) return
    clearTimeout(saveTimer.current)
    saveTimer.current = setTimeout(async () => {
      const ok = await saveToSupabase(db)
      if (!ok) showToast('Error al guardar. Verificá la conexión.')
    }, 800)
  }, [db, loading])

  // Flush any pending save immediately when tab is hidden (mobile: switch app / close tab)
  useEffect(() => {
    const flush = () => {
      if (!saveEnabled.current) return
      clearTimeout(saveTimer.current)
      saveToSupabase(dbRef.current)
    }
    document.addEventListener('visibilitychange', flush)
    return () => document.removeEventListener('visibilitychange', flush)
  }, [])

  // Redirect to inventario if selected article code no longer exists
  useEffect(() => {
    if (view === 'detalle' && selectedCode !== null && !db.articles.find(a => a.code === selectedCode)) {
      setView('inventario')
      setSelectedCode(null)
    }
  }, [db.articles, view, selectedCode])

  const saveUsers = (list) => {
    setDb(prev => ({ ...prev, users: list }))
  }
  const doLogin = () => {
    const found = db.users.find(u => u.username.toLowerCase() === loginForm.user.toLowerCase() && u.password === loginForm.pass)
    if(!found) { setLoginForm(p => ({...p, err:'Usuario o contraseña incorrectos.'})); return }
    if(found.status === 'pendiente') { setLoginForm(p => ({...p, err:'Tu cuenta está pendiente de aprobación por el administrador.'})); return }
    localStorage.setItem(SESSION_KEY, found.username); setSession(found.username); setLoginForm({user:'',pass:'',err:''})
  }
  const doRegister = () => {
    const { displayName, email, telefono, cargo, categoria, division, pass, pass2 } = regForm
    if(!displayName.trim()) { setRegForm(p=>({...p,err:'Ingresá tu nombre completo.'})); return }
    if(!email.trim() || !email.includes('@')) { setRegForm(p=>({...p,err:'Ingresá un correo electrónico válido.'})); return }
    if(!telefono.trim()) { setRegForm(p=>({...p,err:'Ingresá tu teléfono.'})); return }
    if(!cargo) { setRegForm(p=>({...p,err:'Seleccioná tu cargo.'})); return }
    if(!categoria) { setRegForm(p=>({...p,err:'Seleccioná tu sector.'})); return }
    if(!division) { setRegForm(p=>({...p,err:'Seleccioná tu división.'})); return }
    if(!pass || pass.length < 6) { setRegForm(p=>({...p,err:'La contraseña debe tener al menos 6 caracteres.'})); return }
    if(pass !== pass2) { setRegForm(p=>({...p,err:'Las contraseñas no coinciden.'})); return }
    const username = email.trim().toLowerCase()
    if(db.users.find(u => u.username.toLowerCase() === username)) { setRegForm(p=>({...p,err:'Ya existe una cuenta con ese correo.'})); return }
    const newUser = { username, password:pass, role:'receptor', displayName:displayName.trim(), email:username, telefono:telefono.trim(), cargo, categoria, division, status:'pendiente' }
    saveUsers([...db.users, newUser])
    setLoginView('registered')
    setRegForm({ displayName:'', email:'', telefono:'', cargo:'', categoria:'', division:'', pass:'', pass2:'', err:'' })
  }
  const doLogout = () => { localStorage.removeItem(SESSION_KEY); setSession(null) }
  const doForgotStep1 = () => {
    const email = forgotForm.email.trim().toLowerCase()
    if(!email) { setForgotForm(p=>({...p,err:'Ingresá tu correo.'})); return }
    if(!db.users.find(u => u.username.toLowerCase() === email)) { setForgotForm(p=>({...p,err:'No existe una cuenta con ese correo.'})); return }
    setForgotForm(p=>({...p,step:2,err:''}))
  }
  const doForgotStep2 = () => {
    const { email, newPass, newPass2 } = forgotForm
    if(!newPass || newPass.length < 6) { setForgotForm(p=>({...p,err:'La contraseña debe tener al menos 6 caracteres.'})); return }
    if(newPass !== newPass2) { setForgotForm(p=>({...p,err:'Las contraseñas no coinciden.'})); return }
    saveUsers(db.users.map(u => u.username.toLowerCase()===email.trim().toLowerCase() ? {...u,password:newPass} : u))
    setForgotForm({ email:'', newPass:'', newPass2:'', step:1, err:'' })
    setLoginView('login')
    setLoginForm(p=>({...p,err:''}))
  }
  const doChangePass = () => {
    const { current, newPass, newPass2 } = changePassForm
    const me = db.users.find(u => u.username === session)
    if(!me || me.password !== current) { setChangePassForm(p=>({...p,err:'La contraseña actual es incorrecta.'})); return }
    if(!newPass || newPass.length < 6) { setChangePassForm(p=>({...p,err:'La nueva contraseña debe tener al menos 6 caracteres.'})); return }
    if(newPass !== newPass2) { setChangePassForm(p=>({...p,err:'Las contraseñas no coinciden.'})); return }
    saveUsers(db.users.map(u => u.username===session ? {...u,password:newPass} : u))
    setChangePassForm({ current:'', newPass:'', newPass2:'', err:'' })
    closeModal()
    showToast('Contraseña actualizada correctamente.')
  }
  const openUserMgmt = () => { setUserMgmt({ list:db.users, newUser:'', newPass:'', newDisplayName:'', newRole:'receptor', err:'' }); setModal('usuarios') }
  const addUser = () => {
    const u = userMgmt.newUser.trim(); const p = userMgmt.newPass.trim()
    if(!u || !p) { setUserMgmt(x=>({...x,err:'Completá usuario y contraseña.'})); return }
    if(userMgmt.list.find(x => x.username.toLowerCase()===u.toLowerCase())) { setUserMgmt(x=>({...x,err:'Ese usuario ya existe.'})); return }
    const displayName = userMgmt.newDisplayName.trim() || u
    const list = [...userMgmt.list, {username:u, password:p, role:userMgmt.newRole||'receptor', displayName, status:'aprobado'}]
    saveUsers(list)
    setUserMgmt(x=>({...x,list,newUser:'',newPass:'',newDisplayName:'',newRole:'receptor',err:''}))
  }
  const deleteUser = (username) => {
    if(username === session) { showToast('No podés eliminar tu propio usuario.'); return }
    const list = userMgmt.list.filter(u => u.username !== username)
    saveUsers(list)
    setUserMgmt(x=>({...x,list}))
  }
  const approveUser = (username) => {
    const list = db.users.map(u => u.username===username ? {...u, status:'aprobado'} : u)
    saveUsers(list)
    setUserMgmt(x=>({...x, list}))
    showToast('Usuario aprobado.')
  }
  const rejectUser = (username) => {
    const list = db.users.filter(u => u.username !== username)
    saveUsers(list)
    setUserMgmt(x=>({...x, list}))
    showToast('Usuario rechazado y eliminado.')
  }

  const showToast = useCallback((msg) => {
    setToast(msg)
    clearTimeout(toastTimer.current)
    toastTimer.current = setTimeout(() => setToast(''), 2600)
  }, [])

  const closeModal = () => setModal(null)
  const byCode = (code) => db.articles.find(a => a.code === code)
  const curCode = () => selectedCode || ''

  const goView = (v) => { setView(v); setSearch(''); setSidebarOpen(false) }
  const openDetail = (code) => { setSelectedCode(code); setView('detalle'); setSidebarOpen(false) }

  // ---- Entregas / Devoluciones ----
  const openEntrega = () => { setNd({ mode:'entrega', persona:'', receptor:'', cCode:'', cSearch:'', cUbic:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' }); setModal('entrega') }
  const openDevolucion = () => { setNd({ mode:'devolucion', persona:'', receptor:'', cCode:'', cSearch:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' }); setModal('entrega') }
  const openEntregaFromDetail = () => { const a = byCode(selectedCode); setNd({ mode:'entrega', persona:'', receptor:'', cCode:a?a.code:'', cSearch:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' }); setModal('entrega') }
  const openDevolucionFromDetail = () => { const a = byCode(selectedCode); setNd({ mode:'devolucion', persona:'', receptor:'', cCode:a?a.code:'', cSearch:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' }); setModal('entrega') }

  const ndAddLine = () => {
    const qty = parseInt(nd.cQty, 10)
    if(!nd.cCode || !nd.cTalle || !qty || qty <= 0) { showToast('Completá artículo, talle y cantidad.'); return }
    const ndUbicsAll = [...new Set(db.articles.filter(a => a.code === nd.cCode).map(a => a.ubic).filter(Boolean))]
    if(ndUbicsAll.length > 1 && !nd.cUbic) { showToast('Seleccioná la ubicación primero.'); return }
    const ubicToUse = nd.cUbic || (ndUbicsAll[0] || '')
    if(nd.mode !== 'devolucion') {
      const srcArt = db.articles.find(a => a.code === nd.cCode && (!ubicToUse || a.ubic === ubicToUse))
      const avail = srcArt ? (srcArt.sizes.find(x => x.talle === nd.cTalle)?.qty || 0) : 0
      const already = nd.lines.filter(l => l.code === nd.cCode && l.talle === nd.cTalle && l.ubic === ubicToUse).reduce((s,l) => s+l.qty, 0)
      if(avail === 0 || qty + already > avail) { showToast('Stock insuficiente ('+(avail-already)+' disp.).'); return }
    }
    const artName = db.articles.find(a => a.code === nd.cCode)?.name || nd.cCode
    setNd(p => ({...p, lines:[...p.lines,{code:nd.cCode,talle:nd.cTalle,qty,ubic:ubicToUse,name:artName}], cCode:'', cSearch:'', cUbic:'', cTalle:'', cQty:''}))
  }

  const ndConfirm = () => {
    const esDev = nd.mode === 'devolucion'
    if(!nd.persona.trim()) { showToast('Ingresá el nombre del integrante.'); return }
    if(!nd.receptor) { showToast('Elegí un grupo / plantel.'); return }
    if(nd.lines.length === 0) { showToast('Agregá al menos un artículo.'); return }
    setDb(s => {
      const articles = s.articles.map(a => ({...a, sizes: a.sizes.map(z => ({...z}))}))
      const movimientos = [...s.movimientos]
      let mid = s.nextMov
      const fecha = today()
      nd.lines.forEach(l => {
        // Find the entry at the specific location (ubic), fallback to first with the talle
        const a = l.ubic
          ? articles.find(x => x.code === l.code && x.ubic === l.ubic)
          : articles.find(x => x.code === l.code && x.sizes.some(sz => sz.talle === l.talle))
        const z = a && a.sizes.find(x => x.talle === l.talle)
        if(z) z.qty = esDev ? z.qty + l.qty : Math.max(0, z.qty - l.qty)
        if(esDev) {
          movimientos.unshift({id:mid++, code:l.code, name:a?.name||l.code, tipo:'entrada', fecha, talle:l.talle, qty:l.qty, detalle:'Devolución de '+nd.persona+' ('+nd.receptor+')'})
        } else {
          movimientos.unshift({id:mid++, code:l.code, name:a?.name||l.code, tipo:'salida', fecha, talle:l.talle, qty:l.qty, detalle:'Entrega a '+nd.persona+' ('+nd.receptor+')', delId:s.nextDel})
        }
      })
      const activeArticles = articles.filter(a => total(a) > 0)
      if(esDev) return { ...s, articles:activeArticles, movimientos, modal:null, nextMov:mid }
      const toUser = nd.toUser || null
      const status = toUser ? 'pendiente' : 'aceptado'
      const confirmedAt = toUser ? null : fecha
      const deliveries = [{id:s.nextDel, fecha, persona:nd.persona.trim(), receptor:nd.receptor, paga:nd.receptor==='Protocolo'?nd.paga:null, monto:nd.receptor==='Protocolo'&&nd.paga==='si'?ndMonto:null, lines:[...nd.lines], toUser, status, confirmedAt}, ...s.deliveries]
      return { ...s, articles:activeArticles, movimientos, deliveries, nextDel:s.nextDel+1, nextMov:mid }
    })
    // Enviar email de notificación si la entrega va a un usuario específico
    if (nd.toUser && nd.mode !== 'devolucion') {
      const recipient = db.users.find(u => u.username === nd.toUser)
      if (recipient?.email) {
        fetch('/api/send-email', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            to: recipient.email,
            displayName: recipient.displayName || recipient.username,
            lines: nd.lines,
            delId: db.nextDel,
          })
        })
        .then(r => r.json())
        .then(d => { if (d.ok) { showToast('Email de notificación enviado a ' + recipient.email) } else { showToast('Error al enviar email: ' + (d.error || 'error desconocido')) } })
        .catch(() => showToast('No se pudo conectar con el servidor de email.'))
      }
    }
    setModal(null)
    setView(nd.mode === 'devolucion' ? 'inventario' : 'entregas')
    showToast(nd.mode === 'devolucion' ? 'Devolución registrada y stock actualizado.' : 'Entrega registrada y stock actualizado.')
  }

  // ---- Nuevo artículo ----
  const openArticulo = () => { setNa({ code:'', name:'', cat:'Entrenamiento', tipo:'adulto', precio:'', tallesArr:[], tallesMins:{}, tallesQty:{}, estante:'1', altura:'A' }); setModal('articulo') }

  const naToggleTalle = (t) => {
    setNa(p => {
      const arr = p.tallesArr, mins = {...p.tallesMins}, qtys = {...p.tallesQty}
      if(t === 'Único') {
        if(arr.includes('Único')) return {...p, tallesArr:[], tallesMins:{}, tallesQty:{}}
        return {...p, tallesArr:['Único'], tallesMins:{}, tallesQty:{}}
      }
      const without = arr.filter(x => x !== 'Único')
      if(without.includes(t)) { delete mins[t]; delete qtys[t]; return {...p, tallesArr:without.filter(x=>x!==t), tallesMins:mins, tallesQty:qtys} }
      return {...p, tallesArr:[...without,t]}
    })
  }

  const naConfirm = () => {
    const { code, name, cat:ncat, tallesArr, tallesMins, tallesQty, estante, altura, precio } = na
    if(!code || !name) { showToast('Completá código y nombre.'); return }
    if(tallesArr.length === 0) { showToast('Seleccioná al menos un talle.'); return }
    const ubic = (estante||'1') + (altura||'A')
    const sizes = tallesArr.map(t => ({talle:t, qty:tallesQty[t]||0, min:tallesMins[t]||0}))
    const precioNum = parseFloat(precio)||0
    setDb(s => ({ ...s, articles:[{id:s.nextId, code, name, cat:ncat, ubic, precio:precioNum, sizes}, ...s.articles], nextId:s.nextId+1 }))
    setModal(null); setView('inventario')
    showToast('Artículo «'+name+'» creado.')
  }

  // ---- Reponer ----
  const openReponer = () => { setRep({ qtys:{} }); setModal('reponer') }
  const repConfirm = () => {
    const entries = selA.sizes
      .map(s => ({ talle:s.talle, q:parseInt(rep.qtys[s.talle]||0, 10) }))
      .filter(e => e.q > 0)
    if(entries.length === 0) { showToast('Ingresá al menos una cantidad.'); return }
    const fecha = today()
    setDb(s => {
      let nextMov = s.nextMov
      const artName = selA.name
      const code = selA.code
      const articles = s.articles.map(a => {
        if(a.id !== selectedId) return a
        return {...a, sizes: a.sizes.map(z => { const e=entries.find(e=>e.talle===z.talle); return e ? {...z, qty:z.qty+e.q} : z })}
      })
      const newMovs = entries.map(e => ({id:nextMov++, code, name:artName, tipo:'entrada', fecha, talle:e.talle, qty:e.q, detalle:'Ingreso de stock'}))
      return { ...s, articles, movimientos:[...newMovs,...s.movimientos], nextMov }
    })
    setModal(null)
    const tot = entries.reduce((s,e)=>s+e.q, 0)
    showToast('Entrada registrada: +'+tot+' u. en '+entries.length+' talle'+(entries.length>1?'s':'')+'.')
  }

  // ---- Ajuste ----
  const openAjuste = () => { setAj({ talle:'', cantidad:'' }); setModal('ajuste') }
  const ajConfirm = () => {
    if(!aj.talle || aj.cantidad === '') { showToast('Elegí talle e ingresá la cantidad contada.'); return }
    const q = parseInt(aj.cantidad, 10)
    if(isNaN(q) || q < 0) { showToast('Cantidad inválida.'); return }
    const code = selA.code; const fecha = today()
    const z0 = selA.sizes.find(z => z.talle === aj.talle); const cur = z0 ? z0.qty : 0
    const delta = q - cur
    if(delta === 0) { showToast('Sin cambios: el stock ya es '+q+'.'); setModal(null); return }
    setDb(s => {
      const artName = selA.name
      const articles = s.articles.map(a => { if(a.id!==selectedId) return a; return {...a, sizes:a.sizes.map(z => z.talle===aj.talle?{...z,qty:Math.max(0,q)}:z)} })
      const activeArticles = articles.filter(a => total(a) > 0)
      const movimientos = [{id:s.nextMov, code, name:artName, tipo:(delta>0?'entrada':'salida'), fecha, talle:aj.talle, qty:Math.abs(delta), detalle:'Ajuste por recuento (de '+cur+' a '+q+')'}, ...s.movimientos]
      return { ...s, articles:activeArticles, movimientos, nextMov:s.nextMov+1 }
    })
    setModal(null)
    showToast('Stock ajustado: '+aj.talle+' = '+q+' ('+(delta>0?'+':'')+delta+').')
  }

  // ---- Mover talle ----
  const openMover = () => { setMv({ qtys:{}, estante:'1', altura:'A' }); setModal('mover') }

  const exportExcel = () => {
    const sorted = [...articles].sort((a, b) => {
      const pu = u => { if(!u||u==='—') return {n:Infinity,l:''}; const m=u.match(/^(\d+)(.*)/); return m?{n:parseInt(m[1],10),l:m[2]}:{n:Infinity,l:u} }
      const ua=pu(a.ubic), ub=pu(b.ubic)
      return ua.n!==ub.n ? ua.n-ub.n : ua.l.localeCompare(ub.l)
    })
    const rows = sorted.map(a => {
      const row = { UBICACIÓN: a.ubic||'—', CÓDIGO: a.code, ARTÍCULO: a.name, CATEGORÍA: a.cat, PRECIO: a.precio||0 }
      TALLE_ORDER.forEach(t => { const sz=a.sizes.find(s=>s.talle===t); row[t]=sz?sz.qty:'' })
      row['TOTAL'] = a.sizes.reduce((s,z)=>s+z.qty,0)
      return row
    })
    const headers = ['UBICACIÓN','CÓDIGO','ARTÍCULO','CATEGORÍA','PRECIO',...TALLE_ORDER,'TOTAL']
    const ws = XLSX.utils.json_to_sheet(rows, { header: headers })
    ws['!cols'] = headers.map((h,i) => ({ wch: i===2?32 : i===4?10 : i<5?14 : 7 }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Stock por Ubicación')
    XLSX.writeFile(wb, 'stock-deposito-peniarol.xlsx')
  }
  const mvConfirm = () => {
    const toMove = Object.entries(mv.qtys).map(([t,q]) => [t, parseInt(q,10)||0]).filter(([,q]) => q > 0)
    if(toMove.length === 0) { showToast('Ingresá la cantidad a mover en al menos un talle.'); return }
    const newUbic = mv.estante + mv.altura
    if(newUbic === selA.ubic) { showToast('La ubicación destino es la misma que la actual.'); return }
    for(const [t, q] of toMove) {
      const src = selA.sizes.find(sz => sz.talle === t)
      if(!src || q > src.qty) { showToast('Stock insuficiente para talle ' + t + '.'); return }
    }
    setDb(prev => {
      const code = curCode()
      let arts = [...prev.articles]
      let nextId = prev.nextId
      // Reducir stock en origen
      arts = arts.map(a => {
        if(a.id !== selA.id) return a
        const newSizes = a.sizes.map(sz => {
          const q = parseInt(mv.qtys[sz.talle],10)||0
          return q > 0 ? {...sz, qty: sz.qty - q} : sz
        }).filter(sz => sz.qty > 0)
        return newSizes.length === 0 ? null : {...a, sizes: newSizes}
      }).filter(Boolean)
      // Agregar en destino
      const targetEntry = arts.find(a => a.code === code && a.ubic === newUbic)
      if(targetEntry) {
        arts = arts.map(a => {
          if(a.id !== targetEntry.id) return a
          const newSizes = [...a.sizes]
          toMove.forEach(([t, q]) => {
            const src = selA.sizes.find(sz => sz.talle === t)
            const idx = newSizes.findIndex(sz => sz.talle === t)
            if(idx >= 0) newSizes[idx] = {...newSizes[idx], qty: newSizes[idx].qty + q}
            else newSizes.push({talle:t, qty:q, min:src?.min||0})
          })
          return {...a, sizes: newSizes}
        })
      } else {
        const newSizes = toMove.map(([t, q]) => {
          const src = selA.sizes.find(sz => sz.talle === t)
          return {talle:t, qty:q, min:src?.min||0}
        })
        arts = [...arts, {id:nextId++, code, name:selA.name, cat:selA.cat, ubic:newUbic, sizes:newSizes, precio:selA.precio||0}]
      }
      return {...prev, articles:arts, nextId}
    })
    setModal(null); setView('inventario')
    showToast('Movido a ' + mv.estante + mv.altura + '.')
  }

  // ---- Editar ----
  const openEdit = () => {
    const a = db.articles.find(x => x.id === selectedId); if(!a) return
    setEditing({id:a.id, code:a.code, name:a.name, cat:a.cat, ubic:a.ubic||'', precio:a.precio||''}); setModal('edit')
  }
  const saveEdit = () => {
    if(!editing.code.trim() || !editing.name.trim()) { showToast('Completá código y nombre.'); return }
    setDb(s => ({...s, articles:s.articles.map(a => a.id===editing.id?{...a,code:editing.code.trim(),name:editing.name.trim(),cat:editing.cat,ubic:editing.ubic.trim(),precio:parseFloat(editing.precio)||0}:a)}))
    setModal(null); showToast('Artículo actualizado.')
  }

  // ---- Eliminar ----
  const askDeleteDelivery = (id) => {
    const d = db.deliveries.find(x => x.id === id)
    setConfirm({kind:'delivery', id, title:'Eliminar entrega', msg:'Se eliminará la entrega'+(d&&d.persona?' a '+d.persona:'')+' y se restituirá el stock. ¿Confirmás?'})
  }
  const askDeleteMov = (id) => {
    setConfirm({kind:'mov', id, title:'Eliminar movimiento', msg:'Se eliminará el movimiento y se revertirá su efecto en el stock. ¿Confirmás?'})
  }
  const confirmYes = () => {
    const c = confirm; if(!c) return
    if(c.kind === 'delivery') {
      setDb(s => {
        const del = s.deliveries.find(d => d.id === c.id); if(!del) return {...s}
        const articles = s.articles.map(a => ({...a, sizes:a.sizes.map(z=>({...z}))}))
        del.lines.forEach(l => { const a=articles.find(x=>x.code===l.code); const z=a&&a.sizes.find(x=>x.talle===l.talle); if(z) z.qty+=l.qty })
        return {...s, articles, deliveries:s.deliveries.filter(d=>d.id!==c.id), movimientos:s.movimientos.filter(m=>m.delId!==c.id)}
      })
      showToast('Entrega eliminada y stock restituido.')
    } else {
      setDb(s => {
        const m = s.movimientos.find(x => x.id === c.id); if(!m) return {...s}
        const articles = s.articles.map(a => ({...a, sizes:a.sizes.map(z=>({...z}))}))
        const a = articles.find(x => x.code === m.code); const z = a && a.sizes.find(x => x.talle === m.talle)
        if(z) { if(m.tipo==='entrada') z.qty=Math.max(0,z.qty-m.qty); else z.qty+=m.qty }
        return {...s, articles, movimientos:s.movimientos.filter(x=>x.id!==c.id)}
      })
      showToast('Movimiento eliminado y stock corregido.')
    }
    setConfirm(null)
  }

  // ---- Receptor: aceptar / rechazar entrega ----
  const receptorAceptar = (delId) => {
    setDb(s => {
      const deliveries = s.deliveries.map(d => d.id === delId ? {...d, status:'aceptado', confirmedAt:today()} : d)
      return {...s, deliveries}
    })
    showToast('Entrega aceptada.')
  }
  const receptorRechazar = (delId) => {
    setDb(s => {
      const del = s.deliveries.find(d => d.id === delId); if(!del) return {...s}
      // revert stock
      const articles = s.articles.map(a => ({...a, sizes:a.sizes.map(z=>({...z}))}))
      del.lines.forEach(l => { const a=articles.find(x=>x.code===l.code); const z=a&&a.sizes.find(x=>x.talle===l.talle); if(z) z.qty+=l.qty })
      const deliveries = s.deliveries.map(d => d.id === delId ? {...d, status:'rechazado', confirmedAt:today()} : d)
      return {...s, articles, deliveries}
    })
    showToast('Entrega rechazada y stock restituido.')
  }

  // ---- Derived data ----
  const { articles, deliveries, movimientos } = db
  const codeName = articles.reduce((acc, a) => { acc[a.code] = a.name; return acc }, {})

  // Current user role
  const allUsers = db.users
  const currentUser = allUsers.find(u => u.username === session) || null
  const isReceptor = currentUser?.role === 'receptor'

  // Receptor users list (for the delivery modal selector)
  const receptorUsers = allUsers.filter(u => u.role === 'receptor')

  // KPIs: count unique codes, group by code for category totals
  const byCodeMap = {}
  articles.forEach(a => {
    if(!byCodeMap[a.code]) byCodeMap[a.code] = { cat: a.cat, qty: 0 }
    byCodeMap[a.code].qty += total(a)
  })
  const catTotal = cat => Object.values(byCodeMap).filter(v => v.cat === cat).reduce((s, v) => s + v.qty, 0)

  const kpis = {
    articulos: new Set(articles.map(a => a.code)).size,
    unidades: fmt(articles.reduce((s,a) => s + total(a), 0)),
    valorStock: articles.reduce((s,a) => s + (a.precio||0) * total(a), 0),
    entrenamiento: fmt(catTotal('Entrenamiento')),
    juego: fmt(catTotal('Juego')),
    casual: fmt(catTotal('Casual')),
    bajo: articles.filter(isLow).length,
    entregas: deliveries.length,
  }

  const q = search.trim().toLowerCase()
  let filtered = articles.filter(a => cat==='Todas' || a.cat===cat)
  if(q) filtered = filtered.filter(a => a.name.toLowerCase().includes(q) || a.code.toLowerCase().includes(q) || (a.ubic||'').toLowerCase().includes(q))

  const parseUbic = u => {
    if(!u || u==='—') return { n: Infinity, l: '' }
    const m = u.match(/^(\d+)(.*)$/)
    return m ? { n: parseInt(m[1], 10), l: m[2] } : { n: Infinity, l: u }
  }
  filtered = [...filtered].sort((a, b) => {
    const ua = parseUbic(a.ubic), ub = parseUbic(b.ubic)
    if(ua.n !== ub.n) return ua.n - ub.n
    return ua.l.localeCompare(ub.l)
  })

  // Group by code for inventory rows
  const groupedCodes = {}
  filtered.forEach(a => {
    if(!groupedCodes[a.code]) groupedCodes[a.code] = { ...a, _entries: [a] }
    else groupedCodes[a.code]._entries.push(a)
  })
  const invRows = Object.values(groupedCodes).map(g => {
    const allSizes = g._entries.flatMap(e => e.sizes)
    const tot = allSizes.reduce((s, z) => s + z.qty, 0)
    const low = g._entries.some(e => isLow(e))
    const sortedEntries = [...g._entries].sort((a, b) => {
      const ua = parseUbic(a.ubic), ub = parseUbic(b.ubic)
      if(ua.n !== ub.n) return ua.n - ub.n
      return ua.l.localeCompare(ub.l)
    })
    const ubics = [...new Set(sortedEntries.map(e => e.ubic).filter(Boolean))].join(' · ')
    return {
      ...g,
      totalFmt: fmt(tot),
      sizesLabel: TALLE_ORDER.filter(t => allSizes.some(s => s.talle === t)).join(' · '),
      low,
      ubic: ubics || '—',
      _firstUbic: sortedEntries[0]?.ubic || '',
      precio: g.precio || 0,
      dupUbic: false,
    }
  }).sort((a, b) => {
    const ua = parseUbic(a._firstUbic), ub = parseUbic(b._firstUbic)
    if(ua.n !== ub.n) return ua.n - ub.n
    return ua.l.localeCompare(ub.l)
  })

  const lowList = articles.filter(isLow).map(a => ({
    ...a,
    tallesEnMin: a.sizes.filter(s => (s.min||0) > 0 && s.qty === s.min).length,
    tallesBajo:  a.sizes.filter(s => (s.min||0) > 0 && s.qty < s.min).length,
  }))

  // Talles duplicados: mismo código+talle en más de una ubicación
  const codeTalleCounts = {}
  articles.forEach(a => a.sizes.forEach(s => {
    const k = a.code + ':' + s.talle
    codeTalleCounts[k] = (codeTalleCounts[k]||0)+1
  }))
  const dupCodes = [...new Set(
    articles.filter(a => a.sizes.some(s => codeTalleCounts[a.code+':'+s.talle] > 1)).map(a => a.code)
  )]
  const dupList = dupCodes.map(code => {
    const entries = articles.filter(a => a.code === code)
    const talleUbics = {}
    entries.forEach(a => a.sizes.forEach(s => {
      if(!talleUbics[s.talle]) talleUbics[s.talle] = []
      talleUbics[s.talle].push(a.ubic || '—')
    }))
    const tallesDup = Object.entries(talleUbics)
      .filter(([, ubics]) => ubics.length > 1)
      .map(([talle, ubics]) => ({ talle, ubics }))
    return { code, name: entries[0].name, tallesDup }
  }).filter(d => d.tallesDup.length > 0)

  const delEnrich = d => {
    const totalUd = d.lines.reduce((s,l) => s+l.qty, 0)
    const resumen = d.lines.map(l => (codeName[l.code]||l.code)+' '+l.talle+' ×'+l.qty).join(' · ')
    return {...d, totalUd, resumen, ini:ini(d.persona||d.receptor)}
  }
  const deliveryRows = deliveries.map(delEnrich)
  const deliveryReceptores = [...new Set(deliveries.map(d => d.receptor).filter(Boolean))]
  const filteredDeliveryRows = deliveryRows
    .filter(d => !delFilterReceptor || d.receptor === delFilterReceptor)
    .filter(d => !delFilterPersona || d.persona.toLowerCase().includes(delFilterPersona.toLowerCase()))
  const recentDeliveries = deliveries.slice(0,4).map(delEnrich)
  const pendingApprovals = db.users.filter(u => u.status === 'pendiente')
  const pendingDeliveries = deliveries
    .filter(d => d.status === 'pendiente' && d.toUser)
    .map(d => {
      const u = db.users.find(x => x.username === d.toUser)
      const totalUd = d.lines.reduce((s,l) => s+l.qty, 0)
      return { ...d, totalUd, displayName: u?.displayName || d.toUser, ini: ini(u?.displayName || d.toUser) }
    })

  const movKind = m => {
    const d = m.detalle||''
    if(d.startsWith('Entrega')) return 'entrega'
    if(d.startsWith('Devolución')) return 'devolucion'
    if(d.startsWith('Ajuste')) return 'ajuste'
    return 'ingreso'
  }
  const movChipDefs = [['Todos',null],['Entregas','entrega'],['Devoluciones','devolucion'],['Ajustes','ajuste'],['Ingresos','ingreso']]
  const movFilterKind = (movChipDefs.find(c=>c[0]===movFilter)||[])[1]
  const movRows = movimientos.filter(m => !movFilterKind || movKind(m)===movFilterKind)

  const saveUti = () => {
    if(!utiForm.numero.trim()) { showToast('Ingresá el número de camiseta.'); return }
    setDb(prev => {
      const list = prev.camisetasUtileria || []
      if(utiForm.id !== null) {
        return {...prev, camisetasUtileria: list.map(c => c.id === utiForm.id ? {...utiForm} : c)}
      } else {
        const newId = list.length > 0 ? Math.max(...list.map(c => c.id)) + 1 : 1
        return {...prev, camisetasUtileria: [...list, {...utiForm, id: newId}]}
      }
    })
    setUtiModal(false)
    showToast(utiForm.id !== null ? 'Camiseta actualizada.' : 'Camiseta agregada.')
  }
  const deleteUti = (id) => {
    setDb(prev => ({...prev, camisetasUtileria: (prev.camisetasUtileria||[]).filter(c => c.id !== id)}))
    showToast('Camiseta eliminada.')
  }
  const utiFiltered = (db.camisetasUtileria || []).filter(c => !utiFilter || c.competicion === utiFilter)

  const receptorCards = RECEPTORES.map(name => {
    const ds = deliveries.filter(d => d.receptor===name)
    const unidades = ds.reduce((s,d) => s+d.lines.reduce((x,l)=>x+l.qty,0),0)
    return { name, ini:ini(name), count:ds.length, unidades }
  })

  // selA: for modals that operate on a specific entry (selectedId)
  const selA = articles.find(a => a.id === selectedId)

  // selEntries: all entries for the selected code (for detail view)
  const selEntries = selectedCode ? articles.filter(a => a.code === selectedCode) : []

  let detail = null
  if(selEntries.length > 0) {
    const allSizesFlat = selEntries.flatMap(e => e.sizes)
    const tot = allSizesFlat.reduce((s, z) => s + z.qty, 0)
    const low = selEntries.some(e => isLow(e))
    const movs = movimientos.filter(m => m.code === selectedCode)
    const first = selEntries[0]
    const entriesDisplay = selEntries.map(entry => {
      const maxQ = Math.max(1, ...entry.sizes.map(s => s.qty))
      const ordered = [...entry.sizes].sort((a,b) => TALLE_ORDER.indexOf(a.talle) - TALLE_ORDER.indexOf(b.talle))
      const sizes = ordered.map(s => {
        const sLow = (s.min||0) > 0 && s.qty <= (s.min||0)
        return { ...s, isLow: sLow, pct: Math.round(s.qty / maxQ * 100) }
      })
      return { ...entry, sizes, ubic: entry.ubic || '—' }
    })
    detail = {
      ...first,
      entries: entriesDisplay,
      total: tot,
      totalFmt: fmt(tot),
      low,
      ubic: selEntries.map(e => e.ubic || '—').join(' · '),
      movs,
      noMovs: movs.length === 0,
    }
  }

  // nd derived
  const ndUbics = nd.cCode ? [...new Set(db.articles.filter(a => a.code === nd.cCode).map(a => a.ubic).filter(Boolean))] : []
  const ndHasMultiUbic = ndUbics.length > 1
  const effectiveUbic = nd.cUbic || (ndUbics.length === 1 ? ndUbics[0] : '')
  const ndA = nd.cCode ? db.articles.find(a => a.code === nd.cCode && (!effectiveUbic || a.ubic === effectiveUbic)) : null
  const ndTalleOptions = ndA ? ndA.sizes.filter(s => s.qty > 0).map(s => ({value:s.talle, label:s.talle+' ('+s.qty+' disp.)'})) : []
  let stockHint = ''
  if(nd.cCode && nd.cTalle && ndA) { const z=ndA.sizes.find(s=>s.talle===nd.cTalle); if(z) stockHint='Disponible: '+z.qty+' u. en talle '+nd.cTalle+(effectiveUbic?' · Ubic. '+effectiveUbic:'') }
  const ndTotal = nd.lines.reduce((s,l) => s+l.qty, 0)
  const ndMonto = nd.receptor === 'Protocolo' && nd.paga === 'si'
    ? nd.lines.reduce((s,l) => { const art=articles.find(a=>a.code===l.code); return s+(art?.precio||0)*l.qty }, 0) * 0.5
    : 0
  const ndOk = nd.persona && nd.persona.trim() && nd.receptor && nd.lines.length > 0

  const repTalleOptions = selA ? selA.sizes.map(s => ({value:s.talle, label:s.talle})) : []
  const ajTalleOptions = selA ? selA.sizes.map(s => ({value:s.talle, label:s.talle+' (sistema: '+s.qty+')'})) : []

  const ndIsDev = nd.mode === 'devolucion'

  if(!session) return (
    <div style={{display:'flex',alignItems:'center',justifyContent:'center',minHeight:'100dvh',background:'#121212',flexDirection:'column',padding:20,overflowY:'auto'}}>
      <div style={{background:'#1a1a1a',borderRadius:12,padding:'36px 32px',width:'100%',maxWidth:420,boxShadow:'0 20px 60px rgba(0,0,0,.5)'}}>
        <div style={{display:'flex',flexDirection:'column',alignItems:'center',marginBottom:28}}>
          <img src="/escudo.png" alt="Peñarol" style={{height:56,marginBottom:14}} />
          <div style={{fontFamily:'Archivo Black,sans-serif',fontSize:18,color:'#FFD200',letterSpacing:'.05em'}}>INDUMENTARIA PEÑAROL</div>
        </div>

        {/* Tabs login / registro — oculto en vistas auxiliares */}
        {(loginView === 'login' || loginView === 'register') && (
          <div style={{display:'flex',borderBottom:'1px solid #2a2a2a',marginBottom:24}}>
            {[['login','Ingresar'],['register','Registrarse']].map(([v,label]) => (
              <button key={v} onClick={()=>setLoginView(v)} style={{flex:1,padding:'10px 0',background:'none',border:'none',cursor:'pointer',fontWeight:700,fontSize:13,color:loginView===v?'#FFD200':'#8a8a82',borderBottom:loginView===v?'2px solid #FFD200':'2px solid transparent',transition:'all .15s'}}>
                {label}
              </button>
            ))}
          </div>
        )}

        {loginView === 'login' && (
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <div className="form-group">
              <label className="field-label" style={{color:'#8a8a82'}}>CORREO / USUARIO</label>
              <input className="field-input" value={loginForm.user} onChange={e=>setLoginForm(p=>({...p,user:e.target.value,err:''}))} onKeyDown={e=>e.key==='Enter'&&doLogin()} placeholder="usuario o correo" autoComplete="username" />
            </div>
            <div className="form-group">
              <label className="field-label" style={{color:'#8a8a82'}}>CONTRASEÑA</label>
              <input className="field-input" type="password" value={loginForm.pass} onChange={e=>setLoginForm(p=>({...p,pass:e.target.value,err:''}))} onKeyDown={e=>e.key==='Enter'&&doLogin()} placeholder="••••••••" autoComplete="current-password" />
            </div>
            {loginForm.err && <div style={{fontSize:12.5,color:'#C2473D',fontWeight:600}}>{loginForm.err}</div>}
            <button className="btn btn-yellow" style={{width:'100%',justifyContent:'center',marginTop:4,height:44}} onClick={doLogin}>Ingresar</button>
            <button onClick={()=>{setLoginView('forgot');setForgotForm({email:'',newPass:'',newPass2:'',step:1,err:''})}} style={{background:'none',border:'none',color:'#8a8a82',cursor:'pointer',fontSize:12,marginTop:4,padding:0,textDecoration:'underline'}}>¿Olvidaste tu contraseña?</button>
          </div>
        )}

        {loginView === 'forgot' && (
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <div style={{fontSize:13,color:'#8a8a82',marginBottom:2}}>{forgotForm.step===1 ? 'Ingresá el correo con el que te registraste.' : 'Elegí una nueva contraseña.'}</div>
            {forgotForm.step === 1 && (
              <div className="form-group">
                <label className="field-label" style={{color:'#8a8a82'}}>CORREO ELECTRÓNICO</label>
                <input className="field-input" type="email" value={forgotForm.email} onChange={e=>setForgotForm(p=>({...p,email:e.target.value,err:''}))} onKeyDown={e=>e.key==='Enter'&&doForgotStep1()} placeholder="correo@ejemplo.com" autoComplete="email" />
              </div>
            )}
            {forgotForm.step === 2 && (<>
              <div className="form-group">
                <label className="field-label" style={{color:'#8a8a82'}}>NUEVA CONTRASEÑA</label>
                <input className="field-input" type="password" value={forgotForm.newPass} onChange={e=>setForgotForm(p=>({...p,newPass:e.target.value,err:''}))} placeholder="Mín. 6 caracteres" />
              </div>
              <div className="form-group">
                <label className="field-label" style={{color:'#8a8a82'}}>REPETIR CONTRASEÑA</label>
                <input className="field-input" type="password" value={forgotForm.newPass2} onChange={e=>setForgotForm(p=>({...p,newPass2:e.target.value,err:''}))} placeholder="••••••••" />
              </div>
            </>)}
            {forgotForm.err && <div style={{fontSize:12.5,color:'#C2473D',fontWeight:600}}>{forgotForm.err}</div>}
            <button className="btn btn-yellow" style={{width:'100%',justifyContent:'center',height:44}} onClick={forgotForm.step===1?doForgotStep1:doForgotStep2}>{forgotForm.step===1?'Continuar':'Guardar contraseña'}</button>
            <button onClick={()=>setLoginView('login')} style={{background:'none',border:'none',color:'#8a8a82',cursor:'pointer',fontSize:12,padding:0,textDecoration:'underline'}}>Volver al inicio de sesión</button>
          </div>
        )}

        {loginView === 'register' && (
          <div style={{display:'flex',flexDirection:'column',gap:12}}>
            <div className="form-group">
              <label className="field-label" style={{color:'#8a8a82'}}>NOMBRE COMPLETO</label>
              <input className="field-input" value={regForm.displayName} onChange={e=>setRegForm(p=>({...p,displayName:e.target.value,err:''}))} placeholder="Ej. Juan García" />
            </div>
            <div className="form-group">
              <label className="field-label" style={{color:'#8a8a82'}}>CORREO ELECTRÓNICO</label>
              <input className="field-input" type="email" value={regForm.email} onChange={e=>setRegForm(p=>({...p,email:e.target.value,err:''}))} placeholder="correo@ejemplo.com" autoComplete="email" />
            </div>
            <div className="form-cols-2">
              <div className="form-group">
                <label className="field-label" style={{color:'#8a8a82'}}>TELÉFONO</label>
                <input className="field-input" type="tel" value={regForm.telefono} onChange={e=>setRegForm(p=>({...p,telefono:e.target.value,err:''}))} placeholder="09X XXX XXX" />
              </div>
              <div className="form-group">
                <label className="field-label" style={{color:'#8a8a82'}}>CARGO</label>
                <select className="field-input" value={regForm.cargo} onChange={e=>setRegForm(p=>({...p,cargo:e.target.value,err:''}))}>
                  <option value="">Seleccioná tu cargo…</option>
                  {CARGOS_REG.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            </div>
            <div className="form-group">
              <label className="field-label" style={{color:'#8a8a82'}}>SECTOR</label>
              <select className="field-input" value={regForm.categoria} onChange={e=>setRegForm(p=>({...p,categoria:e.target.value,division:'',err:''}))}>
                <option value="">Seleccioná tu sector…</option>
                {OCUPACIONES.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>
            {regForm.categoria && regForm.categoria !== '3° División' && (
              <div className="form-group">
                <label className="field-label" style={{color:'#8a8a82'}}>DIVISIÓN</label>
                <select className="field-input" value={regForm.division} onChange={e=>setRegForm(p=>({...p,division:e.target.value,err:''}))}>
                  <option value="">Seleccioná tu división…</option>
                  {DIVISIONES.map(d => <option key={d} value={d}>{d}</option>)}
                </select>
              </div>
            )}
            <div className="form-cols-2">
              <div className="form-group">
                <label className="field-label" style={{color:'#8a8a82'}}>CONTRASEÑA</label>
                <input className="field-input" type="password" value={regForm.pass} onChange={e=>setRegForm(p=>({...p,pass:e.target.value,err:''}))} placeholder="Mín. 6 caracteres" />
              </div>
              <div className="form-group">
                <label className="field-label" style={{color:'#8a8a82'}}>REPETIR CONTRASEÑA</label>
                <input className="field-input" type="password" value={regForm.pass2} onChange={e=>setRegForm(p=>({...p,pass2:e.target.value,err:''}))} placeholder="••••••••" />
              </div>
            </div>
            {regForm.err && <div style={{fontSize:12.5,color:'#C2473D',fontWeight:600}}>{regForm.err}</div>}
            <button className="btn btn-yellow" style={{width:'100%',justifyContent:'center',marginTop:4,height:44}} onClick={doRegister}>Crear cuenta</button>
          </div>
        )}

        {loginView === 'registered' && (
          <div style={{display:'flex',flexDirection:'column',alignItems:'center',gap:18,padding:'8px 0 4px'}}>
            <div style={{width:56,height:56,borderRadius:'50%',background:'#1e3a2e',display:'flex',alignItems:'center',justifyContent:'center',fontSize:26}}>✓</div>
            <div style={{textAlign:'center'}}>
              <div style={{fontWeight:700,fontSize:15,color:'#fff',marginBottom:6}}>Cuenta creada</div>
              <div style={{fontSize:13,color:'#8a8a82',lineHeight:1.5}}>Tu solicitud fue enviada.<br/>El administrador debe aprobarla antes de que puedas ingresar.</div>
            </div>
            <button className="btn btn-yellow" style={{width:'100%',justifyContent:'center',height:44}} onClick={()=>setLoginView('login')}>Volver al inicio de sesión</button>
          </div>
        )}
      </div>
    </div>
  )

  if (loading) return (
    <div style={{display:'flex',alignItems:'center',justifyContent:'center',height:'100dvh',flexDirection:'column',gap:16,background:'#121212'}}>
      <img src="/escudo.png" alt="Peñarol" style={{height:64,opacity:.9}} />
      <div style={{color:'#FFD200',fontFamily:'Archivo Black,sans-serif',fontSize:14,letterSpacing:'.1em'}}>CARGANDO…</div>
    </div>
  )

  // ---- Vista del receptor ----
  if (isReceptor) {
    const myDeliveries = db.deliveries.filter(d => d.toUser === session)
    const pendientes = myDeliveries.filter(d => (d.status||'aceptado') === 'pendiente')
    const historial  = myDeliveries.filter(d => (d.status||'aceptado') !== 'pendiente')
    const rCodeName  = db.articles.reduce((acc, a) => { acc[a.code] = a.name; return acc }, {})
    return (
      <div style={{minHeight:'100dvh',background:'#F6F6F4',fontFamily:'Archivo,sans-serif'}}>
        {/* Header */}
        <div style={{background:'#121212',padding:'18px 24px',display:'flex',alignItems:'center',gap:16}}>
          <img src="/escudo.png" alt="Peñarol" style={{height:44}} />
          <div style={{flex:1}}>
            <div style={{fontFamily:'Archivo Black,sans-serif',fontSize:14,color:'#FFD200',letterSpacing:'.05em'}}>DEPÓSITO · INDUMENTARIA</div>
            <div style={{fontSize:13,color:'#fff',marginTop:2}}>Hola, <b>{currentUser?.displayName || session}</b></div>
          </div>
          <button onClick={doLogout} style={{background:'#2a2a2a',border:'1px solid #3a3a3a',color:'#ccc',borderRadius:8,padding:'8px 14px',cursor:'pointer',fontSize:13}}>Cerrar sesión</button>
        </div>

        <div style={{maxWidth:680,margin:'0 auto',padding:'28px 16px',display:'flex',flexDirection:'column',gap:28}}>

          {/* Pendientes */}
          <div>
            <div style={{fontFamily:'Archivo Black,sans-serif',fontSize:13,color:'#8a8a82',letterSpacing:'.08em',marginBottom:14}}>PENDIENTES DE CONFIRMACIÓN</div>
            {pendientes.length === 0 && (
              <div style={{background:'#fff',borderRadius:12,padding:'20px 20px',textAlign:'center',color:'#8a8a82',fontSize:13.5,border:'1px solid #E7E7E3'}}>Sin entregas pendientes.</div>
            )}
            {pendientes.map(d => (
              <div key={d.id} style={{background:'#fff',borderRadius:12,border:'1px solid #E7E7E3',marginBottom:12,overflow:'hidden',boxShadow:'0 1px 4px rgba(0,0,0,.06)'}}>
                <div style={{padding:'14px 18px',borderBottom:'1px solid #F0F0EC',display:'flex',alignItems:'center',gap:12}}>
                  <span style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200',borderRadius:6,padding:'2px 9px',fontSize:12,fontWeight:700}}>Pendiente</span>
                  <span style={{fontSize:12.5,color:'#8a8a82',fontFamily:'IBM Plex Mono,monospace'}}>{d.fecha}</span>
                  <span style={{fontSize:12.5,color:'#6a6a62',flex:1,textAlign:'right'}}>{d.receptor}</span>
                </div>
                <div style={{padding:'14px 18px'}}>
                  {d.lines.map((l,i) => (
                    <div key={i} style={{display:'flex',alignItems:'center',gap:10,padding:'5px 0',borderBottom:i<d.lines.length-1?'1px solid #F5F5F0':'none'}}>
                      <span style={{flex:1,fontWeight:600,fontSize:13.5}}>{rCodeName[l.code]||l.code}</span>
                      <span style={{fontSize:12.5,color:'#6a6a62'}}>Talle {l.talle}</span>
                      <span style={{fontFamily:'IBM Plex Mono,monospace',fontWeight:700,fontSize:13.5}}>×{l.qty}</span>
                    </div>
                  ))}
                </div>
                <div style={{padding:'12px 18px',borderTop:'1px solid #F0F0EC',display:'flex',gap:10}}>
                  <button onClick={() => receptorAceptar(d.id)} style={{flex:1,background:'#2e9b5e',color:'#fff',border:'none',borderRadius:8,padding:'10px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>✓ Aceptar</button>
                  <button onClick={() => receptorRechazar(d.id)} style={{flex:1,background:'#C2473D',color:'#fff',border:'none',borderRadius:8,padding:'10px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>✕ Rechazar</button>
                </div>
              </div>
            ))}
          </div>

          {/* Historial */}
          <div>
            <div style={{fontFamily:'Archivo Black,sans-serif',fontSize:13,color:'#8a8a82',letterSpacing:'.08em',marginBottom:14}}>HISTORIAL</div>
            {historial.length === 0 && (
              <div style={{background:'#fff',borderRadius:12,padding:'20px 20px',textAlign:'center',color:'#8a8a82',fontSize:13.5,border:'1px solid #E7E7E3'}}>Sin entregas en el historial.</div>
            )}
            {historial.map(d => {
              const st = d.status || 'aceptado'
              const stColor = st === 'aceptado' ? '#2e9b5e' : '#C2473D'
              const stBg    = st === 'aceptado' ? '#EDF7F2' : '#FBEAE8'
              const stLabel = st === 'aceptado' ? 'Aceptado' : 'Rechazado'
              return (
                <div key={d.id} style={{background:'#fff',borderRadius:12,border:'1px solid #E7E7E3',marginBottom:12,overflow:'hidden',boxShadow:'0 1px 4px rgba(0,0,0,.06)'}}>
                  <div style={{padding:'14px 18px',borderBottom:'1px solid #F0F0EC',display:'flex',alignItems:'center',gap:12}}>
                    <span style={{background:stBg,color:stColor,border:'1px solid '+stColor,borderRadius:6,padding:'2px 9px',fontSize:12,fontWeight:700}}>{stLabel}</span>
                    <span style={{fontSize:12.5,color:'#8a8a82',fontFamily:'IBM Plex Mono,monospace'}}>{d.fecha}</span>
                    {d.confirmedAt && <span style={{fontSize:11.5,color:'#8a8a82',flex:1,textAlign:'right'}}>Confirmado: {d.confirmedAt}</span>}
                  </div>
                  <div style={{padding:'14px 18px'}}>
                    {d.lines.map((l,i) => (
                      <div key={i} style={{display:'flex',alignItems:'center',gap:10,padding:'5px 0',borderBottom:i<d.lines.length-1?'1px solid #F5F5F0':'none'}}>
                        <span style={{flex:1,fontWeight:600,fontSize:13.5}}>{rCodeName[l.code]||l.code}</span>
                        <span style={{fontSize:12.5,color:'#6a6a62'}}>Talle {l.talle}</span>
                        <span style={{fontFamily:'IBM Plex Mono,monospace',fontWeight:700,fontSize:13.5}}>×{l.qty}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )
            })}
          </div>
        </div>

        {/* Toast */}
        {toast && (
          <div className="toast">
            <span className="toast-dot"/>
            {toast}
          </div>
        )}
      </div>
    )
  }

  return (
    <div className="app-shell">
      {/* Mobile overlay */}
      <div className={`mobile-overlay${sidebarOpen?' open':''}`} onClick={() => setSidebarOpen(false)} />

      {/* Sidebar */}
      <aside className={`sidebar${sidebarOpen?' open':''}`}>
        <div className="sidebar-logo">
          <img src="/escudo.png" alt="Peñarol" />
          <div className="sidebar-logo-text">
            <div className="name">PEÑAROL</div>
            <div className="sub">DEPÓSITO · INDUMENTARIA</div>
          </div>
        </div>
        <nav className="sidebar-nav">
          {[['panel','PANEL'],['inventario','INVENTARIO'],['entregas','ENTREGAS'],['movimientos','MOVIMIENTOS'],['receptores','RECEPTORES'],['usuarios-reg','USUARIOS REGISTRADOS'],['utileria','CAMISETAS UTILERÍA']].map(([key,label]) => {
            const isActive = view===key||(key==='inventario'&&view==='detalle')
            return (
              <button key={key} className={`nav-item${isActive?' active':''}`} onClick={() => goView(key)}>
                <span className="nav-dot" />
                {label}
                {isActive && <img src="/escudo.png" alt="" style={{height:20,width:'auto',marginLeft:'auto',opacity:0.85}} />}
              </button>
            )
          })}
        </nav>
        <div className="sidebar-user">
          <div className="user-avatar">{ini(session||'')}</div>
          <div style={{flex:1,minWidth:0}}>
            <div className="user-name">{(session||'').toUpperCase()}</div>
            <div className="user-role">Gestión de depósito</div>
          </div>
          <button title="Cambiar contraseña" onClick={()=>{setChangePassForm({current:'',newPass:'',newPass2:'',err:''});setModal('cambiar-pass')}} style={{background:'none',border:'none',color:'#8a8a82',cursor:'pointer',fontSize:16,padding:'0 4px',flexShrink:0}}>🔑</button>
          <button title="Gestionar usuarios" onClick={openUserMgmt} style={{background:'none',border:'none',color:'#8a8a82',cursor:'pointer',fontSize:18,padding:'0 4px',flexShrink:0}}>⚙</button>
          <button title="Cerrar sesión" onClick={doLogout} style={{background:'none',border:'none',color:'#8a8a82',cursor:'pointer',fontSize:18,padding:'0 4px',flexShrink:0}}>⏻</button>
        </div>
      </aside>

      {/* Main */}
      <div className="main-area">
        <header className="topbar">
          <button className="hamburger" onClick={() => setSidebarOpen(o=>!o)} aria-label="Menú">
            <span/><span/><span/>
          </button>
          <img src="/1891_Amarillo.jpg" alt="1891" style={{height:28,width:'auto'}} />
          <div className="topbar-title">
            {{panel:'PANEL',inventario:'INVENTARIO',detalle:'DETALLE',entregas:'ENTREGAS',movimientos:'MOVIMIENTOS',receptores:'RECEPTORES','usuarios-reg':'USUARIOS REGISTRADOS',utileria:'CAMISETAS UTILERÍA'}[view]}
          </div>
          <div className="topbar-spacer" />
          <div className="search-box">
            <span className="search-icon" />
            <input value={search} onChange={e => { setSearch(e.target.value); if((view==='panel'||view==='detalle')&&e.target.value) setView('inventario') }} placeholder="Buscar…" />
          </div>
          <button className="btn btn-ghost" onClick={openArticulo}>+<span className="btn-label"> Artículo</span></button>
          <button className="btn btn-ghost" onClick={openDevolucion}>↩<span className="btn-label"> Dev.</span></button>
          <button className="btn btn-yellow" onClick={openEntrega}>+<span className="btn-label"> Entrega</span></button>
        </header>

        <div className="content">
          {/* PANEL */}
          {view === 'panel' && (
            <>
              <div className="kpi-grid">
                <div className="kpi-card"><div className="kpi-label">ARTÍCULOS</div><div className="kpi-value">{kpis.articulos}</div><div className="kpi-sub">referencias activas</div></div>
                <div className="kpi-card"><div className="kpi-label">UNIDADES EN STOCK</div><div className="kpi-value">{kpis.unidades}</div><div className="kpi-sub">suma de todos los talles</div></div>

                <div className="kpi-card"><div className="kpi-label">MONTO TOTAL EN ARTÍCULOS</div><div className="kpi-value" style={{fontSize:24}}>$ {kpis.valorStock.toLocaleString('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2})}</div></div>
                <div className="kpi-card" style={{cursor:'pointer'}} onClick={() => { setCat('Entrenamiento'); setView('inventario') }}><div className="kpi-label">PRENDAS DE ENTRENAMIENTO</div><div className="kpi-value">{kpis.entrenamiento}</div><div className="kpi-sub">unidades en stock →</div></div>
                <div className="kpi-card" style={{cursor:'pointer'}} onClick={() => { setCat('Juego'); setView('inventario') }}><div className="kpi-label">PRENDAS DE JUEGO</div><div className="kpi-value">{kpis.juego}</div><div className="kpi-sub">unidades en stock →</div></div>
                <div className="kpi-card" style={{cursor:'pointer'}} onClick={() => { setCat('Casual'); setView('inventario') }}><div className="kpi-label">PRENDAS CASUAL</div><div className="kpi-value">{kpis.casual}</div><div className="kpi-sub">unidades en stock →</div></div>
              </div>
              <div className="panel-grid">
                <div className="card">
                  <div className="card-header">
                    <div className="card-title">Reposición necesaria</div>
                    <div className="card-spacer"/>
                    <span className="badge low">{kpis.bajo} artículos</span>
                  </div>
                  {lowList.length === 0 && <div className="empty">Sin artículos bajo mínimo 🎉</div>}
                  {lowList.map(a => (
                    <div key={a.id} className="table-row clickable" style={{gridTemplateColumns:'1fr auto'}} onClick={() => openDetail(a.code)}>
                      <div>
                        <div style={{fontWeight:600,fontSize:13.5}}>{a.name}</div>
                        <div style={{fontSize:11.5,color:'#8a8a82',fontFamily:'IBM Plex Mono,monospace'}}>{a.code}</div>
                      </div>
                      <div style={{textAlign:'right'}}>
                        {a.tallesBajo > 0 && <div style={{fontSize:12,color:'#C2473D',fontWeight:600}}>{a.tallesBajo} talle{a.tallesBajo>1?'s':''} por debajo</div>}
                        {a.tallesEnMin > 0 && <div style={{fontSize:12,color:'#b87000',fontWeight:600}}>{a.tallesEnMin} talle{a.tallesEnMin>1?'s':''} en el mínimo</div>}
                      </div>
                    </div>
                  ))}
                </div>
                <div className="card">
                  <div className="card-header">
                    <div className="card-title">Entregas pendientes de respuesta</div>
                    <div className="card-spacer"/>
                    {pendingDeliveries.length > 0 && <span className="badge" style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200'}}>{pendingDeliveries.length}</span>}
                    <button className="back-link" style={{color:'#9a7d00',margin:0}} onClick={() => goView('entregas')}>Ver todas →</button>
                  </div>
                  {pendingDeliveries.length === 0
                    ? <div className="empty">No hay entregas pendientes de confirmación.</div>
                    : pendingDeliveries.map(d => (
                      <div key={d.id} className="table-row" style={{gridTemplateColumns:'34px 1fr auto'}}>
                        <div className="avatar" style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200'}}>{d.ini}</div>
                        <div style={{minWidth:0}}>
                          <div style={{fontWeight:600,fontSize:13.5,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.displayName}</div>
                          <div style={{fontSize:11.5,color:'#8a8a82'}}>
                            {d.lines.length} artículo{d.lines.length!==1?'s':''} · {d.totalUd} unidades
                          </div>
                        </div>
                        <div style={{textAlign:'right',flexShrink:0}}>
                          <span style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200',borderRadius:5,padding:'2px 8px',fontSize:11,fontWeight:700}}>Pendiente</span>
                          <div style={{fontSize:11,color:'#8a8a82',marginTop:3}}>{d.fecha}</div>
                        </div>
                      </div>
                    ))
                  }
                </div>
                {!isReceptor && pendingApprovals.length > 0 && (
                  <div className="card" style={{marginTop:16}}>
                    <div className="card-header">
                      <div className="card-title">Solicitudes de acceso pendientes</div>
                      <div className="card-spacer"/>
                      <span className="badge" style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200'}}>{pendingApprovals.length}</span>
                      <button className="back-link" style={{color:'#9a7d00',margin:0}} onClick={() => goView('usuarios-reg')}>Ver →</button>
                    </div>
                    {pendingApprovals.map(u => (
                      <div key={u.username} className="table-row" style={{gridTemplateColumns:'34px 1fr auto'}}>
                        <div className="avatar" style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200',opacity:.8}}>{ini(u.displayName||u.username)}</div>
                        <div style={{minWidth:0}}>
                          <div style={{fontWeight:600,fontSize:13.5,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{u.displayName||u.username}</div>
                          <div style={{fontSize:11.5,color:'#8a8a82'}}>{u.cargo}{u.categoria ? ' · '+u.categoria : ''}</div>
                        </div>
                        <div style={{display:'flex',gap:6,flexShrink:0}}>
                          <button onClick={()=>approveUser(u.username)} style={{padding:'4px 10px',borderRadius:5,border:'none',cursor:'pointer',fontWeight:700,fontSize:11.5,background:'#FFD200',color:'#121212'}}>Aprobar</button>
                          <button onClick={()=>rejectUser(u.username)} style={{padding:'4px 10px',borderRadius:5,border:'1px solid #C2473D',cursor:'pointer',fontWeight:700,fontSize:11.5,background:'none',color:'#C2473D'}}>Rechazar</button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
                <div className="card">
                <div className="card-header">
                  <div className="card-title">⚠ Talles duplicados en múltiples ubicaciones</div>
                  <div className="card-spacer"/>
                  {dupList.length > 0 && <span className="badge" style={{background:'#FFF0C2',color:'#7a5800',border:'1px solid #FFD200'}}>{dupList.length} artículo{dupList.length>1?'s':''}</span>}
                </div>
                {dupList.length === 0 && <div className="empty">No hay artículos en ubicaciones repetidas.</div>}
                {dupList.map(d => (
                  <div key={d.code} className="table-row clickable" style={{gridTemplateColumns:'1fr auto',cursor:'pointer'}} onClick={() => openDetail(d.code)}>
                    <div>
                      <div style={{fontWeight:600,fontSize:13.5}}>{d.name}</div>
                      <div style={{fontSize:11.5,color:'#8a8a82',fontFamily:'IBM Plex Mono,monospace'}}>{d.code}</div>
                    </div>
                    <div style={{textAlign:'right',display:'flex',flexDirection:'column',gap:3}}>
                      {d.tallesDup.map(t => (
                        <div key={t.talle} style={{fontSize:12,fontWeight:600,color:'#C2473D'}}>
                          Talle <b>{t.talle}</b>: {t.ubics.join(' · ')}
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
              </div>
            </>
          )}

          {/* INVENTARIO */}
          {view === 'inventario' && (
            <>
              <div className="chips" style={{display:'flex',alignItems:'center',gap:8,flexWrap:'wrap'}}>
                {['Todas',...CATEGORIAS].map(c => (
                  <button key={c} className={`chip${cat===c?' active':''}`} onClick={() => setCat(c)}>{c}</button>
                ))}
                <div style={{flex:1}}/>
                <button className="btn btn-ghost" style={{fontSize:12.5,padding:'5px 12px'}} onClick={exportExcel}>↓ Exportar Excel</button>
              </div>
              <div className="card table-wrap">
                <div className="table-header inv-cols">
                  <div>CÓDIGO</div><div>ARTÍCULO</div>
                  <div className="inv-col-ubic">UBIC.</div>
                  <div className="inv-col-cat">CATEGORÍA</div>
                  <div className="inv-col-sizes">TALLES</div>
                  <div style={{textAlign:'right'}}>STOCK</div>
                  <div style={{textAlign:'right'}}>ESTADO</div>
                  <div className="inv-col-precio" style={{textAlign:'right'}}>PRECIO SOCIO</div>
                </div>
                {invRows.map(r => (
                  <div key={r.code} className="table-row clickable inv-cols" onClick={() => openDetail(r.code)}>
                    <div className="mono" style={{fontSize:12.5,color:'#1a1a1a',fontWeight:500}}>{r.code}</div>
                    <div style={{fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.name}</div>
                    <div className="inv-col-ubic"><span className="ubic-badge">{r.ubic}</span></div>
                    <div className="inv-col-cat" style={{color:'#1a1a1a'}}>{r.cat}</div>
                    <div className="inv-col-sizes" style={{color:'#1a1a1a'}}>{r.sizesLabel}</div>
                    <div style={{textAlign:'right',fontWeight:700,fontFamily:'IBM Plex Mono,monospace'}}>{r.totalFmt}</div>
                    <div style={{textAlign:'right',display:'flex',gap:4,justifyContent:'flex-end',flexWrap:'wrap'}}>
                      {r.low && <span className="badge low">Bajo mín.</span>}
                    </div>
                    <div className="inv-col-precio mono" style={{textAlign:'right',fontSize:12.5,color:'#1a1a1a'}}>
                      {r.precio > 0 ? '$ '+r.precio.toLocaleString('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2}) : '—'}
                    </div>
                  </div>
                ))}
                {invRows.length === 0 && <div className="empty">{search ? `Sin resultados para «${search}».` : 'Sin artículos. Creá el primero con + Artículo.'}</div>}
              </div>
            </>
          )}

          {/* DETALLE */}
          {view === 'detalle' && detail && (
            <>
              <button className="back-link" onClick={() => setView('inventario')}>← Volver al inventario</button>
              <div className="detail-grid">
                <div className="card">
                  {/* Header global del artículo */}
                  <div style={{padding:'22px 24px',borderBottom:'1px solid #E7E7E3'}}>
                    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',gap:16}}>
                      <div style={{flex:1}}>
                        <div className="mono" style={{fontSize:12.5,color:'#8a8a82'}}>{detail.code}</div>
                        <div style={{fontWeight:800,fontSize:22,marginTop:3}}>{detail.name}</div>
                        <div style={{display:'flex',gap:8,marginTop:9,alignItems:'center',flexWrap:'wrap'}}>
                          <span className="badge gray">{detail.cat}</span>
                          {detail.low && <span className="badge low">Bajo mínimo</span>}
                        </div>
                      </div>
                      <div style={{textAlign:'right',flexShrink:0}}>
                        <div style={{fontFamily:'Archivo Black,sans-serif',fontSize:30,lineHeight:1}}>{detail.totalFmt}</div>
                        <div style={{fontSize:11.5,color:'#8a8a82',marginTop:4}}>unidades totales</div>
                        {detail.precio > 0 && <div style={{marginTop:8,fontSize:13,fontWeight:700,color:'#1a1a1a'}}>$ {detail.precio.toLocaleString('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2})}</div>}
                      </div>
                    </div>
                  </div>

                  {/* Stock por ubicación */}
                  <div style={{padding:'18px 24px'}}>
                    {detail.entries.map((entry, idx) => {
                      const entryTot = entry.sizes.reduce((s, z) => s + z.qty, 0)
                      return (
                        <div key={entry.id}>
                          {idx > 0 && <div style={{borderTop:'1px solid #E7E7E3',margin:'20px 0'}} />}
                          <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:14}}>
                            <span className="ubic-badge" style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200',fontWeight:700}}>
                              <span style={{fontSize:11,color:'#9a7d00',fontFamily:'Archivo,sans-serif'}}>UBIC. </span>{entry.ubic}
                            </span>
                            <span style={{fontSize:12.5,color:'#8a8a82'}}>{fmt(entryTot)} u.</span>
                          </div>
                          <div style={{fontSize:12,color:'#8a8a82',fontWeight:700,letterSpacing:'.04em',marginBottom:14}}>STOCK POR TALLE</div>
                          {entry.sizes.map(s => (
                            <div key={s.talle} className="bar-row">
                              <div style={{width:46,fontWeight:700,fontSize:13.5}}>{s.talle}</div>
                              <div className="bar-track"><div className="bar-fill" style={{width:s.pct+'%',background:s.isLow||s.qty<=0?'#C2473D':'#FFD200'}} /></div>
                              <div style={{textAlign:'right',flexShrink:0}}>
                                <div className="mono" style={{fontWeight:600,fontSize:13.5}}>{s.qty}</div>
                                <div style={{fontSize:10.5,color:'#8a8a82'}}>mín {s.min}</div>
                              </div>
                            </div>
                          ))}
                          {/* Botones por ubicación */}
                          <div style={{display:'flex',gap:8,flexWrap:'wrap',marginTop:16}}>
                            <button className="btn btn-yellow" onClick={() => { setSelectedId(entry.id); openReponer() }}>＋ Registrar entrada</button>
                            <button className="btn btn-dark" onClick={() => { setSelectedId(entry.id); openAjuste() }}>Ajustar stock</button>
                            <button className="btn btn-ghost" onClick={() => { setSelectedId(entry.id); openMover() }}>⇄ Cambiar de ubicación</button>
                          </div>
                        </div>
                      )
                    })}

                    {/* Acciones globales */}
                    <div className="detail-actions" style={{marginTop:24,paddingTop:20,borderTop:'1px solid #E7E7E3'}}>
                      <button className="btn btn-ghost" onClick={openEntregaFromDetail}>Registrar entrega</button>
                      <button className="btn btn-ghost" onClick={openDevolucionFromDetail}>↩ Devolución</button>
                      <button className="btn btn-ghost btn-full" onClick={() => { setSelectedId(detail.entries[0].id); openEdit() }}>✎ Editar artículo</button>
                    </div>
                  </div>
                </div>
                <div className="card">
                  <div className="card-header"><div className="card-title">Movimientos</div></div>
                  {detail.noMovs && <div className="empty">Sin movimientos registrados.</div>}
                  {detail.movs.map(m => (
                    <div key={m.id} style={{display:'flex',alignItems:'center',gap:12,padding:'13px 20px',borderBottom:'1px solid #F0F0EC'}}>
                      <span style={{width:9,height:9,borderRadius:'50%',flexShrink:0,background:m.tipo==='entrada'?'#2e9b5e':'#C2473D'}} />
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontWeight:600,fontSize:13}}>{m.detalle}</div>
                        <div style={{fontSize:11.5,color:'#8a8a82'}}>{m.fecha} · Talle {m.talle}</div>
                      </div>
                      <div className="mono" style={{fontWeight:700,fontSize:14,color:m.tipo==='entrada'?'#2e9b5e':'#C2473D',flexShrink:0}}>
                        {m.tipo==='entrada'?'+':'−'}{m.qty}
                      </div>
                      <button className="btn-del" onClick={() => m.delId ? askDeleteDelivery(m.delId) : askDeleteMov(m.id)}>✕</button>
                    </div>
                  ))}
                </div>
              </div>
            </>
          )}

          {/* ENTREGAS */}
          {view === 'entregas' && (
            <>
              <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center',marginBottom:12}}>
                <div style={{display:'flex',gap:6,flexWrap:'wrap',flex:1}}>
                  <button className={`chip${delFilterReceptor===''?' active':''}`} onClick={() => setDelFilterReceptor('')}>Todos</button>
                  {deliveryReceptores.map(r => (
                    <button key={r} className={`chip${delFilterReceptor===r?' active':''}`} onClick={() => setDelFilterReceptor(r)}>{r}</button>
                  ))}
                </div>
                <input className="field-input" style={{width:200,flexShrink:0}} placeholder="Buscar integrante…" value={delFilterPersona} onChange={e => setDelFilterPersona(e.target.value)} />
              </div>
            <div className="card table-wrap">
              <div className="card-header">
                <div className="card-title">Historial de entregas</div>
                <div className="card-spacer"/>
                <span style={{fontSize:12.5,color:'#8a8a82'}}>{filteredDeliveryRows.length} de {kpis.entregas} entregas</span>
              </div>
              <div className="table-header del-cols">
                <div>FECHA</div><div>INTEGRANTE / GRUPO</div>
                <div className="del-col-detail">DETALLE</div>
                <div style={{textAlign:'right'}}>UNID.</div><div/>
              </div>
              {filteredDeliveryRows.map(d => (
                <div key={d.id} className="table-row del-cols clickable" onClick={() => setSelectedDeliveryId(d.id)}>
                  <div className="mono" style={{fontSize:12.5,color:'#6a6a62'}}>{d.fecha}</div>
                  <div style={{display:'flex',alignItems:'center',gap:11,minWidth:0}}>
                    <div className="avatar lg">{d.ini}</div>
                    <div style={{minWidth:0}}>
                      <div style={{fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.persona}</div>
                      <div style={{fontSize:11.5,color:'#8a8a82'}}>
                        {d.receptor}
                        {d.paga !== null && d.paga !== undefined && <span style={{marginLeft:6,fontWeight:600,color:d.paga==='si'?'#2e9b5e':'#C2473D'}}>· Paga: {d.paga==='si'?'Sí':'No'}{d.paga==='si'&&d.monto>0?' — $ '+d.monto.toLocaleString('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2}):''}</span>}
                      </div>
                    </div>
                  </div>
                  <div className="del-col-detail" style={{color:'#6a6a62',fontSize:13,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.resumen}</div>
                  <div style={{textAlign:'right',fontWeight:700,fontFamily:'IBM Plex Mono,monospace'}}>{d.totalUd}</div>
                  <div style={{display:'flex',justifyContent:'flex-end',alignItems:'center',gap:6}}>
                    {(() => { const st=d.status||'aceptado'; return st==='pendiente'?<span style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200',borderRadius:5,padding:'2px 7px',fontSize:11,fontWeight:700,whiteSpace:'nowrap'}}>Pendiente</span>:st==='rechazado'?<span style={{background:'#FBEAE8',color:'#C2473D',border:'1px solid #C2473D',borderRadius:5,padding:'2px 7px',fontSize:11,fontWeight:700,whiteSpace:'nowrap'}}>Rechazado</span>:<span style={{background:'#EDF7F2',color:'#2e9b5e',border:'1px solid #2e9b5e',borderRadius:5,padding:'2px 7px',fontSize:11,fontWeight:700,whiteSpace:'nowrap'}}>Aceptado</span> })()}
                    <button className="btn-del" onClick={e => { e.stopPropagation(); askDeleteDelivery(d.id) }}>✕</button>
                  </div>
                </div>
              ))}
              {filteredDeliveryRows.length === 0 && <div className="empty">{delFilterReceptor||delFilterPersona ? 'Sin entregas para este filtro.' : 'Sin entregas registradas.'}</div>}
            </div>
            </>
          )}

          {/* MOVIMIENTOS */}
          {view === 'movimientos' && (
            <>
              <div className="chips">
                {movChipDefs.map(([label]) => (
                  <button key={label} className={`chip${movFilter===label?' active':''}`} onClick={() => setMovFilter(label)}>{label}</button>
                ))}
              </div>
              <div className="card table-wrap">
                <div className="table-header mov-cols">
                  <div>FECHA</div><div>ARTÍCULO / DETALLE</div>
                  <div className="mov-col-tipo">TIPO</div>
                  <div className="mov-col-talle">TALLE</div>
                  <div style={{textAlign:'right'}}>CANT.</div><div/>
                </div>
                {movRows.map(m => {
                  const kind = movKind(m)
                  const kindLabel = {entrega:'Entrega',devolucion:'Devolución',ajuste:'Ajuste',ingreso:'Ingreso'}[kind]
                  const kindClass = {entrega:'exit',devolucion:'entry',ajuste:'yellow',ingreso:'entry'}[kind]
                  return (
                    <div key={m.id} className="table-row mov-cols">
                      <div className="mono" style={{fontSize:12.5,color:'#6a6a62'}}>{m.fecha}</div>
                      <div style={{minWidth:0}}>
                        <div style={{fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{m.name||codeName[m.code]||m.code}</div>
                        <div style={{fontSize:11.5,color:'#8a8a82',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{m.detalle}</div>
                      </div>
                      <div className="mov-col-tipo"><span className={`badge ${kindClass}`}>{kindLabel}</span></div>
                      <div className="mov-col-talle" style={{color:'#6a6a62'}}>{m.talle}</div>
                      <div style={{textAlign:'right'}}>
                        <span className="mono" style={{fontWeight:700,fontSize:14,color:m.tipo==='entrada'?'#2e9b5e':'#C2473D'}}>
                          {m.tipo==='entrada'?'+':'−'}{m.qty}
                        </span>
                      </div>
                      <div style={{display:'flex',justifyContent:'flex-end'}}>
                        <button className="btn-del" style={{width:28,height:28}} onClick={() => m.delId ? askDeleteDelivery(m.delId) : askDeleteMov(m.id)}>✕</button>
                      </div>
                    </div>
                  )
                })}
                {movRows.length === 0 && <div className="empty">Sin movimientos para este filtro.</div>}
              </div>
            </>
          )}

          {/* RECEPTORES */}
          {view === 'usuarios-reg' && (
            <div style={{display:'flex',flexDirection:'column',gap:10,padding:'0 2px'}}>
              {db.users.map(u => (
                <div key={u.username} className="card" style={{padding:'16px 20px',display:'flex',alignItems:'center',gap:14,borderLeft: u.status==='pendiente' ? '3px solid #FFD200' : undefined}}>
                  <div className="avatar" style={{flexShrink:0,width:42,height:42,fontSize:15,opacity:u.status==='pendiente'?.6:1}}>{ini(u.displayName||u.username)}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontWeight:700,fontSize:14,display:'flex',alignItems:'center',gap:8}}>
                      {u.displayName||u.username}
                      {u.status==='pendiente' && <span style={{background:'#FFF8DC',color:'#8a6200',border:'1px solid #e6be00',borderRadius:4,padding:'1px 7px',fontSize:10,fontWeight:700}}>PENDIENTE</span>}
                    </div>
                    <div style={{fontSize:12,color:'#8a8a82',marginTop:2}}>{u.email||u.username}</div>
                    {(u.cargo||u.categoria||u.division) && (
                      <div style={{fontSize:12,color:'#8a8a82',marginTop:2,display:'flex',flexWrap:'wrap',gap:4,alignItems:'center'}}>
                        {u.cargo && <span>{u.cargo}</span>}
                        {u.categoria && <span style={{background:'#F0F0EC',borderRadius:4,padding:'1px 6px',fontSize:11}}>{u.categoria}</span>}
                        {u.division && <span style={{background:'#E8F0FE',borderRadius:4,padding:'1px 6px',fontSize:11,color:'#1a56db'}}>{u.division}</span>}
                      </div>
                    )}
                    {u.telefono && <div style={{fontSize:12,color:'#8a8a82',marginTop:2}}>{u.telefono}</div>}
                    {u.status==='pendiente' && session==='compras' && (
                      <div style={{display:'flex',gap:8,marginTop:10}}>
                        <button onClick={()=>approveUser(u.username)} style={{padding:'5px 14px',borderRadius:5,border:'none',cursor:'pointer',fontWeight:700,fontSize:12,background:'#FFD200',color:'#121212'}}>Aprobar</button>
                        <button onClick={()=>rejectUser(u.username)} style={{padding:'5px 14px',borderRadius:5,border:'1px solid #C2473D',cursor:'pointer',fontWeight:700,fontSize:12,background:'none',color:'#C2473D'}}>Rechazar</button>
                      </div>
                    )}
                  </div>
                  <span style={{background:u.role==='admin'?'#121212':'#EDF7F2',color:u.role==='admin'?'#FFD200':'#2e9b5e',border:'1px solid '+(u.role==='admin'?'#3a3a3a':'#2e9b5e'),borderRadius:5,padding:'2px 8px',fontSize:11,fontWeight:700,flexShrink:0}}>
                    {u.role==='admin'?'Admin':'Receptor'}
                  </span>
                </div>
              ))}
            </div>
          )}

          {view === 'receptores' && (
            <div className="receptor-grid">
              {receptorCards.map(r => (
                <div key={r.name} className="card" style={{padding:20,display:'flex',gap:14,alignItems:'center',cursor:'pointer'}} onClick={() => setSelectedReceptor(r.name)}>
                  <div className="avatar xl">{r.ini}</div>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:700,fontSize:15}}>{r.name}</div>
                    <div style={{fontSize:12.5,color:'#8a8a82',marginTop:3}}>{r.count} entregas · {r.unidades} unidades</div>
                  </div>
                  <span style={{color:'#C8C8C0',fontSize:20}}>›</span>
                </div>
              ))}
            </div>
          )}
          {/* CAMISETAS UTILERÍA */}
          {view === 'utileria' && (
            <div style={{display:'flex',flexDirection:'column',gap:12}}>
              <div style={{display:'flex',gap:8,alignItems:'center',justifyContent:'space-between',flexWrap:'wrap'}}>
                <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                  <button className={`chip${utiFilter===''?' active':''}`} onClick={()=>setUtiFilter('')}>Todas</button>
                  {COMPETICIONES.map(c => (
                    <button key={c} className={`chip${utiFilter===c?' active':''}`} onClick={()=>setUtiFilter(c)}>{c}</button>
                  ))}
                </div>
                <button className="btn btn-dark" style={{flexShrink:0}} onClick={()=>{ setUtiForm({competicion:COMPETICIONES[0],numero:'',jugador:'',talle:'S',modelo:'',estampado:'',parches:'',detalle:'',id:null}); setUtiModal(true) }}>+ Camiseta</button>
              </div>
              <div className="card" style={{overflow:'hidden'}}>
                <div style={{display:'grid',gridTemplateColumns:'52px 1fr 60px 76px 36px',background:'#121212',padding:'9px 16px',gap:8}}>
                  <div style={{fontSize:11,fontWeight:700,color:'#FFD200',letterSpacing:.5}}>NRO.</div>
                  <div style={{fontSize:11,fontWeight:700,color:'#FFD200',letterSpacing:.5}}>JUGADOR / COMPETICIÓN</div>
                  <div style={{fontSize:11,fontWeight:700,color:'#FFD200',letterSpacing:.5}}>TALLE</div>
                  <div></div>
                  <div></div>
                </div>
                {utiFiltered.length === 0
                  ? <div style={{padding:28,textAlign:'center',color:'#8a8a82',fontSize:13}}>No hay camisetas registradas{utiFilter ? ' para esta competición' : ''}.</div>
                  : utiFiltered.map(c => (
                      <div key={c.id} style={{display:'grid',gridTemplateColumns:'52px 1fr 60px 76px 36px',padding:'11px 16px',borderBottom:'1px solid #F0F0EC',alignItems:'center',gap:8}}>
                        <div style={{fontWeight:800,fontSize:18,fontFamily:'IBM Plex Mono,monospace',color:'#1a1a1a'}}>{c.numero}</div>
                        <div>
                          <div style={{fontWeight:600,fontSize:13.5}}>{c.jugador || <span style={{color:'#aaa',fontStyle:'italic',fontWeight:400}}>Sin asignar</span>}</div>
                          <div style={{fontSize:11,color:'#8a8a82',marginTop:2}}>{c.competicion}</div>
                          {(c.modelo||c.estampado||c.parches||c.detalle) && (
                            <div style={{fontSize:11,color:'#aaa',marginTop:2}}>
                              {[c.modelo&&`Modelo: ${c.modelo}`,c.estampado&&`Estampado: ${c.estampado}`,c.parches&&`Parches: ${c.parches}`,c.detalle&&`Detalle: ${c.detalle}`].filter(Boolean).join(' · ')}
                            </div>
                          )}
                        </div>
                        <div style={{fontSize:13,fontWeight:600}}>{c.talle}</div>
                        <button className="btn btn-ghost" style={{padding:'4px 10px',fontSize:12}} onClick={()=>{setUtiForm({...c}); setUtiModal(true)}}>Editar</button>
                        <button onClick={()=>deleteUti(c.id)} style={{background:'none',border:'none',cursor:'pointer',fontSize:18,color:'#C2473D',padding:'0 4px',lineHeight:1}}>×</button>
                      </div>
                    ))
                }
              </div>
            </div>
          )}
        </div>
      </div>

      {/* ===== MODALES ===== */}

      {/* Modal: Detalle de entrega */}
      {selectedDeliveryId && (() => {
        const d = deliveryRows.find(x => x.id === selectedDeliveryId)
        if (!d) return null
        const st = d.status || 'aceptado'
        const stStyle = st==='pendiente'
          ? {background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200'}
          : st==='rechazado'
          ? {background:'#FBEAE8',color:'#C2473D',border:'1px solid #C2473D'}
          : {background:'#EDF7F2',color:'#2e9b5e',border:'1px solid #2e9b5e'}
        const stLabel = st==='pendiente'?'Pendiente':st==='rechazado'?'Rechazado':'Aceptado'
        return (
          <div className="modal-backdrop" onClick={() => setSelectedDeliveryId(null)}>
            <div className="modal modal-md" onClick={e => e.stopPropagation()}>
              <div className="modal-header">
                <div>
                  <div className="modal-title">{d.persona}</div>
                  <div style={{fontSize:12.5,color:'#8a8a82',marginTop:2}}>{d.receptor} · {d.fecha}</div>
                </div>
                <span style={{...stStyle,borderRadius:5,padding:'3px 9px',fontSize:11,fontWeight:700,marginLeft:'auto',marginRight:12,whiteSpace:'nowrap'}}>{stLabel}</span>
                <button className="modal-close" onClick={() => setSelectedDeliveryId(null)}>×</button>
              </div>
              <div className="modal-body" style={{padding:0}}>
                <div style={{display:'grid',gridTemplateColumns:'1fr 60px 55px',background:'#121212',padding:'9px 20px'}}>
                  <div style={{fontSize:11,fontWeight:700,color:'#FFD200',letterSpacing:'.04em'}}>PRENDA</div>
                  <div style={{fontSize:11,fontWeight:700,color:'#FFD200',letterSpacing:'.04em',textAlign:'center'}}>TALLE</div>
                  <div style={{fontSize:11,fontWeight:700,color:'#FFD200',letterSpacing:'.04em',textAlign:'right'}}>CANT.</div>
                </div>
                {d.lines.map((l, i) => (
                  <div key={i} style={{display:'grid',gridTemplateColumns:'1fr 60px 55px',padding:'11px 20px',borderBottom:'1px solid #F0F0EC',alignItems:'center'}}>
                    <div>
                      <div style={{fontWeight:600,fontSize:13.5}}>{l.name || l.code}</div>
                      {l.ubic && <div style={{fontSize:11.5,color:'#8a8a82',marginTop:2}}>Ubic. {l.ubic}</div>}
                    </div>
                    <div style={{textAlign:'center',fontWeight:700,fontSize:13}}>{l.talle}</div>
                    <div style={{textAlign:'right',fontWeight:700,fontFamily:'IBM Plex Mono,monospace',fontSize:14}}>{l.qty}</div>
                  </div>
                ))}
                <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'12px 20px',background:'#FAFAF8',borderTop:'2px solid #E7E7E3'}}>
                  <span style={{fontSize:13,color:'#6a6a62',fontWeight:600}}>Total unidades</span>
                  <span style={{fontWeight:800,fontSize:16,fontFamily:'IBM Plex Mono,monospace'}}>{d.totalUd}</span>
                </div>
                {d.paga === 'si' && d.monto > 0 && (
                  <div style={{padding:'10px 20px',background:'#F0FAF4',borderTop:'1px solid #b6e4c8',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                    <span style={{fontSize:13,color:'#1a5c33',fontWeight:600}}>Total a cobrar</span>
                    <span style={{fontWeight:800,fontSize:15,color:'#1a5c33'}}>$ {d.monto.toLocaleString('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2})}</span>
                  </div>
                )}
              </div>
              <div className="modal-footer">
                <button className="btn btn-ghost" onClick={() => setSelectedDeliveryId(null)}>Cerrar</button>
                <button className="btn btn-red" onClick={() => { setSelectedDeliveryId(null); askDeleteDelivery(d.id) }}>Eliminar entrega</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* Modal: Camiseta Utilería */}
      {utiModal && (
        <div className="modal-backdrop" onClick={()=>setUtiModal(false)}>
          <div className="modal modal-sm" onClick={e=>e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">{utiForm.id !== null ? 'Editar camiseta' : 'Nueva camiseta'}</div>
              <button className="modal-close" onClick={()=>setUtiModal(false)}>×</button>
            </div>
            <div className="modal-body" style={{display:'flex',flexDirection:'column',gap:12}}>
              <div className="form-group">
                <label className="field-label">Competición</label>
                <select className="field-input" value={utiForm.competicion} onChange={e=>setUtiForm(p=>({...p,competicion:e.target.value}))}>
                  {COMPETICIONES.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div className="form-cols-2">
                <div className="form-group">
                  <label className="field-label">Número</label>
                  <input className="field-input" value={utiForm.numero} onChange={e=>setUtiForm(p=>({...p,numero:e.target.value}))} placeholder="10" />
                </div>
                <div className="form-group">
                  <label className="field-label">Talle</label>
                  <select className="field-input" value={utiForm.talle} onChange={e=>setUtiForm(p=>({...p,talle:e.target.value}))}>
                    {[...TALLES_ADULTO,...TALLES_NINO].map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
              </div>
              <div className="form-group">
                <label className="field-label">Jugador / Asignado</label>
                <input className="field-input" value={utiForm.jugador} onChange={e=>setUtiForm(p=>({...p,jugador:e.target.value}))} placeholder="Nombre del jugador" />
              </div>
              <div className="form-cols-2">
                <div className="form-group">
                  <label className="field-label">Modelo</label>
                  <select className="field-input" value={utiForm.modelo} onChange={e=>setUtiForm(p=>({...p,modelo:e.target.value}))}>
                    <option value="">Seleccionar…</option>
                    <option value="TRADICIONAL">TRADICIONAL</option>
                    <option value="GRIS">GRIS</option>
                  </select>
                </div>
                <div className="form-group">
                  <label className="field-label">Estampado</label>
                  <input className="field-input" value={utiForm.estampado} onChange={e=>setUtiForm(p=>({...p,estampado:e.target.value}))} placeholder="" />
                </div>
              </div>
              <div className="form-cols-2">
                <div className="form-group">
                  <label className="field-label">Parches</label>
                  <input className="field-input" value={utiForm.parches} onChange={e=>setUtiForm(p=>({...p,parches:e.target.value}))} placeholder="" />
                </div>
                <div className="form-group">
                  <label className="field-label">Detalle</label>
                  <input className="field-input" value={utiForm.detalle} onChange={e=>setUtiForm(p=>({...p,detalle:e.target.value}))} placeholder="" />
                </div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={()=>setUtiModal(false)}>Cancelar</button>
              <button className="btn btn-dark" onClick={saveUti}>Guardar</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Detalle de receptor */}
      {selectedReceptor && (() => {
        const rData = receptorCards.find(r => r.name === selectedReceptor)
        const rDeliveries = deliveryRows.filter(d => d.receptor === selectedReceptor).sort((a,b) => b.fecha.localeCompare(a.fecha))
        return (
          <div className="modal-backdrop" onClick={() => setSelectedReceptor(null)}>
            <div className="modal modal-md" onClick={e => e.stopPropagation()}>
              <div className="modal-header">
                <div style={{display:'flex',gap:12,alignItems:'center'}}>
                  <div className="avatar xl">{rData?.ini}</div>
                  <div>
                    <div className="modal-title">{selectedReceptor}</div>
                    <div style={{fontSize:12.5,color:'#8a8a82',marginTop:2}}>{rData?.count} entregas · {rData?.unidades} unidades totales</div>
                  </div>
                </div>
                <button className="modal-close" onClick={() => setSelectedReceptor(null)}>×</button>
              </div>
              <div className="modal-body" style={{padding:0,maxHeight:'65vh',overflowY:'auto'}}>
                {rDeliveries.length === 0
                  ? <div style={{padding:24,textAlign:'center',color:'#8a8a82'}}>Sin entregas registradas.</div>
                  : rDeliveries.map(d => {
                      const st = d.status || 'aceptado'
                      const stStyle = st==='pendiente'
                        ? {background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200'}
                        : st==='rechazado'
                        ? {background:'#FBEAE8',color:'#C2473D',border:'1px solid #C2473D'}
                        : {background:'#EDF7F2',color:'#2e9b5e',border:'1px solid #2e9b5e'}
                      const stLabel = st==='pendiente'?'Pendiente':st==='rechazado'?'Rechazado':'Aceptado'
                      return (
                        <div key={d.id} className="clickable"
                          style={{display:'flex',alignItems:'center',gap:10,padding:'12px 20px',borderBottom:'1px solid #F0F0EC'}}
                          onClick={() => { setSelectedReceptor(null); setSelectedDeliveryId(d.id) }}>
                          <div style={{flex:1}}>
                            <div style={{fontWeight:600,fontSize:13.5}}>{d.persona}</div>
                            <div style={{fontSize:11.5,color:'#8a8a82',marginTop:2}}>{d.fecha} · {d.totalUd} u.</div>
                          </div>
                          <span style={{...stStyle,borderRadius:5,padding:'3px 9px',fontSize:11,fontWeight:700,whiteSpace:'nowrap'}}>{stLabel}</span>
                          <span style={{color:'#C8C8C0',fontSize:20,marginLeft:4}}>›</span>
                        </div>
                      )
                    })
                }
              </div>
              <div className="modal-footer">
                <button className="btn btn-ghost" onClick={() => setSelectedReceptor(null)}>Cerrar</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* Modal: Entrega / Devolución */}
      {modal === 'entrega' && (
        <div className="modal-backdrop" onClick={closeModal}>
          <div className="modal modal-lg" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">{ndIsDev ? 'Registrar devolución' : 'Registrar entrega'}</div>
              <button className="modal-close" onClick={closeModal}>×</button>
            </div>
            <div className="modal-body" style={{display:'flex',flexDirection:'column',gap:14}}>
              <div className="form-group">
                <label className="field-label">{ndIsDev ? 'Integrante que devuelve' : 'Integrante que recibe'}</label>
                <input className="field-input" value={nd.persona} onChange={e => setNd(p=>({...p,persona:e.target.value}))} placeholder="Ej. Maximiliano Olivera" />
              </div>
              <div className="form-group">
                <label className="field-label">Grupo / Plantel</label>
                <select className="field-input" value={nd.receptor} onChange={e => setNd(p=>({...p,receptor:e.target.value,paga:null}))}>
                  <option value="">Seleccionar grupo…</option>
                  {RECEPTORES.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
              </div>
              {!ndIsDev && receptorUsers.length > 0 && (
                <div className="form-group">
                  <label className="field-label">Enviar a usuario registrado <span style={{fontSize:11,color:'#8a8a82',fontWeight:400}}>(opcional)</span></label>
                  <select className="field-input" value={nd.toUser} onChange={e => {
                    const u = receptorUsers.find(x => x.username === e.target.value)
                    const norm = s => (s||'').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'')
                    const matchedReceptor = u ? (RECEPTORES.find(r => norm(r) === norm(u.categoria)) || '') : ''
                    setNd(p=>({...p, toUser:e.target.value, persona: u ? u.displayName : p.persona, receptor: matchedReceptor || p.receptor}))
                  }}>
                    <option value="">Sin usuario específico</option>
                    {receptorUsers.map(u => <option key={u.username} value={u.username}>{u.displayName} ({u.username})</option>)}
                  </select>
                  {nd.toUser && <div style={{marginTop:6,fontSize:12,color:'#7a5800',background:'#FFF8D6',border:'1px solid #FFD200',borderRadius:6,padding:'6px 10px'}}>La entrega quedará pendiente de confirmación por el receptor.</div>}
                </div>
              )}
              {nd.receptor === 'Protocolo' && !ndIsDev && (
                <div className="form-group">
                  <label className="field-label">¿Paga?</label>
                  <div style={{display:'flex',gap:8}}>
                    {[['si','SÍ'],['no','NO']].map(([v,label]) => (
                      <button key={v} style={{flex:1,padding:'7px 0',borderRadius:6,border:'1px solid',cursor:'pointer',fontWeight:700,fontSize:13,
                        background:nd.paga===v?'#FFD200':'#F5F5F0',
                        borderColor:nd.paga===v?'#e6be00':'#E0E0DA',
                        color:nd.paga===v?'#121212':'#8a8a82'}}
                        onClick={() => setNd(p=>({...p,paga:v}))}>
                        {label}
                      </button>
                    ))}
                  </div>
                  {nd.paga === 'si' && nd.lines.length > 0 && (
                    <div style={{marginTop:10,padding:'8px 12px',background:'#F0FAF4',border:'1px solid #b6e4c8',borderRadius:6,fontSize:13,color:'#1a5c33'}}>
                      Total a cobrar: <b style={{fontSize:15}}>$ {ndMonto.toLocaleString('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2})}</b>
                    </div>
                  )}
                </div>
              )}
              <div style={{background:'#FAFAF8',border:'1px solid #ECECE8',borderRadius:8,padding:16}}>
                <div style={{fontSize:12.5,fontWeight:700,color:'#4a4a42',marginBottom:11}}>Agregar artículo</div>
                <div style={{display:'flex',flexDirection:'column',gap:10}}>
                  <div style={{position:'relative'}}>
                    <input
                      className="field-input"
                      placeholder="Buscar prenda por nombre…"
                      autoComplete="off"
                      value={nd.cCode ? (ndA ? ndA.name : nd.cCode) : nd.cSearch}
                      onChange={e => setNd(p=>({...p, cSearch:e.target.value, cCode:'', cTalle:'', cQty:''}))}
                      onBlur={() => setTimeout(() => setNd(p => p.cCode ? p : {...p, cSearch:''}), 150)}
                    />
                    {!nd.cCode && nd.cSearch && (
                      <div onMouseDown={e => e.preventDefault()} style={{position:'absolute',top:'calc(100% + 4px)',left:0,right:0,background:'#fff',border:'1px solid #ECECE8',borderRadius:8,zIndex:200,maxHeight:220,overflowY:'auto',boxShadow:'0 4px 16px rgba(0,0,0,.12)'}}>
                        {(() => {
                          const term = nd.cSearch.toLowerCase()
                          const seen = new Set()
                          const unique = articles.filter(a => {
                            if(seen.has(a.code)) return false
                            seen.add(a.code)
                            return a.name?.toLowerCase().includes(term) || a.code?.toLowerCase().includes(term)
                          })
                          return unique.length === 0
                            ? <div style={{padding:'10px 14px',fontSize:13,color:'#8a8a82'}}>Sin resultados</div>
                            : unique.map(a => (
                              <div key={a.code} style={{padding:'9px 14px',cursor:'pointer',borderBottom:'1px solid #F2F2EE',fontSize:13}} onClick={() => setNd(p=>({...p, cCode:a.code, cSearch:'', cUbic:'', cTalle:'', cQty:''}))}>
                                <span style={{fontWeight:600}}>{a.name}</span>
                                <span style={{color:'#8a8a82',fontSize:11.5,marginLeft:8}}>{a.code}</span>
                              </div>
                            ))
                        })()}
                      </div>
                    )}
                  </div>
                  {nd.cCode && ndHasMultiUbic && (
                    <div>
                      <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',letterSpacing:'.04em',marginBottom:6}}>UBICACIÓN</div>
                      <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                        {ndUbics.map(ubic => (
                          <button key={ubic} onClick={() => setNd(p=>({...p, cUbic:ubic, cTalle:'', cQty:''}))}
                            style={{padding:'5px 14px',borderRadius:5,border:'1px solid',cursor:'pointer',fontWeight:700,fontSize:12.5,
                              background: nd.cUbic===ubic ? '#121212' : '#F5F5F0',
                              color: nd.cUbic===ubic ? '#FFD200' : '#5a5a52',
                              borderColor: nd.cUbic===ubic ? '#121212' : '#E0E0DA'}}>
                            {ubic}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                  <div className="nd-tallerow">
                    <select className="field-input" value={nd.cTalle} onChange={e => setNd(p=>({...p,cTalle:e.target.value}))}
                      disabled={ndHasMultiUbic && !nd.cUbic}>
                      <option value="">{ndHasMultiUbic && !nd.cUbic ? 'Elegí ubicación primero' : 'Talle…'}</option>
                      {ndTalleOptions.map(t => <option key={t.value} value={t.value}>{t.label}</option>)}
                    </select>
                    <input type="number" min="1" className="field-input" value={nd.cQty} onChange={e => setNd(p=>({...p,cQty:e.target.value}))} placeholder="Cantidad" />
                    <button className="btn btn-dark" onClick={ndAddLine}>Agregar</button>
                  </div>
                  {stockHint && <div style={{fontSize:11.5,color:'#8a8a82'}}>{stockHint}</div>}
                </div>
              </div>
              {nd.lines.length > 0 && (
                <div style={{border:'1px solid #ECECE8',borderRadius:8,overflow:'hidden'}}>
                  {nd.lines.map((l,i) => (
                    <div key={i} style={{display:'flex',alignItems:'center',gap:12,padding:'11px 14px',borderBottom:'1px solid #F2F2EE'}}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontWeight:600,fontSize:13}}>{codeName[l.code]||l.code}</div>
                        <div style={{fontSize:11.5,color:'#8a8a82'}}>Talle {l.talle}</div>
                      </div>
                      <div className="mono" style={{fontWeight:600,fontSize:13}}>×{l.qty}</div>
                      <button style={{background:'none',border:'none',color:'#C2473D',fontSize:18,cursor:'pointer',padding:0,lineHeight:1}} onClick={() => setNd(p=>({...p,lines:p.lines.filter((_,j)=>j!==i)}))}>×</button>
                    </div>
                  ))}
                </div>
              )}
            </div>
            <div className="modal-footer">
              <span style={{fontSize:13,color:'#8a8a82'}}>Total: <b style={{color:'#1a1a1a'}}>{ndTotal}</b> unidades</span>
              <div style={{flex:1}}/>
              <button className="btn btn-ghost" onClick={closeModal}>Cancelar</button>
              <button className="btn" style={{background:ndOk?'#FFD200':'#EDE9D2',color:ndOk?'#121212':'#a89e6a',cursor:ndOk?'pointer':'not-allowed'}} onClick={ndConfirm}>
                {ndIsDev ? 'Confirmar devolución' : 'Confirmar entrega'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Nuevo artículo */}
      {modal === 'articulo' && (
        <div className="modal-backdrop" onClick={closeModal}>
          <div className="modal modal-lg" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">Nuevo artículo</div>
              <button className="modal-close" onClick={closeModal}>×</button>
            </div>
            <div className="modal-body" style={{display:'flex',flexDirection:'column',gap:14}}>
              <div className="form-cols-2" style={{gap:12}}>
                <div className="form-group">
                  <label className="field-label">Código (SKU)</label>
                  <input className="field-input mono" value={na.code} onChange={e => setNa(p=>({...p,code:e.target.value.toUpperCase()}))} placeholder="CAM-XXX-26" />
                </div>
                <div className="form-group">
                  <label className="field-label">Categoría</label>
                  <select className="field-input" value={na.cat} onChange={e => setNa(p=>({...p,cat:e.target.value}))}>
                    {CATEGORIAS.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              </div>
              <div className="form-group">
                <label className="field-label">Nombre del artículo</label>
                <input className="field-input" value={na.name} onChange={e => setNa(p=>({...p,name:e.target.value}))} placeholder="Ej. Camiseta Titular 2026" />
              </div>
              <div className="form-group">
                <label className="field-label">Precio Tienda (Socio) <span style={{fontSize:11,color:'#8a8a82',fontWeight:400}}>(opcional)</span></label>
                <input type="number" min="0" step="0.01" className="field-input" value={na.precio} onChange={e => setNa(p=>({...p,precio:e.target.value}))} placeholder="0.00" />
              </div>
              <div className="form-group">
                <label className="field-label">Talles</label>
                <div style={{display:'flex',gap:8,marginBottom:10}}>
                  {['adulto','nino'].map(t => (
                    <button key={t} style={{flex:1,padding:'7px 0',borderRadius:6,border:'1px solid',cursor:'pointer',fontWeight:700,fontSize:13,
                      background:na.tipo===t?'#FFD200':'#F5F5F0',
                      borderColor:na.tipo===t?'#b89900':'#E0E0DA',
                      color:na.tipo===t?'#121212':'#8a8a82'}}
                      onClick={() => setNa(p=>({...p,tipo:t,tallesArr:[],tallesMins:{},tallesQty:{}}))}>
                      {t==='adulto'?'ADULTO':'NIÑO'}
                    </button>
                  ))}
                </div>
                <div className="talle-grid">
                  {(na.tipo==='nino' ? TALLES_NINO : TALLES_ADULTO).map(t => (
                    <button key={t} className={`talle-btn${na.tallesArr.includes(t)?' active':''}`} onClick={() => naToggleTalle(t)}>{t}</button>
                  ))}
                </div>
                {na.tallesArr.length > 0 && (
                  <div style={{marginTop:14,border:'1px solid #E7E7E3',borderRadius:8,overflow:'hidden'}}>
                    <div style={{display:'grid',gridTemplateColumns:'56px 1fr 1fr',background:'#FAFAF8',padding:'8px 12px',borderBottom:'1px solid #E7E7E3'}}>
                      <div style={{fontSize:11,fontWeight:700,color:'#8a8a82'}}>TALLE</div>
                      <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',textAlign:'center'}}>STOCK INICIAL</div>
                      <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',textAlign:'center'}}>MÍN.</div>
                    </div>
                    {na.tallesArr.map(t => (
                      <div key={t} style={{display:'grid',gridTemplateColumns:'56px 1fr 1fr',gap:8,padding:'8px 12px',borderBottom:'1px solid #F0F0EC',alignItems:'center'}}>
                        <span style={{fontWeight:700,fontSize:13.5}}>{t}</span>
                        <input type="number" min="0" className="field-input" style={{height:36,textAlign:'center',padding:'0 8px'}}
                          value={na.tallesQty[t]||''} onChange={e => setNa(p=>({...p,tallesQty:{...p.tallesQty,[t]:parseInt(e.target.value,10)||0}}))} placeholder="0" />
                        <input type="number" min="0" className="field-input" style={{height:36,textAlign:'center',padding:'0 8px'}}
                          value={na.tallesMins[t]||''} onChange={e => setNa(p=>({...p,tallesMins:{...p.tallesMins,[t]:parseInt(e.target.value,10)||0}}))} placeholder="0" />
                      </div>
                    ))}
                    <div style={{padding:'8px 14px',background:'#FAFAF8',textAlign:'right',fontSize:12.5,color:'#8a8a82'}}>
                      Total: <b style={{color:'#1a1a1a',fontFamily:'IBM Plex Mono,monospace'}}>{Object.values(na.tallesQty).reduce((a,b)=>a+b,0)}</b> u.
                    </div>
                  </div>
                )}
              </div>
              <div className="form-group">
                <label className="field-label">Ubicación en depósito</label>
                <div className="form-cols-2" style={{gap:12}}>
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <span style={{fontSize:12,color:'#8a8a82',whiteSpace:'nowrap'}}>Estantería</span>
                    <select className="field-input" style={{flex:1}} value={na.estante} onChange={e => setNa(p=>({...p,estante:e.target.value}))}>
                      {ESTANTES.map(o => <option key={o} value={o}>{o}</option>)}
                    </select>
                  </div>
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <span style={{fontSize:12,color:'#8a8a82',whiteSpace:'nowrap'}}>Altura</span>
                    <select className="field-input" style={{flex:1}} value={na.altura} onChange={e => setNa(p=>({...p,altura:e.target.value}))}>
                      {ALTURAS.map(o => <option key={o} value={o}>{o}</option>)}
                    </select>
                  </div>
                </div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={closeModal}>Cancelar</button>
              <button className="btn" style={{background:na.code&&na.name?'#121212':'#E7E7E3',color:na.code&&na.name?'#fff':'#a0a098',cursor:na.code&&na.name?'pointer':'not-allowed'}} onClick={naConfirm}>
                Crear artículo
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Registrar entrada */}
      {modal === 'reponer' && selA && (
        <div className="modal-backdrop" onClick={closeModal}>
          <div className="modal modal-sm" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">Registrar entrada</div>
              <button className="modal-close" onClick={closeModal}>×</button>
            </div>
            <div className="modal-body">
              <div style={{fontSize:13,color:'#8a8a82',marginBottom:4}}>{selA.name} <span className="mono">· {selA.code}</span></div>
              <div style={{fontSize:12,color:'#9a7d00',background:'#FBF7E3',padding:'6px 10px',borderRadius:6,marginBottom:14}}>Ubicación: <b>{selA.ubic||'—'}</b></div>
              <div style={{border:'1px solid #E7E7E3',borderRadius:8,overflow:'hidden'}}>
                <div style={{display:'grid',gridTemplateColumns:'56px 1fr 1fr',background:'#FAFAF8',padding:'8px 12px',borderBottom:'1px solid #E7E7E3'}}>
                  <div style={{fontSize:11,fontWeight:700,color:'#8a8a82'}}>TALLE</div>
                  <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',textAlign:'center'}}>STOCK ACTUAL</div>
                  <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',textAlign:'center'}}>AGREGAR</div>
                </div>
                {selA.sizes.map(s => (
                  <div key={s.talle} style={{display:'grid',gridTemplateColumns:'56px 1fr 1fr',gap:8,padding:'8px 12px',borderBottom:'1px solid #F0F0EC',alignItems:'center'}}>
                    <div style={{fontWeight:700,fontSize:13}}>{s.talle}</div>
                    <div style={{textAlign:'center',fontSize:13,color:'#8a8a82'}}>{s.qty}</div>
                    <input type="number" min="0" className="field-input" style={{textAlign:'center',padding:'4px 8px'}}
                      value={rep.qtys[s.talle]||''} onChange={e => setRep(p=>({...p,qtys:{...p.qtys,[s.talle]:e.target.value}}))} placeholder="0" />
                  </div>
                ))}
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={closeModal}>Cancelar</button>
              <button className="btn btn-yellow" onClick={repConfirm}>Sumar al stock</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Ajuste de stock */}
      {modal === 'ajuste' && selA && (
        <div className="modal-backdrop" onClick={closeModal}>
          <div className="modal modal-sm" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">Ajuste de stock</div>
              <button className="modal-close" onClick={closeModal}>×</button>
            </div>
            <div className="modal-body">
              <div style={{fontSize:13,color:'#8a8a82',marginBottom:4}}>{selA.name} <span className="mono">· {selA.code}</span></div>
              <div style={{fontSize:12,color:'#9a7d00',background:'#FBF7E3',padding:'6px 10px',borderRadius:6,marginBottom:10}}>Ubicación: <b>{selA.ubic||'—'}</b></div>
              <div style={{fontSize:12,color:'#9a7d00',background:'#FBF7E3',padding:'8px 12px',borderRadius:6,marginBottom:16}}>
                Corrección por recuento: ingresá la cantidad física real contada. El sistema registra la diferencia.
              </div>
              <div className="form-cols-2" style={{gap:12}}>
                <div className="form-group">
                  <label className="field-label">Talle</label>
                  <select className="field-input" value={aj.talle} onChange={e => setAj(p=>({...p,talle:e.target.value}))}>
                    <option value="">Talle…</option>
                    {ajTalleOptions.map(t => <option key={t.value} value={t.value}>{t.label}</option>)}
                  </select>
                </div>
                <div className="form-group">
                  <label className="field-label">Cantidad real contada</label>
                  <input type="number" min="0" className="field-input" value={aj.cantidad} onChange={e => setAj(p=>({...p,cantidad:e.target.value}))} placeholder="0" />
                </div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={closeModal}>Cancelar</button>
              <button className="btn btn-dark" onClick={ajConfirm}>Aplicar ajuste</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Editar artículo */}
      {modal === 'edit' && editing && (
        <div className="modal-backdrop" onClick={closeModal}>
          <div className="modal modal-md" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">Editar artículo</div>
              <button className="modal-close" onClick={closeModal}>×</button>
            </div>
            <div className="modal-body" style={{display:'flex',flexDirection:'column',gap:14}}>
              <div className="form-cols-2" style={{gap:12}}>
                <div className="form-group">
                  <label className="field-label">Código (SKU)</label>
                  <input className="field-input mono" value={editing.code} onChange={e => setEditing(p=>({...p,code:e.target.value}))} />
                </div>
                <div className="form-group">
                  <label className="field-label">Categoría</label>
                  <select className="field-input" value={editing.cat} onChange={e => setEditing(p=>({...p,cat:e.target.value}))}>
                    {CATEGORIAS.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              </div>
              <div className="form-group">
                <label className="field-label">Nombre del artículo</label>
                <input className="field-input" value={editing.name} onChange={e => setEditing(p=>({...p,name:e.target.value}))} />
              </div>
              <div className="form-cols-2" style={{gap:12}}>
                <div className="form-group">
                  <label className="field-label">Ubicación <span style={{fontSize:11,color:'#8a8a82',fontWeight:400}}>(ej. 3B, 0O)</span></label>
                  <input className="field-input mono" value={editing.ubic} onChange={e => setEditing(p=>({...p,ubic:e.target.value}))} placeholder="1A" />
                </div>
                <div className="form-group">
                  <label className="field-label">Precio Tienda (Socio)</label>
                  <input type="number" min="0" step="0.01" className="field-input" value={editing.precio} onChange={e => setEditing(p=>({...p,precio:e.target.value}))} placeholder="0.00" />
                </div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={closeModal}>Cancelar</button>
              <button className="btn btn-yellow" onClick={saveEdit}>Guardar cambios</button>
            </div>
          </div>
        </div>
      )}

      {/* Confirm dialog */}
      {confirm && (
        <div className="modal-backdrop" style={{zIndex:55}} onClick={() => setConfirm(null)}>
          <div className="modal modal-sm" onClick={e => e.stopPropagation()}>
            <div className="modal-body">
              <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:12}}>
                <span style={{width:38,height:38,borderRadius:'50%',background:'#FBEAE8',color:'#C2473D',display:'flex',alignItems:'center',justifyContent:'center',fontSize:18,flexShrink:0}}>✕</span>
                <div style={{fontWeight:800,fontSize:17}}>{confirm.title}</div>
              </div>
              <div style={{fontSize:13.5,color:'#6a6a62',lineHeight:1.5}}>{confirm.msg}</div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={() => setConfirm(null)}>Cancelar</button>
              <button className="btn btn-red" onClick={confirmYes}>Eliminar</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Cambiar de ubicación */}
      {modal === 'mover' && selA && (
        <div className="modal-backdrop" onClick={closeModal}>
          <div className="modal modal-sm" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">Cambiar de ubicación</div>
              <button className="modal-close" onClick={closeModal}>×</button>
            </div>
            <div className="modal-body" style={{display:'flex',flexDirection:'column',gap:14}}>
              <div style={{fontSize:13,color:'#8a8a82'}}>{selA.name} <span className="mono">· {selA.code}</span></div>
              <div style={{fontSize:12,background:'#F5F5F0',border:'1px solid #E0E0DA',borderRadius:6,padding:'8px 12px',color:'#6a6a62'}}>
                Ubicación actual: <b style={{color:'#1a1a1a'}}>{selA.ubic||'—'}</b>
              </div>
              <div className="form-group">
                <label className="field-label">Cantidad a mover por talle</label>
                <div style={{border:'1px solid #E7E7E3',borderRadius:8,overflow:'hidden'}}>
                  <div style={{display:'grid',gridTemplateColumns:'56px 1fr 84px',background:'#FAFAF8',padding:'7px 12px',borderBottom:'1px solid #E7E7E3'}}>
                    <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',letterSpacing:.5}}>TALLE</div>
                    <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',letterSpacing:.5}}>STOCK</div>
                    <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',letterSpacing:.5,textAlign:'center'}}>MOVER</div>
                  </div>
                  {selA.sizes.map(s => (
                    <div key={s.talle} style={{display:'grid',gridTemplateColumns:'56px 1fr 84px',gap:6,padding:'8px 12px',borderBottom:'1px solid #F0F0EC',alignItems:'center'}}>
                      <div style={{fontWeight:700,fontSize:13.5}}>{s.talle}</div>
                      <div style={{fontSize:13,color:'#6a6a62'}}>{s.qty} u.</div>
                      <input type="number" min="0" max={s.qty} className="field-input"
                        style={{height:34,textAlign:'center',padding:'0 6px',fontSize:13}}
                        value={mv.qtys[s.talle]||''}
                        onChange={e => setMv(p => ({...p, qtys:{...p.qtys,[s.talle]:e.target.value}}))}
                        placeholder="0"
                      />
                    </div>
                  ))}
                </div>
              </div>
              <div className="form-group">
                <label className="field-label">Ubicación destino</label>
                <div style={{display:'flex',alignItems:'center',gap:10}}>
                  <select className="field-input" style={{flex:1}} value={mv.estante} onChange={e => setMv(p=>({...p,estante:e.target.value}))}>
                    {ESTANTES.map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                  <span style={{fontSize:12,color:'#8a8a82',whiteSpace:'nowrap'}}>Altura</span>
                  <select className="field-input" style={{flex:1}} value={mv.altura} onChange={e => setMv(p=>({...p,altura:e.target.value}))}>
                    {ALTURAS.map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </div>
                <div style={{marginTop:6,fontSize:12,color:'#8a8a82'}}>Destino: <b style={{color:'#1a1a1a'}}>{mv.estante}{mv.altura}</b></div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={closeModal}>Cancelar</button>
              <button className="btn btn-dark" onClick={mvConfirm}>Confirmar movimiento</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Gestión de usuarios */}
      {modal === 'usuarios' && (
        <div className="modal-backdrop" onClick={closeModal}>
          <div className="modal modal-md" onClick={e=>e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">Gestión de usuarios</div>
              <button className="modal-close" onClick={closeModal}>✕</button>
            </div>
            <div className="modal-body" style={{display:'flex',flexDirection:'column',gap:16}}>
              <div>
                <div style={{fontSize:12,fontWeight:700,color:'#8a8a82',letterSpacing:'.04em',marginBottom:10}}>USUARIOS ACTIVOS</div>
                {userMgmt.list.map(u => (
                  <div key={u.username} style={{padding:'12px 0',borderBottom:'1px solid #F0F0EC'}}>
                    <div style={{display:'flex',alignItems:'center',gap:12}}>
                      <div className="avatar" style={{flexShrink:0}}>{ini(u.displayName||u.username)}</div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontWeight:700,fontSize:13.5}}>{u.displayName||u.username}</div>
                        <div style={{fontSize:11.5,color:'#8a8a82'}}>{u.email||u.username}</div>
                        {(u.cargo||u.categoria||u.division) && (
                          <div style={{fontSize:11.5,color:'#8a8a82'}}>
                            {u.cargo||''}
                            {u.categoria ? <span style={{marginLeft:6,background:'#F0F0EC',borderRadius:4,padding:'1px 6px',fontSize:11}}>{u.categoria}</span> : null}
                            {u.division ? <span style={{marginLeft:4,background:'#E8F0FE',borderRadius:4,padding:'1px 6px',fontSize:11,color:'#1a56db'}}>{u.division}</span> : null}
                          </div>
                        )}
                        {u.telefono && <div style={{fontSize:11.5,color:'#8a8a82'}}>{u.telefono}</div>}
                      </div>
                      <span style={{background:u.role==='admin'?'#121212':'#EDF7F2',color:u.role==='admin'?'#FFD200':'#2e9b5e',border:'1px solid '+(u.role==='admin'?'#3a3a3a':'#2e9b5e'),borderRadius:5,padding:'2px 8px',fontSize:11,fontWeight:700,flexShrink:0}}>{u.role==='admin'?'Admin':'Receptor'}</span>
                      {u.username === session && <span className="badge gray">Vos</span>}
                      {u.username !== session && <button className="btn-del" onClick={()=>deleteUser(u.username)}>✕</button>}
                    </div>
                    {session === 'compras' && u.username !== 'compras' && (
                      <div style={{display:'flex',gap:6,marginTop:8,paddingLeft:42}}>
                        {[['admin','Administrador'],['receptor','Receptor']].map(([v,label]) => (
                          <button key={v} onClick={()=>{
                            const list = userMgmt.list.map(x => x.username===u.username?{...x,role:v}:x)
                            saveUsers(list)
                            setUserMgmt(p=>({...p,list}))
                          }} style={{padding:'4px 12px',borderRadius:5,border:'1px solid',cursor:'pointer',fontWeight:700,fontSize:11.5,background:u.role===v?'#FFD200':'#F5F5F0',borderColor:u.role===v?'#e6be00':'#E0E0DA',color:u.role===v?'#121212':'#8a8a82'}}>
                            {label}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                ))}
              </div>
              <div style={{borderTop:'1px solid #E7E7E3',paddingTop:16}}>
                <div style={{fontSize:12,fontWeight:700,color:'#8a8a82',letterSpacing:'.04em',marginBottom:10}}>AGREGAR USUARIO</div>
                <div style={{display:'flex',flexDirection:'column',gap:10,marginBottom:10}}>
                  <div className="form-cols-2">
                    <div className="form-group">
                      <label className="field-label">Usuario</label>
                      <input className="field-input" value={userMgmt.newUser} onChange={e=>setUserMgmt(p=>({...p,newUser:e.target.value,err:''}))} placeholder="nombre de usuario" />
                    </div>
                    <div className="form-group">
                      <label className="field-label">Contraseña</label>
                      <input className="field-input" type="password" value={userMgmt.newPass} onChange={e=>setUserMgmt(p=>({...p,newPass:e.target.value,err:''}))} placeholder="••••••••" />
                    </div>
                  </div>
                  <div className="form-group">
                    <label className="field-label">Nombre completo</label>
                    <input className="field-input" value={userMgmt.newDisplayName||''} onChange={e=>setUserMgmt(p=>({...p,newDisplayName:e.target.value,err:''}))} placeholder="Ej. Juan Pérez" />
                  </div>
                  {session === 'compras' && (
                  <div className="form-group">
                    <label className="field-label">Rol</label>
                    <div style={{display:'flex',gap:8}}>
                      {[['admin','Administrador'],['receptor','Receptor']].map(([v,label]) => (
                        <button key={v} style={{flex:1,padding:'8px 0',borderRadius:6,border:'1px solid',cursor:'pointer',fontWeight:700,fontSize:13,
                          background:(userMgmt.newRole||'receptor')===v?'#FFD200':'#F5F5F0',
                          borderColor:(userMgmt.newRole||'receptor')===v?'#e6be00':'#E0E0DA',
                          color:(userMgmt.newRole||'receptor')===v?'#121212':'#8a8a82'}}
                          onClick={() => setUserMgmt(p=>({...p,newRole:v}))}>
                          {label}
                        </button>
                      ))}
                    </div>
                  </div>
                  )}
                </div>
                {userMgmt.err && <div style={{fontSize:12.5,color:'#C2473D',fontWeight:600,marginBottom:8}}>{userMgmt.err}</div>}
                <button className="btn btn-dark" onClick={addUser}>+ Agregar usuario</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {modal === 'cambiar-pass' && (
        <div className="modal-overlay" onClick={closeModal}>
          <div className="modal-box" onClick={e=>e.stopPropagation()} style={{maxWidth:400}}>
            <div className="modal-header">
              <div className="modal-title">Cambiar contraseña</div>
              <button className="modal-close" onClick={closeModal}>✕</button>
            </div>
            <div className="modal-body" style={{display:'flex',flexDirection:'column',gap:14}}>
              <div className="form-group">
                <label className="field-label">CONTRASEÑA ACTUAL</label>
                <input className="field-input" type="password" value={changePassForm.current} onChange={e=>setChangePassForm(p=>({...p,current:e.target.value,err:''}))} placeholder="••••••••" />
              </div>
              <div className="form-group">
                <label className="field-label">NUEVA CONTRASEÑA</label>
                <input className="field-input" type="password" value={changePassForm.newPass} onChange={e=>setChangePassForm(p=>({...p,newPass:e.target.value,err:''}))} placeholder="Mín. 6 caracteres" />
              </div>
              <div className="form-group">
                <label className="field-label">REPETIR NUEVA CONTRASEÑA</label>
                <input className="field-input" type="password" value={changePassForm.newPass2} onChange={e=>setChangePassForm(p=>({...p,newPass2:e.target.value,err:''}))} placeholder="••••••••" />
              </div>
              {changePassForm.err && <div style={{fontSize:12.5,color:'#C2473D',fontWeight:600}}>{changePassForm.err}</div>}
              <button className="btn btn-dark" style={{width:'100%',justifyContent:'center',height:42}} onClick={doChangePass}>Guardar contraseña</button>
            </div>
          </div>
        </div>
      )}

      {/* Toast */}
      {toast && (
        <div className="toast">
          <span className="toast-dot"/>
          {toast}
        </div>
      )}
    </div>
  )
}
