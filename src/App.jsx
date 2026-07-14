import { useState, useEffect, useCallback, useRef } from 'react'
import { supabase } from './supabase.js'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'

const TALLE_ORDER = ['2','4','6','8','10','12','14','Único','S','M','L','XL','XXL','XXXL']
const TALLES_ADULTO = ['S','M','L','XL','XXL','XXXL','Único']
const TALLES_NINO   = ['2','4','6','8','10','12','14']
const RECEPTORES = ['1° División','3° División','Juveniles','Captación','Femenino','Juveniles Femenino','Fútbol Sala Masculino','Fútbol Sala Femenino','Basket','Deportes Anexos','Funcionarios','Protocolo']
const DISCIPLINAS_DEPORTES_ANEXOS = ['Atletismo','Bowling','Esports','Fútbol Inclusivo','Fútbol Playa Masculino','Fútbol Sala Femenino','Handball','Volley','Teqball','Cricket','Footgolf','Ciclismo','Paracaidismo','Maxi Basket','Automovilismo','Motociclismo','Hockey Patín','Esgrima']
const CATEGORIAS = ['Entrenamiento','Juego','Casual']
const OCUPACIONES = ['3° División','Juveniles','Juveniles Femenino','Captacion']
const DIVISIONES            = ['Sub 19','Sub 17','Sub 16','Sub 15','Sub 14']
const DIVISIONES_FEM        = ['Sub 19','Sub 16','Sub 14']
const CARGOS_REG = ['Coordinación','Director Técnico','Ayudante Técnico','Videoanalista','Preparador Físico','Entrenador de Arqueros','Doctor/a','Kinesiólogo/a','Utilero','Administración Palacio']
const CARGOS_SIN_SECTOR = ['Administración Palacio']
const ESTANTES = ['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','50','51','TRANSITO']
const ALTURAS = ['A','B','C','D','E','O']
const CAMISETA_TIPOS = ['Titular','Alternativa','3°']
const SHORT_TIPOS = ['Titular','Alternativa']
const REP_TIPOS_JUGADOR = ['TRADICIONAL','AMARILLA','VERDE']
const REP_TIPOS_GOLERO  = ['NEGRO','NARANJA','CREMA']
const getRepTipos = (posicion) => posicion === 'Golero' ? REP_TIPOS_GOLERO : REP_TIPOS_JUGADOR

const DEFAULT_USERS = [
  { username:'compras', password:'peniarol1891', role:'admin', displayName:'Compras Peñarol', status:'aprobado' },
  { username:'iabella@capenarol.com.uy', password:'Temporal2026', role:'receptor', displayName:'Ignacio Abella Goday', status:'aprobado' },
  { username:'jfalero@capenarol.com.uy', password:'Temporal2026', role:'receptor', displayName:'Joaquín Falero', status:'aprobado' },
  { username:'rferrari@capenarol.com.uy', password:'Temporal2026', role:'receptor', displayName:'Rodrigo Ferrari', status:'aprobado' },
  { username:'clauria@capenarol.com.uy', password:'Temporal2026', role:'receptor', displayName:'Camilo Lauria', status:'aprobado' },
]
const EMPTY_DB = { articles:[], deliveries:[], movimientos:[], nextId:1, nextDel:1, nextMov:1, nextRep:1, users: DEFAULT_USERS, camisetasUtileria:[], reposiciones:[], plantel:[] }
const COMPETICIONES = ['CAMPEONATO URUGUAYO','CONMEBOL','COPA LIBERTADORES FEMENINA','COPA LIBERTADORES FÚTBOL SALA','COPA INTERCONTINENTAL SUB 20']
const MODELOS_JUGADOR = ['TRADICIONAL','GRIS','AMARILLA','DORADA','NEGRA Y DORADA','NEGRA Y AMARILLA','AMARILLA FLÚO']
const MODELOS_GOLERO  = ['VERDE','NARANJA','NEGRO','GRIS','ROSADO','CREMA','AMARILLO FLÚO','AMARILLO']

const USERS_KEY = 'dep_usuarios_v1'
const SESSION_KEY = 'dep_session'


async function loadFromSupabase() {
  const [{ data, error }, { data: usersRow }, { data: utiRow }] = await Promise.all([
    supabase.from('deposito_state').select('*').eq('id', 1).single(),
    supabase.from('deposito_state').select('deliveries').eq('id', 2).single(),
    supabase.from('deposito_state').select('articles,deliveries,movimientos,next_del').eq('id', 3).single(),
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
  // Eliminar usuario legacy con email como username
  users = users.filter(u => u.username !== 'compras@capenarol.com.uy')
  // Siempre garantizar que los usuarios de DEFAULT_USERS estén presentes
  DEFAULT_USERS.forEach(du => {
    if (!users.find(u => u.username === du.username)) users.push(du)
  })
  const rawDeliveries = data.deliveries || []
  const needsMigration = rawDeliveries.some(d => !d.creadoPor)
  const deliveries = needsMigration
    ? rawDeliveries.map(d => d.creadoPor ? d : { ...d, creadoPor: 'Emiliano Domínguez' })
    : rawDeliveries
  if (needsMigration) {
    supabase.from('deposito_state').upsert({ id: 1, articles: data.articles, deliveries, movimientos: data.movimientos, next_id: data.next_id, next_del: data.next_del, next_mov: data.next_mov })
      .then(({ error }) => { if (error) console.error('Error migrando creadoPor:', error.message) })
  }
  return {
    articles: (data.articles || []).map(a => ({
      ...a, sizes: (a.sizes || []).map(s => ({ talle: s.talle, qty: Number(s.qty)||0, min: Number(s.min)||0 }))
    })),
    deliveries,
    movimientos: data.movimientos || [],
    nextId: data.next_id || 1,
    nextDel: data.next_del || 1,
    nextMov: data.next_mov || 1,
    users,
    camisetasUtileria: utiRow?.articles || [],
    reposiciones: utiRow?.deliveries || [],
    nextRep: utiRow?.next_del || 1,
    plantel: utiRow?.movimientos || [],
  }
}

async function saveToSupabase(db) {
  // Row id=2 (users) is managed exclusively by saveUsers() to avoid session-collision overwrites
  const [r1, r3] = await Promise.all([
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
      id: 3,
      articles: db.camisetasUtileria || [],
      deliveries: db.reposiciones || [],
      movimientos: db.plantel || [],
      next_id: 0,
      next_del: db.nextRep || 1,
      next_mov: 0,
      updated_at: new Date().toISOString(),
    }),
  ])
  return !r1.error && !r3.error
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
  const [filterUbic, setFilterUbic] = useState('')
  const [modal, setModal] = useState(null)
  const [confirm, setConfirm] = useState(null)
  const [editing, setEditing] = useState(null)
  const [movFilter, setMovFilter] = useState('Todos')
  const [delFilterReceptor, setDelFilterReceptor] = useState('')
  const [delFilterDisciplina, setDelFilterDisciplina] = useState('')
  const [delFilterPersona, setDelFilterPersona] = useState('')
  const [delFilterPaga, setDelFilterPaga] = useState('')
  const [selectedDeliveryId, setSelectedDeliveryId] = useState(null)
  const [selectedReceptor, setSelectedReceptor] = useState(null)
  const [utiFilter, setUtiFilter] = useState('')
  const [utiFilterTipo, setUtiFilterTipo] = useState('')
  const [utiFilterTemp, setUtiFilterTemp] = useState('')
  const [utiFilterModelo, setUtiFilterModelo] = useState('')
  const [utiForm, setUtiForm] = useState({ tipo:'', competicion:'', numero:'', jugador:'', talle:'S', modelo:'', estampado:'', parches:'', detalle:'', temporada:'', id:null })
  const [utiModal, setUtiModal] = useState(false)
  const [repForm, setRepForm] = useState({ editId:null, concepto:'', descuento:true, rows:[] })
  const [repModal, setRepModal] = useState(false)
  const [repDetail, setRepDetail] = useState(null)
  const [repResumen, setRepResumen] = useState(null)
  const [repFilterTorneo, setRepFilterTorneo] = useState('')
  const [repConceptoEdit, setRepConceptoEdit] = useState(null)
  const [disciplinaEdit, setDisciplinaEdit] = useState(null)
  const [pumaMetric, setPumaMetric] = useState('unidades')
  const [repTab, setRepTab] = useState('reposiciones')
  const [plantelForm, setPlantelForm] = useState({id:null,numero:'',nombre:'',posicion:'Jugador',talleCamiseta:'L',talleShort:'L'})
  const [plantelModal, setPlantelModal] = useState(false)
  const [rechazarModal, setRechazarModal] = useState({ delId: null, motivo: '' })
  const [toast, setToast] = useState('')
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [depositosOpen, setDepositosOpen] = useState(true)
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
  const initialLoadDone = useRef(false)
  const hasPendingSave = useRef(false)
  const dbRef = useRef(db)

  // delivery/devolución form
  const [nd, setNd] = useState({ mode:'entrega', persona:'', receptor:'', disciplina:'', fecha:'', cCode:'', cSearch:'', cUbic:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' })
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

  // Save to Supabase whenever data changes (debounced 800ms).
  // Skip the first fire right after the initial load to avoid overwriting data that
  // another session saved between our load and our first debounce tick.
  useEffect(() => {
    if (loading || !saveEnabled.current) return
    if (!initialLoadDone.current) { initialLoadDone.current = true; return }
    clearTimeout(saveTimer.current)
    hasPendingSave.current = true
    saveTimer.current = setTimeout(async () => {
      hasPendingSave.current = false
      const ok = await saveToSupabase(db)
      if (!ok) showToast('Error al guardar. Verificá la conexión.')
    }, 800)
  }, [db, loading])

  // Flush any pending save immediately when tab is hidden (mobile: switch app / close tab).
  // Only flush if there is actually a pending debounce save — never write stale load data.
  useEffect(() => {
    const flush = () => {
      if (!saveEnabled.current || !hasPendingSave.current) return
      clearTimeout(saveTimer.current)
      hasPendingSave.current = false
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
    const merged = [...list]
    DEFAULT_USERS.forEach(du => { if (!merged.find(u => u.username === du.username)) merged.push(du) })
    setDb(prev => ({ ...prev, users: merged }))
    supabase.from('deposito_state').upsert({ id: 2, deliveries: merged })
      .then(({ error }) => { if (error) console.error('Error guardando usuarios:', error.message) })
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
    if(!CARGOS_SIN_SECTOR.includes(cargo) && !categoria) { setRegForm(p=>({...p,err:'Seleccioná tu sector.'})); return }
    if(!CARGOS_SIN_SECTOR.includes(cargo) && ['Juveniles','Juveniles Femenino'].includes(categoria) && cargo !== 'Coordinación' && !division) { setRegForm(p=>({...p,err:'Seleccioná tu división.'})); return }
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
  const openEntrega = () => { setNd({ mode:'entrega', persona:'', receptor:'', disciplina:'', cCode:'', cSearch:'', cUbic:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' }); setModal('entrega') }
  const openDevolucion = () => { setNd({ mode:'devolucion', persona:'', receptor:'', disciplina:'', cCode:'', cSearch:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' }); setModal('entrega') }
  const openEntregaFromDetail = () => { const a = byCode(selectedCode); setNd({ mode:'entrega', persona:'', receptor:'', disciplina:'', cCode:a?a.code:'', cSearch:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' }); setModal('entrega') }
  const openDevolucionFromDetail = () => { const a = byCode(selectedCode); setNd({ mode:'devolucion', persona:'', receptor:'', disciplina:'', cCode:a?a.code:'', cSearch:'', cTalle:'', cQty:'', paga:null, lines:[], toUser:'' }); setModal('entrega') }

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
    if(nd.receptor === 'Deportes Anexos' && !nd.disciplina.trim()) { showToast('Ingresá la disciplina.'); return }
    if(nd.lines.length === 0) { showToast('Agregá al menos un artículo.'); return }
    let newDbState = null
    setDb(s => {
      const articles = s.articles.map(a => ({...a, sizes: a.sizes.map(z => ({...z}))}))
      const movimientos = [...s.movimientos]
      let mid = s.nextMov
      const fecha = nd.fecha || today()
      nd.lines.forEach(l => {
        // Find the entry at the specific location (ubic), fallback to first with the talle
        const a = l.ubic
          ? articles.find(x => x.code === l.code && x.ubic === l.ubic)
          : articles.find(x => x.code === l.code && x.sizes.some(sz => sz.talle === l.talle))
        const z = a && a.sizes.find(x => x.talle === l.talle)
        if(z) z.qty = esDev ? z.qty + l.qty : Math.max(0, z.qty - l.qty)
        if(esDev) {
          movimientos.unshift({id:mid++, code:l.code, name:a?.name||l.code, tipo:'entrada', fecha, talle:l.talle, qty:l.qty, detalle:'Devolución de '+nd.persona+' ('+nd.receptor+')', creadoPor:currentUser?.displayName||session})
        } else {
          movimientos.unshift({id:mid++, code:l.code, name:a?.name||l.code, tipo:'salida', fecha, talle:l.talle, qty:l.qty, detalle:'Entrega a '+nd.persona+' ('+nd.receptor+(nd.disciplina?' - '+nd.disciplina:'')+')', delId:s.nextDel, creadoPor:currentUser?.displayName||session})
        }
      })
      const activeArticles = articles.filter(a => total(a) > 0)
      if(esDev) { const r = { ...s, articles:activeArticles, movimientos, modal:null, nextMov:mid }; newDbState = r; return r }
      const toUser = nd.toUser || null
      const status = toUser ? 'pendiente' : 'aceptado'
      const confirmedAt = toUser ? null : fecha
      const deliveries = [{id:s.nextDel, fecha, persona:nd.persona.trim(), receptor:nd.receptor, disciplina:nd.receptor==='Deportes Anexos'?nd.disciplina.trim():undefined, paga:nd.receptor==='Protocolo'?nd.paga:null, monto:nd.receptor==='Protocolo'&&nd.paga==='si'?ndMonto:null, lines:[...nd.lines], toUser, status, confirmedAt, creadoPor:currentUser?.displayName||session}, ...s.deliveries]
      const r = { ...s, articles:activeArticles, movimientos, deliveries, nextDel:s.nextDel+1, nextMov:mid }
      newDbState = r; return r
    })
    // Guardar inmediatamente en Supabase sin esperar el debounce de 800ms
    if (newDbState) saveToSupabase(newDbState)
    // Enviar email de notificación si la entrega va a un usuario específico
    if (nd.toUser && nd.mode !== 'devolucion') {
      const recipient = db.users.find(u => u.username === nd.toUser)
      const recipientEmail = recipient?.email || (recipient?.username?.includes('@') ? recipient?.username : null)
      if (recipientEmail) {
        fetch('/api/send-email', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            to: recipientEmail,
            displayName: recipient.displayName || recipient.username,
            lines: nd.lines,
            delId: db.nextDel,
          })
        })
        .then(r => r.json())
        .then(d => { if (d.ok) { showToast('Email de notificación enviado a ' + recipientEmail) } else { showToast('Error al enviar email: ' + (d.error || 'error desconocido')) } })
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
    const ubic = estante === 'TRANSITO' ? 'TRANSITO' : (estante||'1') + (altura||'A')
    const sizes = tallesArr.map(t => ({talle:t, qty:tallesQty[t]||0, min:tallesMins[t]||0}))
    const precioNum = parseFloat(precio)||0
    const fecha = today()
    setDb(s => {
      let nextMov = s.nextMov
      const newMovs = sizes
        .filter(sz => sz.qty > 0)
        .map(sz => ({id:nextMov++, code, name, tipo:'entrada', fecha, talle:sz.talle, qty:sz.qty, detalle:'Stock inicial', creadoPor:currentUser?.displayName||session}))
      return {
        ...s,
        articles: [{id:s.nextId, code, name, cat:ncat, ubic, precio:precioNum, sizes}, ...s.articles],
        nextId: s.nextId+1,
        movimientos: [...newMovs, ...s.movimientos],
        nextMov
      }
    })
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
      const newMovs = entries.map(e => ({id:nextMov++, code, name:artName, tipo:'entrada', fecha, talle:e.talle, qty:e.q, detalle:'Ingreso de stock', creadoPor:currentUser?.displayName||session}))
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
      const movimientos = [{id:s.nextMov, code, name:artName, tipo:(delta>0?'entrada':'salida'), fecha, talle:aj.talle, qty:Math.abs(delta), detalle:'Ajuste por recuento (de '+cur+' a '+q+')', creadoPor:currentUser?.displayName||session}, ...s.movimientos]
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
    const newUbic = mv.estante === 'TRANSITO' ? 'TRANSITO' : mv.estante + mv.altura
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
    showToast('Movido a ' + (mv.estante === 'TRANSITO' ? 'TRANSITO' : mv.estante + mv.altura) + '.')
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
  const receptorRechazar = (delId, motivo) => {
    setDb(s => {
      const del = s.deliveries.find(d => d.id === delId); if(!del) return {...s}
      // revert stock
      const articles = s.articles.map(a => ({...a, sizes:a.sizes.map(z=>({...z}))}))
      del.lines.forEach(l => { const a=articles.find(x=>x.code===l.code); const z=a&&a.sizes.find(x=>x.talle===l.talle); if(z) z.qty+=l.qty })
      const deliveries = s.deliveries.map(d => d.id === delId ? {...d, status:'rechazado', confirmedAt:today(), motivoRechazo: motivo||''} : d)
      return {...s, articles, deliveries}
    })
    setRechazarModal({ delId: null, motivo: '' })
    showToast('Entrega rechazada y stock restituido.')
  }

  // ---- Derived data ----
  const { articles, deliveries, movimientos } = db
  const codeName = articles.reduce((acc, a) => { acc[a.code] = a.name; return acc }, {})

  // Current user role
  const allUsers = db.users
  const currentUser = allUsers.find(u => u.username === session) || null
  const isReceptor  = currentUser?.role === 'receptor'
  const isSoloVista = currentUser?.role === 'solo-vista'

  // Receptor users list (for the delivery modal selector)
  const receptorUsers = allUsers.filter(u => u.status === 'aprobado')

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

  const normStr = s => s.normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase()
  const q = normStr(search.trim())
  let filtered = articles.filter(a => cat==='Todas' || a.cat===cat)
  if(q) filtered = filtered.filter(a => normStr(a.name).includes(q) || normStr(a.code).includes(q) || normStr(a.ubic||'').includes(q))
  if(filterUbic) filtered = filtered.filter(a => (a.ubic||'') === filterUbic)

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

  const availableUbics = [...new Set(articles.map(a => a.ubic).filter(Boolean))].sort((a,b) => {
    const pa = parseUbic(a), pb = parseUbic(b)
    if(pa.n !== pb.n) return pa.n - pb.n
    return pa.l.localeCompare(pb.l)
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
  const REP_FILTER = 'Reposiciones 1° División'
  const deportesAnexosDisciplinas = [...new Set(deliveries.filter(d => d.receptor==='Deportes Anexos' && d.disciplina).map(d => d.disciplina))].sort()
  const filteredDeliveryRows = deliveryRows
    .filter(d => !delFilterReceptor || delFilterReceptor === REP_FILTER || d.receptor === delFilterReceptor)
    .filter(d => delFilterReceptor !== 'Deportes Anexos' || !delFilterDisciplina || d.disciplina === delFilterDisciplina)
    .filter(d => !delFilterPersona || d.persona.toLowerCase().includes(delFilterPersona.toLowerCase()))
    .filter(d => delFilterReceptor !== 'Protocolo' || !delFilterPaga || d.paga === delFilterPaga)
  const repRows = (db.reposiciones||[]).map(r => {
    const totalCamisetas = (r.jugadores||[]).reduce((s,j)=>s+(Number(j.cantCamiseta)||0),0)
    const totalShorts = (r.jugadores||[]).reduce((s,j)=>s+(Number(j.cantShort)||0),0)
    const parts = []
    if (totalCamisetas) parts.push(`${totalCamisetas} camiseta${totalCamisetas!==1?'s':''}`)
    if (totalShorts) parts.push(`${totalShorts} short${totalShorts!==1?'s':''}`)
    return { id:r.id, fecha:r.fecha, persona:r.concepto, subLabel:r.torneo||'', ini:ini(r.concepto), resumen:parts.join(' · ')||'—', totalUd:totalCamisetas+totalShorts, creadoPor:r.creadoPor, _rep:r }
  })
  const filteredRepRows = repRows.filter(r => !delFilterPersona || r.persona.toLowerCase().includes(delFilterPersona.toLowerCase()))
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

  const TORNEOS_CON_FECHA = ['APERTURA','CLAUSURA','INTERMEDIO']
  const openRepModal = () => {
    setRepForm({
      editId:null,
      concepto:'',
      descuento:true,
      torneo:'APERTURA',
      fechaTorneo:'1',
      tipoCamisetaJugador: REP_TIPOS_JUGADOR[0],
      tipoCamisetaGolero: REP_TIPOS_GOLERO[0],
      rows:(db.plantel||[]).sort((a,b)=>(Number(a.numero)||0)-(Number(b.numero)||0)).map(j=>({...j,cantCamiseta:'',cantShort:'',descuento:true}))
    })
    setRepModal(true)
  }
  const openRepEdit = (rep) => {
    const plantelRows = (db.plantel||[]).sort((a,b)=>(Number(a.numero)||0)-(Number(b.numero)||0)).map(j => {
      const ex = (rep.jugadores||[]).find(jj => jj.nombre===j.nombre)
      return { ...j, cantCamiseta: ex ? String(ex.cantCamiseta) : '', cantShort: ex ? String(ex.cantShort) : '', descuento: ex ? ex.descuento !== false : true }
    })
    setRepForm({
      editId: rep.id,
      concepto: rep.concepto,
      descuento: rep.descuento !== false,
      torneo: rep.torneo || 'APERTURA',
      fechaTorneo: rep.fechaTorneo != null ? String(rep.fechaTorneo) : '1',
      tipoCamisetaJugador: rep.tipoCamisetaJugador || REP_TIPOS_JUGADOR[0],
      tipoCamisetaGolero: rep.tipoCamisetaGolero || REP_TIPOS_GOLERO[0],
      rows: plantelRows
    })
    setRepDetail(null)
    setRepModal(true)
  }
  const saveReposicion = () => {
    if (!repForm.concepto.trim()) { showToast('Ingresá el concepto.'); return }
    const jugadores = repForm.rows
      .filter(r => Number(r.cantCamiseta)>0 || Number(r.cantShort)>0)
      .map(r => {
        const tipo = (r.posicion||'Jugador')==='Golero' ? repForm.tipoCamisetaGolero : repForm.tipoCamisetaJugador
        return { numero:r.numero, nombre:r.nombre, posicion:r.posicion||'Jugador',
          talleCamiseta:r.talleCamiseta, talleShort:r.talleShort, descuento:r.descuento !== false,
          tipoCamiseta: tipo, cantCamiseta:Number(r.cantCamiseta)||0, cantShort:Number(r.cantShort)||0 }
      })
    if (!jugadores.length) { showToast('Ingresá al menos una cantidad.'); return }
    const tieneFecha = TORNEOS_CON_FECHA.includes(repForm.torneo)
    if (repForm.editId) {
      setDb(s => ({...s, reposiciones:(s.reposiciones||[]).map(r => r.id===repForm.editId
        ? {...r, concepto:repForm.concepto.trim(), torneo:repForm.torneo, descuento:repForm.descuento,
            fechaTorneo: tieneFecha ? Number(repForm.fechaTorneo) : null,
            tipoCamisetaJugador:repForm.tipoCamisetaJugador, tipoCamisetaGolero:repForm.tipoCamisetaGolero, jugadores}
        : r)}))
      showToast('Reposición actualizada.')
    } else {
      setDb(s => {
        const rep = { id:s.nextRep, fecha:today(), concepto:repForm.concepto.trim(), creadoPor:currentUser?.displayName||session,
          torneo:repForm.torneo, fechaTorneo: tieneFecha ? Number(repForm.fechaTorneo) : null, descuento:repForm.descuento,
          tipoCamisetaJugador:repForm.tipoCamisetaJugador, tipoCamisetaGolero:repForm.tipoCamisetaGolero, jugadores }
        return { ...s, reposiciones:[rep,...(s.reposiciones||[])], nextRep:s.nextRep+1 }
      })
      showToast('Reposición registrada.')
    }
    setRepModal(false)
  }
  const deleteReposicion = (id) => {
    setDb(s => ({...s, reposiciones:(s.reposiciones||[]).filter(r=>r.id!==id)}))
    setRepDetail(null)
    showToast('Reposición eliminada.')
  }
  const exportRepToExcel = async (rep) => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet(rep.concepto.slice(0,31))
    ws.columns = [
      {width:8.5},{width:17.625},{width:9.25},{width:10.25},{width:6.5},{width:9.75}
    ]
    const YELLOW    = {type:'pattern',pattern:'solid',fgColor:{argb:'FFFFD966'}}
    const FILL_WHT  = {type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFFFF'}}
    const FILL_GRAY = {type:'pattern',pattern:'solid',fgColor:{argb:'FFD0D0D0'}}
    const FILL_ORAN = {type:'pattern',pattern:'solid',fgColor:{argb:'FFFFE5CC'}}
    const FILL_CREM = {type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF8E8'}}
    const F_BOLD    = {name:'Arial',size:12,bold:true}
    const F_NORM    = {name:'Arial',size:12}
    const CENTER    = {horizontal:'center',vertical:'middle'}
    const BORDER    = {left:{style:'thin'},right:{style:'thin'},top:{style:'thin'},bottom:{style:'thin'}}
    const style = (cell,fill,font) => { cell.fill=fill; cell.font=font; cell.alignment=CENTER; cell.border=BORDER }
    const goleroFill = (j) => {
      if (j.posicion !== 'Golero') return FILL_WHT
      if (j.tipoCamiseta === 'NEGRO')   return FILL_GRAY
      if (j.tipoCamiseta === 'NARANJA') return FILL_ORAN
      if (j.tipoCamiseta === 'CREMA')   return FILL_CREM
      return FILL_WHT
    }

    // Fila 1: concepto - torneo fecha
    ws.mergeCells('A1:F1')
    const titulo = [rep.concepto.toUpperCase(), rep.torneo?(rep.torneo+(rep.fechaTorneo?' FECHA '+rep.fechaTorneo:'')):null].filter(Boolean).join(' - ')
    style(ws.getCell('A1'), YELLOW, F_BOLD); ws.getCell('A1').value = titulo; ws.getRow(1).height = 20

    // Fila 2: equipo tipo jugador / golero
    ws.mergeCells('A2:F2')
    const equipo = ['EQUIPO',rep.tipoCamisetaJugador||'','GOLERO',rep.tipoCamisetaGolero||''].join(' ').replace(/\s+/g,' ').trim()
    style(ws.getCell('A2'), YELLOW, F_BOLD); ws.getCell('A2').value = equipo; ws.getRow(2).height = 20

    // Fila 3: headers
    ;['NÚMERO','NOMBRE','CAMISETA','CANTIDAD','SHORT','CANTIDAD'].forEach((h,i) => {
      const c = ws.getRow(3).getCell(i+1); style(c, YELLOW, F_BOLD); c.value = h
    })
    ws.getRow(3).height = 20

    // Filas de jugadores
    ;(rep.jugadores||[]).forEach((j,idx) => {
      const r = ws.getRow(idx+4)
      r.height = 18
      const fill = goleroFill(j)
      ;[j.numero||'—', j.nombre||'—', j.talleCamiseta||'—', Number(j.cantCamiseta)||0, j.talleShort||'—', Number(j.cantShort)||0].forEach((v,i) => {
        const c = r.getCell(i+1); style(c, fill, F_NORM); c.value = v
      })
    })

    // Fila TOTAL
    const totN = (rep.jugadores||[]).length + 4
    const totCam = (rep.jugadores||[]).reduce((s,j)=>s+(Number(j.cantCamiseta)||0),0)
    const totSht = (rep.jugadores||[]).reduce((s,j)=>s+(Number(j.cantShort)||0),0)
    ws.mergeCells(`A${totN}:B${totN}`)
    ws.getRow(totN).height = 20
    ;[[1,'TOTAL'],[3,''],[4,totCam],[5,''],[6,totSht]].forEach(([col,val]) => {
      const c = ws.getRow(totN).getCell(col); style(c, YELLOW, F_BOLD); c.value = val
    })

    // Descargar
    const buf = await wb.xlsx.writeBuffer()
    const blob = new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a'); a.href=url; a.download=`${rep.concepto.replace(/[\\/:*?"<>|]/g,'-')}.xlsx`; a.click()
    URL.revokeObjectURL(url)
  }
  const saveDisciplinaEdit = (deliveryId) => {
    if (!disciplinaEdit?.trim()) { showToast('Ingresá la disciplina.'); return }
    setDb(s => ({...s, deliveries:s.deliveries.map(d=>d.id===deliveryId?{...d,disciplina:disciplinaEdit.trim()}:d)}))
    setDisciplinaEdit(null)
    showToast('Disciplina actualizada.')
  }
  const saveRepConcepto = () => {
    if (!repConceptoEdit?.trim()) { showToast('El concepto no puede estar vacío.'); return }
    setDb(s => ({...s, reposiciones:(s.reposiciones||[]).map(r=>r.id===repDetail.id?{...r,concepto:repConceptoEdit.trim()}:r)}))
    setRepDetail(p => ({...p, concepto:repConceptoEdit.trim()}))
    setRepConceptoEdit(null)
    showToast('Concepto actualizado.')
  }
  const savePlantelJugador = () => {
    if (!plantelForm.nombre.trim()) { showToast('Ingresá el nombre del jugador.'); return }
    setDb(s => {
      const list = s.plantel || []
      if (plantelForm.id !== null) {
        return {...s, plantel:list.map(j=>j.id===plantelForm.id?{...plantelForm}:j)}
      }
      const newId = list.length > 0 ? Math.max(...list.map(j=>j.id))+1 : 1
      return {...s, plantel:[...list, {...plantelForm, id:newId}]}
    })
    setPlantelModal(false)
    showToast(plantelForm.id!==null ? 'Jugador actualizado.' : 'Jugador agregado al plantel.')
  }
  const deletePlantelJugador = (id) => {
    setDb(s => ({...s, plantel:(s.plantel||[]).filter(j=>j.id!==id)}))
    showToast('Jugador eliminado.')
  }
  const utiFiltered = (db.camisetasUtileria || []).filter(c =>
    (!utiFilter      || c.competicion === utiFilter) &&
    (!utiFilterTipo  || c.tipo === utiFilterTipo) &&
    (!utiFilterTemp  || c.temporada === utiFilterTemp) &&
    (!utiFilterModelo|| c.modelo === utiFilterModelo)
  )

  const receptorCards = RECEPTORES.map(name => {
    const ds = deliveries.filter(d => d.receptor===name)
    const unidades = ds.reduce((s,d) => s+d.lines.reduce((x,l)=>x+l.qty,0),0)
    const monto = ds.reduce((s,d) => s+d.lines.reduce((x,l)=>{
      const art = articles.find(a=>a.code===l.code)
      return x+(art?.precio||0)*l.qty
    },0),0)
    return { name, ini:ini(name), count:ds.length, unidades, monto }
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
  const ndOk = nd.persona && nd.persona.trim() && nd.receptor && nd.lines.length > 0 && (nd.receptor !== 'Deportes Anexos' || nd.disciplina.trim())

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
            {!CARGOS_SIN_SECTOR.includes(regForm.cargo) && (
            <div className="form-group">
              <label className="field-label" style={{color:'#8a8a82'}}>SECTOR</label>
              <select className="field-input" value={regForm.categoria} onChange={e=>setRegForm(p=>({...p,categoria:e.target.value,division:'',err:''}))}>
                <option value="">Seleccioná tu sector…</option>
                {OCUPACIONES.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>
            )}
            {!CARGOS_SIN_SECTOR.includes(regForm.cargo) && ['Juveniles','Juveniles Femenino'].includes(regForm.categoria) && regForm.cargo !== 'Coordinación' && (
              <div className="form-group">
                <label className="field-label" style={{color:'#8a8a82'}}>DIVISIÓN</label>
                <select className="field-input" value={regForm.division} onChange={e=>setRegForm(p=>({...p,division:e.target.value,err:''}))}>
                  <option value="">Seleccioná tu división…</option>
                  {(regForm.categoria === 'Juveniles Femenino' ? DIVISIONES_FEM : DIVISIONES).map(d => <option key={d} value={d}>{d}</option>)}
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
            <div style={{fontFamily:'Archivo Black,sans-serif',fontSize:14,color:'#FFD200',letterSpacing:'.05em'}}>{currentUser?.role==='receptor' ? 'INDUMENTARIA CLUB ATLÉTICO PEÑAROL' : 'DEPÓSITO · INDUMENTARIA'}</div>
            <div style={{fontSize:13,color:'#fff',marginTop:2}}>Hola, <b>{currentUser?.displayName || session}</b></div>
          </div>
          <button onClick={() => window.location.reload()} style={{background:'#2a2a2a',border:'1px solid #3a3a3a',color:'#ccc',borderRadius:8,padding:'8px 14px',cursor:'pointer',fontSize:13,marginRight:8}}>↺</button>
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
                  <button onClick={() => setRechazarModal({ delId: d.id, motivo: '' })} style={{flex:1,background:'#C2473D',color:'#fff',border:'none',borderRadius:8,padding:'10px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>✕ Rechazar</button>
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
                  {st === 'rechazado' && d.motivoRechazo && (
                    <div style={{padding:'8px 18px',background:'#FBEAE8',borderBottom:'1px solid #f5c6c3',fontSize:12.5,color:'#C2473D'}}>
                      <b>Motivo:</b> {d.motivoRechazo}
                    </div>
                  )}
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

        {/* Rechazar modal */}
        {rechazarModal.delId !== null && (
          <div className="modal-overlay" onClick={() => setRechazarModal({ delId: null, motivo: '' })}>
            <div className="modal-box" onClick={e => e.stopPropagation()} style={{maxWidth:420}}>
              <div className="modal-header">
                <div className="modal-title">Motivo de rechazo</div>
                <button className="modal-close" onClick={() => setRechazarModal({ delId: null, motivo: '' })}>×</button>
              </div>
              <div className="modal-body" style={{display:'flex',flexDirection:'column',gap:14}}>
                <p style={{margin:0,fontSize:13.5,color:'#6a6a62'}}>Explicá brevemente por qué rechazás esta entrega.</p>
                <div className="form-group">
                  <label className="field-label">Motivo</label>
                  <textarea
                    className="field-input"
                    rows={3}
                    style={{resize:'vertical',fontFamily:'inherit'}}
                    placeholder="Ej: Talle incorrecto, artículo dañado…"
                    value={rechazarModal.motivo}
                    onChange={e => setRechazarModal(p => ({...p, motivo: e.target.value}))}
                  />
                </div>
                {!rechazarModal.motivo.trim() && (
                  <div style={{fontSize:12,color:'#8a8a82'}}>El motivo es obligatorio para rechazar.</div>
                )}
                <div style={{display:'flex',gap:10}}>
                  <button className="btn btn-ghost" style={{flex:1}} onClick={() => setRechazarModal({ delId: null, motivo: '' })}>Cancelar</button>
                  <button
                    style={{flex:1,padding:'10px 0',borderRadius:8,border:'none',cursor:rechazarModal.motivo.trim()?'pointer':'not-allowed',fontWeight:700,fontSize:14,background:rechazarModal.motivo.trim()?'#C2473D':'#e0a09a',color:'#fff'}}
                    onClick={() => { if(rechazarModal.motivo.trim()) receptorRechazar(rechazarModal.delId, rechazarModal.motivo.trim()) }}
                  >✕ Rechazar entrega</button>
                </div>
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

  return (
    <div className="app-shell">
      {/* Mobile overlay */}
      <div className={`mobile-overlay${sidebarOpen?' open':''}`} onClick={() => setSidebarOpen(false)} />

      {/* Sidebar */}
      <aside className={`sidebar${sidebarOpen?' open':''}`}>
        <div className="sidebar-logo">
          <img src="/escudo.png" alt="Peñarol" />
          <div className="sidebar-logo-text">
            <div className="name">CLUB ATLÉTICO PEÑAROL</div>
            <div className="sub">Indumentaria</div>
          </div>
        </div>
        <nav className="sidebar-nav">
          {/* PANEL PRINCIPAL top-level */}
          {(() => { const isActive = view==='panel'; return (
            <button className={`nav-item${isActive?' active':''}`} onClick={() => goView('panel')}>
              <span className="nav-dot" />
              PANEL PRINCIPAL
              {isActive && <img src="/escudo.png" alt="" style={{height:20,width:'auto',marginLeft:'auto',opacity:0.85}} />}
            </button>
          )})()}
          {/* Grupo DEPÓSITO */}
          {(() => { const isGrpActive = ['inventario','detalle','entregas','movimientos','receptores','utileria'].includes(view); return (
            <button className={`nav-item${isGrpActive?' active':''}`} onClick={() => setDepositosOpen(o => !o)}>
              <span className="nav-dot" />
              DEPÓSITO
              <span style={{marginLeft:'auto',display:'flex',alignItems:'center',gap:6}}>
                <span style={{fontSize:10,opacity:0.6}}>{depositosOpen ? '▾' : '▸'}</span>
                {isGrpActive && <img src="/escudo.png" alt="" style={{height:20,width:'auto',opacity:0.85}} />}
              </span>
            </button>
          )})()}
          {depositosOpen && [['inventario','INVENTARIO'],['entregas','ENTREGAS'],['movimientos','MOVIMIENTOS'],['receptores','RECEPTORES'],['utileria','CAMISETAS UTILERÍA']].map(([key,label]) => {
            const isActive = view===key||(key==='inventario'&&view==='detalle')
            return (
              <button key={key} className={`nav-item nav-item-sub${isActive?' active':''}`} onClick={() => goView(key)}>
                <span className="nav-dot" />
                {label}
                {isActive && <img src="/escudo.png" alt="" style={{height:20,width:'auto',marginLeft:'auto',opacity:0.85}} />}
              </button>
            )
          })}
          {/* Items top-level */}
          {[['reposiciones','REPOSICIÓN CAMISETAS'],['contrato-puma','CONTRATO PUMA'],['usuarios-reg','USUARIOS REGISTRADOS']].map(([key,label]) => {
            const isActive = view===key
            return (
              <button key={key} className={`nav-item${isActive?' active':''}`} onClick={() => goView(key)}>
                <span className="nav-dot" />
                {label}
                {isActive && <img src="/escudo.png" alt="" style={{height:20,width:'auto',marginLeft:'auto',opacity:0.85}} />}
              </button>
            )
          })}
        </nav>
        <div className="sidebar-user" style={{flexDirection:'column',gap:10}}>
          <div style={{display:'flex',alignItems:'center',gap:11,minWidth:0}} title={`${currentUser?.displayName || session}\n${session}`}>
            <div className="user-avatar" style={{flexShrink:0}}>{ini(currentUser?.displayName || session || '')}</div>
            <div style={{flex:1,minWidth:0}}>
              <div className="user-name" style={{overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{currentUser?.displayName || session}</div>
              <div className="user-role" style={{overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{session}</div>
            </div>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',gap:2}}>
            <button title="Sincronizar con servidor" onClick={() => saveToSupabase(db).then(ok => showToast(ok ? 'Datos sincronizados.' : 'Error al sincronizar.'))} style={{background:'none',border:'none',color:'#8a8a82',cursor:'pointer',fontSize:18,padding:'0 6px'}}>↺</button>
            <button title="Cambiar contraseña" onClick={()=>{setChangePassForm({current:'',newPass:'',newPass2:'',err:''});setModal('cambiar-pass')}} style={{background:'none',border:'none',color:'#8a8a82',cursor:'pointer',fontSize:16,padding:'0 6px'}}>🔑</button>
            <button title="Gestionar usuarios" onClick={openUserMgmt} style={{background:'none',border:'none',color:'#8a8a82',cursor:'pointer',fontSize:18,padding:'0 6px'}}>⚙</button>
            <button title="Cerrar sesión" onClick={doLogout} style={{background:'none',border:'none',color:'#8a8a82',cursor:'pointer',fontSize:18,padding:'0 6px'}}>⏻</button>
          </div>
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
            {{panel:'PANEL PRINCIPAL',inventario:'INVENTARIO',detalle:'DETALLE',entregas:'ENTREGAS',movimientos:'MOVIMIENTOS',receptores:'RECEPTORES','usuarios-reg':'USUARIOS REGISTRADOS',utileria:'CAMISETAS UTILERÍA',reposiciones:'REPOSICIÓN CAMISETAS','contrato-puma':'CONTRATO PUMA'}[view]}
          </div>
          <div className="topbar-spacer" />
          <div className="search-box">
            <span className="search-icon" />
            <input value={search} onChange={e => { setSearch(e.target.value); if((view==='panel'||view==='detalle')&&e.target.value) setView('inventario') }} placeholder="Buscar artículo…" />
          </div>
          {!isSoloVista && <button className="btn btn-ghost" onClick={openArticulo}>+<span className="btn-label"> Artículo</span></button>}
          {!isSoloVista && <button className="btn btn-ghost" onClick={openDevolucion}>↩<span className="btn-label"> Dev.</span></button>}
          {!isSoloVista && <button className="btn btn-yellow" onClick={openEntrega}>+<span className="btn-label"> Entrega</span></button>}
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
                {!isReceptor && (() => {
                  const myPending = db.deliveries.filter(d => d.toUser === session && (d.status||'aceptado') === 'pendiente')
                  if (!myPending.length) return null
                  return (
                    <div className="card" style={{marginTop:16}}>
                      <div className="card-header">
                        <div className="card-title">Entregas pendientes de confirmación</div>
                        <div className="card-spacer"/>
                        <span className="badge" style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200'}}>{myPending.length}</span>
                      </div>
                      {myPending.map(d => (
                        <div key={d.id} className="table-row" style={{gridTemplateColumns:'1fr auto'}}>
                          <div>
                            <div style={{fontWeight:600,fontSize:13.5}}>Entrega #{d.id} — {d.receptor}</div>
                            <div style={{fontSize:11.5,color:'#8a8a82'}}>{d.fecha} · {d.lines.length} artículo{d.lines.length!==1?'s':''}</div>
                          </div>
                          <div style={{display:'flex',gap:6}}>
                            <button onClick={()=>receptorAceptar(d.id)} style={{padding:'4px 10px',borderRadius:5,border:'none',cursor:'pointer',fontWeight:700,fontSize:11.5,background:'#2e9b5e',color:'#fff'}}>✓ Aceptar</button>
                            <button onClick={()=>setRechazarModal({delId:d.id,motivo:''})} style={{padding:'4px 10px',borderRadius:5,border:'1px solid #C2473D',cursor:'pointer',fontWeight:700,fontSize:11.5,background:'none',color:'#C2473D'}}>✕ Rechazar</button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )
                })()}
                {!isReceptor && !isSoloVista && pendingApprovals.length > 0 && (
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
                <select className="field-input" style={{height:32,fontSize:12.5,padding:'0 8px',minWidth:110,maxWidth:140}}
                  value={filterUbic} onChange={e => setFilterUbic(e.target.value)}>
                  <option value="">Ubicación…</option>
                  {availableUbics.map(u => <option key={u} value={u}>{u}</option>)}
                </select>
                {filterUbic && <button className="chip active" onClick={()=>setFilterUbic('')}>× {filterUbic}</button>}
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
                            {!isSoloVista && <button className="btn btn-yellow" onClick={() => { setSelectedId(entry.id); openReponer() }}>＋ Registrar entrada</button>}
                            {!isSoloVista && <button className="btn btn-dark" onClick={() => { setSelectedId(entry.id); openAjuste() }}>Ajustar stock</button>}
                            {!isSoloVista && <button className="btn btn-ghost" onClick={() => { setSelectedId(entry.id); openMover() }}>⇄ Cambiar de ubicación</button>}
                          </div>
                        </div>
                      )
                    })}

                    {/* Acciones globales */}
                    <div className="detail-actions" style={{marginTop:24,paddingTop:20,borderTop:'1px solid #E7E7E3'}}>
                      {!isSoloVista && <button className="btn btn-ghost" onClick={openEntregaFromDetail}>Registrar entrega</button>}
                      {!isSoloVista && <button className="btn btn-ghost" onClick={openDevolucionFromDetail}>↩ Devolución</button>}
                      {!isSoloVista && <button className="btn btn-ghost btn-full" onClick={() => { setSelectedId(detail.entries[0].id); openEdit() }}>✎ Editar artículo</button>}
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
                      {!isSoloVista && <button className="btn-del" onClick={() => m.delId ? askDeleteDelivery(m.delId) : askDeleteMov(m.id)}>✕</button>}
                    </div>
                  ))}
                </div>
              </div>
            </>
          )}

          {/* ENTREGAS */}
          {view === 'entregas' && (
            <>
              <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center',marginBottom:delFilterReceptor==='Deportes Anexos'?6:12}}>
                <div style={{display:'flex',gap:6,flexWrap:'wrap',flex:1}}>
                  <button className={`chip${delFilterReceptor===''?' active':''}`} onClick={() => { setDelFilterReceptor(''); setDelFilterDisciplina(''); setDelFilterPaga('') }}>Todos</button>
                  {deliveryReceptores.map(r => (
                    <button key={r} className={`chip${delFilterReceptor===r?' active':''}`} onClick={() => { setDelFilterReceptor(r); setDelFilterDisciplina(''); setDelFilterPaga('') }}>{r}</button>
                  ))}
                  <button className={`chip${delFilterReceptor===REP_FILTER?' active':''}`} onClick={() => { setDelFilterReceptor(REP_FILTER); setDelFilterDisciplina(''); setDelFilterPaga('') }}>Reposiciones 1° División</button>
                </div>
                <input className="field-input" style={{width:200,flexShrink:0}} placeholder="Buscar integrante…" value={delFilterPersona} onChange={e => setDelFilterPersona(e.target.value)} />
              </div>
              {delFilterReceptor==='Deportes Anexos' && deportesAnexosDisciplinas.length > 0 && (
                <div style={{display:'flex',gap:6,flexWrap:'wrap',marginBottom:12}}>
                  <button className={`chip${delFilterDisciplina===''?' active':''}`} onClick={() => setDelFilterDisciplina('')}>Todas</button>
                  {deportesAnexosDisciplinas.map(d => (
                    <button key={d} className={`chip${delFilterDisciplina===d?' active':''}`} onClick={() => setDelFilterDisciplina(d)}>{d}</button>
                  ))}
                </div>
              )}
              {delFilterReceptor==='Protocolo' && (
                <div style={{display:'flex',gap:6,flexWrap:'wrap',alignItems:'center',marginBottom:12}}>
                  <span style={{fontSize:12,color:'#8a8a82',fontWeight:600}}>Paga:</span>
                  <button className={`chip${delFilterPaga===''?' active':''}`} onClick={() => setDelFilterPaga('')}>Todos</button>
                  <button className={`chip${delFilterPaga==='si'?' active':''}`} onClick={() => setDelFilterPaga('si')}>SÍ</button>
                  <button className={`chip${delFilterPaga==='no'?' active':''}`} onClick={() => setDelFilterPaga('no')}>NO</button>
                </div>
              )}
            <div className="card table-wrap">
              <div className="card-header">
                <div className="card-title">{delFilterReceptor===REP_FILTER ? 'Reposiciones 1° División' : 'Historial de entregas'}</div>
                <div className="card-spacer"/>
                <span style={{fontSize:12.5,color:'#8a8a82'}}>
                  {delFilterReceptor===REP_FILTER
                    ? `${filteredRepRows.length} de ${(db.reposiciones||[]).length} reposiciones`
                    : `${filteredDeliveryRows.length} de ${kpis.entregas} entregas`}
                </span>
              </div>
              <div className={`table-header ${delFilterReceptor==='Deportes Anexos'?'del-cols-disc':'del-cols'}`}>
                <div>FECHA</div><div>INTEGRANTE / GRUPO</div>
                {delFilterReceptor==='Deportes Anexos' && <div>DISCIPLINA</div>}
                <div className="del-col-detail">DETALLE</div>
                <div style={{textAlign:'center'}}>TOTAL</div>
                <div className="del-col-por">USUARIO</div>
                <div style={{textAlign:'right'}}>ESTADO</div><div/>
              </div>
              {delFilterReceptor===REP_FILTER
                ? filteredRepRows.map(d => (
                    <div key={d.id} className="table-row del-cols clickable" onClick={() => setRepDetail(d._rep)}>
                      <div className="mono" style={{fontSize:12.5,color:'#6a6a62'}}>{d.fecha}</div>
                      <div style={{display:'flex',alignItems:'center',gap:11,minWidth:0}}>
                        <div className="avatar lg">{d.ini}</div>
                        <div style={{minWidth:0}}>
                          <div style={{fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.persona}</div>
                          {d.subLabel && <div style={{fontSize:11.5,color:'#8a8a82'}}>{d.subLabel}</div>}
                        </div>
                      </div>
                      <div className="del-col-detail" style={{color:'#6a6a62',fontSize:13,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.resumen}</div>
                      <div style={{textAlign:'center',fontWeight:700,fontFamily:'IBM Plex Mono,monospace'}}>{d.totalUd}</div>
                      <div className="del-col-por" style={{fontSize:12.5,color:'#8a8a82',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.creadoPor || '—'}</div>
                      <div/><div/>
                    </div>
                  ))
                : filteredDeliveryRows.map(d => (
                    <div key={d.id} className={`table-row ${delFilterReceptor==='Deportes Anexos'?'del-cols-disc':'del-cols'} clickable`} onClick={() => setSelectedDeliveryId(d.id)}>
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
                      {delFilterReceptor==='Deportes Anexos' && <div style={{fontSize:12.5,color:'#1a1a1a',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.disciplina||<span style={{color:'#ccc'}}>—</span>}</div>}
                      <div className="del-col-detail" style={{color:'#6a6a62',fontSize:13,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.resumen}</div>
                      <div style={{textAlign:'center',fontWeight:700,fontFamily:'IBM Plex Mono,monospace'}}>{d.totalUd}</div>
                      <div className="del-col-por" style={{fontSize:12.5,color:'#8a8a82',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.creadoPor || '—'}</div>
                      <div style={{textAlign:'right'}}>
                        {(() => { const st=d.status||'aceptado'; return st==='pendiente'?<span style={{background:'#FFF8D6',color:'#7a5800',border:'1px solid #FFD200',borderRadius:5,padding:'2px 7px',fontSize:11,fontWeight:700,whiteSpace:'nowrap'}}>Pendiente</span>:st==='rechazado'?<span style={{background:'#FBEAE8',color:'#C2473D',border:'1px solid #C2473D',borderRadius:5,padding:'2px 7px',fontSize:11,fontWeight:700,whiteSpace:'nowrap'}}>Rechazado</span>:<span style={{background:'#EDF7F2',color:'#2e9b5e',border:'1px solid #2e9b5e',borderRadius:5,padding:'2px 7px',fontSize:11,fontWeight:700,whiteSpace:'nowrap'}}>Aceptado</span> })()}
                      </div>
                      <div style={{display:'flex',justifyContent:'flex-end',alignItems:'center'}}>
                        {!isSoloVista && <button className="btn-del" onClick={e => { e.stopPropagation(); askDeleteDelivery(d.id) }}>✕</button>}
                      </div>
                    </div>
                  ))
              }
              {(delFilterReceptor===REP_FILTER ? filteredRepRows : filteredDeliveryRows).length === 0 && <div className="empty">{delFilterReceptor===REP_FILTER ? 'No hay reposiciones registradas.' : delFilterReceptor||delFilterPersona ? 'Sin entregas para este filtro.' : 'Sin entregas registradas.'}</div>}
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
                  <div className="mov-col-por">USUARIO</div>
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
                      <div className="mov-col-por" style={{fontSize:12.5,color:'#8a8a82',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{m.creadoPor || '—'}</div>
                      <div style={{textAlign:'right'}}>
                        <span className="mono" style={{fontWeight:700,fontSize:14,color:m.tipo==='entrada'?'#2e9b5e':'#C2473D'}}>
                          {m.tipo==='entrada'?'+':'−'}{m.qty}
                        </span>
                      </div>
                      <div style={{display:'flex',justifyContent:'flex-end'}}>
                        {!isSoloVista && <button className="btn-del" style={{width:28,height:28}} onClick={() => m.delId ? askDeleteDelivery(m.delId) : askDeleteMov(m.id)}>✕</button>}
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
              {[...db.users].sort((a,b) => (a.role==='admin'?0:1)-(b.role==='admin'?0:1)).map(u => (
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
                    {u.status==='pendiente' && currentUser?.role==='admin' && (
                      <div style={{display:'flex',gap:8,marginTop:10}}>
                        <button onClick={()=>approveUser(u.username)} style={{padding:'5px 14px',borderRadius:5,border:'none',cursor:'pointer',fontWeight:700,fontSize:12,background:'#FFD200',color:'#121212'}}>Aprobar</button>
                        <button onClick={()=>rejectUser(u.username)} style={{padding:'5px 14px',borderRadius:5,border:'1px solid #C2473D',cursor:'pointer',fontWeight:700,fontSize:12,background:'none',color:'#C2473D'}}>Rechazar</button>
                      </div>
                    )}
                  </div>
                  <span style={{background:u.role==='admin'?'#121212':u.role==='solo-vista'?'#FFF4E6':'#EDF7F2',color:u.role==='admin'?'#FFD200':u.role==='solo-vista'?'#c2560a':'#2e9b5e',border:'1px solid '+(u.role==='admin'?'#3a3a3a':u.role==='solo-vista'?'#e8834a':'#2e9b5e'),borderRadius:5,padding:'2px 8px',fontSize:11,fontWeight:700,flexShrink:0}}>
                    {u.role==='admin'?'Admin':u.role==='solo-vista'?'Solo Vista':'Receptor'}
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
              <div className="card" style={{padding:20,display:'flex',gap:14,alignItems:'center',cursor:'pointer'}} onClick={() => setView('reposiciones')}>
                <div className="avatar xl" style={{background:'#FFD200',color:'#121212'}}>R</div>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:15}}>Reposiciones Primera División</div>
                  <div style={{fontSize:12.5,color:'#8a8a82',marginTop:3}}>{(db.reposiciones||[]).length} reposiciones registradas</div>
                </div>
                <span style={{color:'#C8C8C0',fontSize:20}}>›</span>
              </div>
            </div>
          )}
          {/* CAMISETAS UTILERÍA */}
          {view === 'utileria' && (
            <div style={{display:'flex',flexDirection:'column',gap:10}}>
              <div className="kpi-card" style={{alignSelf:'flex-start',minWidth:180}}>
                <div className="kpi-label">CAMISETAS REGISTRADAS</div>
                <div className="kpi-value">{(db.camisetasUtileria||[]).length}</div>
                <div className="kpi-sub">en utilería</div>
              </div>
              <div style={{display:'flex',gap:8,alignItems:'flex-start',justifyContent:'space-between'}}>
                <div style={{display:'flex',flexDirection:'column',gap:6,flex:1}}>
                  {/* Filtro: Competición */}
                  <div style={{display:'flex',gap:5,flexWrap:'wrap',alignItems:'center'}}>
                    <span style={{fontSize:11,fontWeight:700,color:'#8a8a82',minWidth:82}}>COMPETICIÓN</span>
                    <button className={`chip${utiFilter===''?' active':''}`} onClick={()=>setUtiFilter('')}>Todas</button>
                    {COMPETICIONES.map(c => <button key={c} className={`chip${utiFilter===c?' active':''}`} onClick={()=>setUtiFilter(c)}>{c}</button>)}
                  </div>
                  {/* Filtro: Tipo */}
                  <div style={{display:'flex',gap:5,flexWrap:'wrap',alignItems:'center'}}>
                    <span style={{fontSize:11,fontWeight:700,color:'#8a8a82',minWidth:82}}>TIPO</span>
                    <button className={`chip${utiFilterTipo===''?' active':''}`} onClick={()=>setUtiFilterTipo('')}>Todos</button>
                    {['JUGADOR','GOLERO'].map(t => <button key={t} className={`chip${utiFilterTipo===t?' active':''}`} onClick={()=>setUtiFilterTipo(t)}>{t}</button>)}
                  </div>
                  {/* Filtro: Temporada */}
                  <div style={{display:'flex',gap:5,flexWrap:'wrap',alignItems:'center'}}>
                    <span style={{fontSize:11,fontWeight:700,color:'#8a8a82',minWidth:82}}>TEMPORADA</span>
                    <button className={`chip${utiFilterTemp===''?' active':''}`} onClick={()=>setUtiFilterTemp('')}>Todas</button>
                    {['2012/2013','2013/2014','2015/2016','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025','2026'].map(t => <button key={t} className={`chip${utiFilterTemp===t?' active':''}`} onClick={()=>setUtiFilterTemp(t)}>{t}</button>)}
                  </div>
                  {/* Filtro: Modelo */}
                  <div style={{display:'flex',gap:5,flexWrap:'wrap',alignItems:'center'}}>
                    <span style={{fontSize:11,fontWeight:700,color:'#8a8a82',minWidth:82}}>MODELO</span>
                    <button className={`chip${utiFilterModelo===''?' active':''}`} onClick={()=>setUtiFilterModelo('')}>Todos</button>
                    {[...new Set([...MODELOS_JUGADOR,...MODELOS_GOLERO])].map(m => <button key={m} className={`chip${utiFilterModelo===m?' active':''}`} onClick={()=>setUtiFilterModelo(m)}>{m}</button>)}
                  </div>
                </div>
                {!isSoloVista && <button className="btn btn-dark" style={{flexShrink:0,marginTop:2}} onClick={()=>{ setUtiForm({tipo:'',competicion:COMPETICIONES[0],numero:'',jugador:'',talle:'S',modelo:'',estampado:'',parches:'',detalle:'',temporada:'',id:null}); setUtiModal(true) }}>+ Camiseta</button>}
              </div>
              <div className="card" style={{overflow:'auto'}}>
                <div style={{display:'grid',gridTemplateColumns:'62px 90px 50px 46px 1fr 110px 1fr 44px 1fr 65px 28px',background:'#121212',padding:'9px 16px',gap:8,minWidth:860}}>
                  {['TIPO','MODELO','TEMP.','NRO.','NOMBRE','ESTAMPADO','COMPETICIÓN','TALLE','PARCHES','',''].map((h,i) => (
                    <div key={i} style={{fontSize:11,fontWeight:700,color:'#FFD200',letterSpacing:.5}}>{h}</div>
                  ))}
                </div>
                {utiFiltered.length === 0
                  ? <div style={{padding:28,textAlign:'center',color:'#8a8a82',fontSize:13}}>No hay camisetas que coincidan con los filtros.</div>
                  : utiFiltered.map(c => (
                      <div key={c.id} style={{display:'grid',gridTemplateColumns:'62px 90px 50px 46px 1fr 110px 1fr 44px 1fr 65px 28px',padding:'10px 16px',borderBottom:'1px solid #F0F0EC',alignItems:'center',gap:8,minWidth:860}}>
                        <div>{c.tipo && <span style={{fontSize:10,fontWeight:700,background:c.tipo==='GOLERO'?'#EDF7F2':'#F0F0EC',color:c.tipo==='GOLERO'?'#2e9b5e':'#5a5a52',border:'1px solid '+(c.tipo==='GOLERO'?'#2e9b5e':'#D0D0CA'),borderRadius:4,padding:'2px 5px'}}>{c.tipo}</span>}</div>
                        <div style={{fontSize:12,fontWeight:600,color:'#1a1a1a',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.modelo||<span style={{color:'#ccc'}}>—</span>}</div>
                        <div style={{fontSize:12,color:'#1a1a1a'}}>{c.temporada||<span style={{color:'#ccc'}}>—</span>}</div>
                        <div style={{fontWeight:800,fontSize:17,fontFamily:'IBM Plex Mono,monospace',color:'#1a1a1a'}}>{c.numero}</div>
                        <div style={{fontWeight:600,fontSize:13,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',color:'#1a1a1a'}}>
                          {c.jugador || <span style={{color:'#aaa',fontStyle:'italic',fontWeight:400}}>Sin asignar</span>}
                        </div>
                        <div style={{fontSize:12,color:'#1a1a1a',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.estampado||<span style={{color:'#ccc'}}>—</span>}</div>
                        <div style={{fontSize:12,color:'#1a1a1a',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.competicion||<span style={{color:'#ccc'}}>—</span>}</div>
                        <div style={{fontSize:13,fontWeight:700,textAlign:'center',color:'#1a1a1a'}}>{c.talle}</div>
                        <div style={{fontSize:12,color:'#1a1a1a',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.parches||<span style={{color:'#ccc'}}>—</span>}</div>
                        {!isSoloVista && <button className="btn btn-ghost" style={{padding:'4px 10px',fontSize:12}} onClick={()=>{setUtiForm({...c}); setUtiModal(true)}}>Editar</button>}
                        {!isSoloVista && <button onClick={()=>deleteUti(c.id)} style={{background:'none',border:'none',cursor:'pointer',fontSize:18,color:'#C2473D',padding:'0 4px',lineHeight:1}}>×</button>}
                        {c.detalle && (
                          <div style={{gridColumn:'1 / -1',fontSize:11,color:'#8a8a82',paddingTop:4,borderTop:'1px dashed #F0F0EC',marginTop:2}}>
                            {`Detalle: ${c.detalle}`}
                          </div>
                        )}
                      </div>
                    ))
                }
              </div>
            </div>
          )}

          {/* CONTRATO PUMA */}
          {view === 'contrato-puma' && (() => {
            const TOTAL_CONTRATO = 17200
            const RECEPTOR_ORDER = ['Protocolo','1° División','3 División','Juveniles','Femenino','Juveniles Femenino','Basket','Captación','Futbol Sala Masculino','Futbol Sala Femenino','Funcionarios']
            const RECEPTOR_COLORS = {
              'Protocolo':             '#7BC67E',
              '1° División':           '#FFD200',
              '3° División':           '#BDBDBD',
              'Juveniles':             '#5a5a5a',
              'Femenino':              '#9B59B6',
              'Juveniles Femenino':    '#C2185B',
              'Basket':                '#FF7043',
              'Captación':             '#4FC3F7',
              'Fútbol Sala Masculino': '#1565C0',
              'Fútbol Sala Femenino':  '#F48FB1',
              'Deportes Anexos':       '#E53935',
              'Funcionarios':          '#90A4AE',
            }
            const repUnidades = (db.reposiciones||[]).reduce((s, r) =>
              s + (r.jugadores||[]).reduce((a, j) => a + (Number(j.cantCamiseta)||0) + (Number(j.cantShort)||0), 0), 0)
            const baseData = receptorCards
              .map(r => {
                const extra = r.name === '1° División' ? repUnidades : 0
                const total = r.unidades + extra
                return { name: r.name, unidades: total, pct: total / TOTAL_CONTRATO * 100, monto: r.monto }
              })
            const data = baseData.sort((a, b) => {
                const ia = RECEPTOR_ORDER.indexOf(a.name)
                const ib = RECEPTOR_ORDER.indexOf(b.name)
                if (ia === -1 && ib === -1) return 0
                if (ia === -1) return 1
                if (ib === -1) return -1
                return ia - ib
              })
            const totalUsado = data.reduce((s, r) => s + r.unidades, 0)
            const pctTotal = totalUsado / TOTAL_CONTRATO * 100
            const maxPct = Math.max(...data.map(r => r.pct), 1)
            const BAR_HEIGHT = 220
            return (
              <div style={{display:'flex',flexDirection:'column',gap:20}}>
                {/* KPI global */}
                <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
                  <div className="kpi-card" style={{minWidth:180}}>
                    <div className="kpi-label">TOTAL CONTRATO</div>
                    <div className="kpi-value">{TOTAL_CONTRATO.toLocaleString('es-UY')}</div>
                    <div className="kpi-sub">unidades totales</div>
                  </div>
                  <div className="kpi-card" style={{minWidth:180}}>
                    <div className="kpi-label">UTILIZADO</div>
                    <div className="kpi-value">{totalUsado.toLocaleString('es-UY')}</div>
                    <div className="kpi-sub">{pctTotal.toFixed(1)}% del contrato</div>
                  </div>
                  <div className="kpi-card" style={{minWidth:180}}>
                    <div className="kpi-label">DISPONIBLE</div>
                    <div className="kpi-value">{(TOTAL_CONTRATO - totalUsado).toLocaleString('es-UY')}</div>
                    <div className="kpi-sub">{(100 - pctTotal).toFixed(1)}% restante</div>
                  </div>
                  <div className="kpi-card" style={{minWidth:180}}>
                    <div className="kpi-label">MONTO TOTAL ENTREGADO</div>
                    <div className="kpi-value" style={{fontSize:20}}>$ {data.reduce((s,r)=>s+(r.monto||0),0).toLocaleString('es-UY',{minimumFractionDigits:0,maximumFractionDigits:0})}</div>
                    <div className="kpi-sub">según precios del inventario</div>
                  </div>
                </div>

                {/* Gráfica de barras verticales */}
                {(() => {
                  const isMonto = pumaMetric === 'monto'
                  const maxVal = Math.max(...data.map(r => isMonto ? (r.monto||0) : r.unidades), 1)
                  const fmtVal = v => isMonto
                    ? '$ '+v.toLocaleString('es-UY',{minimumFractionDigits:0,maximumFractionDigits:0})
                    : v.toLocaleString('es-UY')+' uds.'
                  return (
                    <div className="card" style={{padding:'20px 24px'}}>
                      {/* Header con toggle */}
                      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
                        <div style={{fontWeight:700,fontSize:13}}>
                          {isMonto ? 'Monto entregado por receptor' : 'Artículos por receptor · % sobre '+TOTAL_CONTRATO.toLocaleString('es-UY')+' totales'}
                        </div>
                        <div style={{display:'flex',gap:4,background:'#F0F0EC',borderRadius:8,padding:3}}>
                          {[['unidades','Artículos'],['monto','Monto $']].map(([key,label]) => (
                            <button key={key} onClick={() => setPumaMetric(key)}
                              style={{padding:'4px 14px',borderRadius:6,border:'none',cursor:'pointer',fontSize:12,fontWeight:600,
                                background: pumaMetric===key ? '#121212' : 'transparent',
                                color: pumaMetric===key ? '#FFD200' : '#6a6a62',
                                transition:'all .15s'}}>
                              {label}
                            </button>
                          ))}
                        </div>
                      </div>
                      <div style={{display:'flex',alignItems:'flex-end',gap:12,height:BAR_HEIGHT+40,overflowX:'auto',paddingBottom:4}}>
                        {/* Eje Y */}
                        <div style={{display:'flex',flexDirection:'column',justifyContent:'space-between',height:BAR_HEIGHT,alignItems:'flex-end',flexShrink:0,paddingRight:6}}>
                          {[100,75,50,25,0].map(v => (
                            <span key={v} style={{fontSize:10,color:'#aaa',lineHeight:1}}>
                              {isMonto
                                ? '$ '+(maxVal*v/100/1000).toFixed(0)+'k'
                                : v+'%'}
                            </span>
                          ))}
                        </div>
                        {/* Líneas guía + barras */}
                        <div style={{position:'relative',flex:1,height:BAR_HEIGHT,minWidth:0}}>
                          {[0,25,50,75,100].map(v => (
                            <div key={v} style={{position:'absolute',left:0,right:0,bottom:`${v}%`,borderTop:'1px dashed #E8E8E0',zIndex:0}} />
                          ))}
                          <div style={{display:'flex',alignItems:'flex-end',gap:8,height:'100%',position:'relative',zIndex:1}}>
                            {data.map(r => {
                              const val = isMonto ? (r.monto||0) : r.unidades
                              const barH = (val / maxVal) * BAR_HEIGHT
                              const color = RECEPTOR_COLORS[r.name]||'#999'
                              return (
                                <div key={r.name}
                                  onClick={() => { setDelFilterReceptor(r.name); setView('entregas') }}
                                  style={{display:'flex',flexDirection:'column',alignItems:'center',flex:1,minWidth:40,height:'100%',justifyContent:'flex-end',cursor:'pointer'}}
                                  title={`Ver entregas de ${r.name}`}
                                >
                                  <div style={{fontSize:10,fontWeight:700,color,marginBottom:2,whiteSpace:'nowrap',textAlign:'center'}}>
                                    {isMonto
                                      ? (val>0 ? '$ '+val.toLocaleString('es-UY',{minimumFractionDigits:0,maximumFractionDigits:0}) : '—')
                                      : (r.pct.toFixed(1)+'%')}
                                  </div>
                                  <div style={{fontSize:9,fontWeight:700,color:'#1a1a1a',marginBottom:3,whiteSpace:'nowrap',textAlign:'center'}}>
                                    {isMonto ? (r.unidades>0 ? r.unidades.toLocaleString('es-UY')+' uds.' : '') : (r.unidades.toLocaleString('es-UY')+' uds.')}
                                  </div>
                                  <div style={{
                                    width:'100%', height: barH||2,
                                    background: color,
                                    borderRadius:'4px 4px 0 0',
                                    transition:'height .4s, filter .2s',
                                    minHeight: val>0 ? 4 : 2,
                                    opacity: val===0 ? 0.2 : 1
                                  }}
                                    onMouseEnter={e => e.currentTarget.style.filter='brightness(1.15)'}
                                    onMouseLeave={e => e.currentTarget.style.filter=''}
                                  />
                                </div>
                              )
                            })}
                          </div>
                        </div>
                      </div>
                      {/* Etiquetas eje X */}
                      <div style={{display:'flex',gap:8,marginLeft:42,marginTop:8}}>
                        {data.map(r => (
                          <div key={r.name} style={{flex:1,minWidth:40,display:'flex',flexDirection:'column',alignItems:'center',gap:3}}>
                            <div style={{width:8,height:8,borderRadius:'50%',background:RECEPTOR_COLORS[r.name]||'#999'}} />
                            <div style={{fontSize:10,fontWeight:700,color:'#1a1a1a',textAlign:'center',lineHeight:1.2,textTransform:'uppercase'}}>{r.name}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )
                })()}
              </div>
            )
          })()}

          {/* REPOSICIÓN CAMISETAS */}
          {view === 'reposiciones' && (
            <div style={{display:'flex',flexDirection:'column',gap:10}}>
              {/* Tabs */}
              <div style={{display:'flex',gap:6,borderBottom:'2px solid #ECECE8',paddingBottom:0}}>
                {[['reposiciones','Reposiciones'],['plantel','Plantel']].map(([k,l]) => (
                  <button key={k} onClick={() => setRepTab(k)} style={{padding:'7px 18px',border:'none',background:'none',fontWeight:700,fontSize:13,cursor:'pointer',borderBottom:repTab===k?'2px solid #FFD200':'2px solid transparent',marginBottom:-2,color:repTab===k?'#121212':'#8a8a82'}}>
                    {l}{k==='plantel'&&(db.plantel||[]).length>0?` (${(db.plantel||[]).length})`:''}
                  </button>
                ))}
              </div>

              {/* Tab: Reposiciones */}
              {repTab === 'reposiciones' && (<>
                {(() => {
                  const totalEquipos = (db.reposiciones||[]).reduce((acc,r)=>acc+(r.jugadores||[]).reduce((a,j)=>a+(j.descuento!==false?Number(j.cantCamiseta)||0:0),0),0)
                  const totalShorts  = (db.reposiciones||[]).reduce((acc,r)=>acc+(r.jugadores||[]).reduce((a,j)=>a+(j.descuento!==false?Number(j.cantShort)||0:0),0),0)
                  const totalCamTodas = (db.reposiciones||[]).reduce((acc,r)=>acc+(r.jugadores||[]).reduce((a,j)=>a+(Number(j.cantCamiseta)||0),0),0)
                  const totalShtTodas = (db.reposiciones||[]).reduce((acc,r)=>acc+(r.jugadores||[]).reduce((a,j)=>a+(Number(j.cantShort)||0),0),0)
                  return (
                    <div style={{display:'flex',alignItems:'flex-start',gap:12,flexWrap:'wrap'}}>
                      <div className="kpi-card" style={{alignSelf:'flex-start',minWidth:150}}>
                        <div className="kpi-label">ENTREGAS</div>
                        <div className="kpi-value">{(db.reposiciones||[]).length}</div>
                        <div className="kpi-sub">registradas</div>
                      </div>
                      <div className="kpi-card" style={{alignSelf:'flex-start',minWidth:150,cursor:'pointer'}} onClick={()=>setRepResumen('ambos')}>
                        <div className="kpi-label">CAMISETAS PARA DESCONTAR</div>
                        <div className="kpi-value">{totalEquipos}</div>
                        <div className="kpi-sub">camisetas en total →</div>
                      </div>
                      <div className="kpi-card" style={{alignSelf:'flex-start',minWidth:150,cursor:'pointer'}} onClick={()=>setRepResumen('ambos')}>
                        <div className="kpi-label">SHORTS PARA DESCONTAR</div>
                        <div className="kpi-value">{totalShorts}</div>
                        <div className="kpi-sub">shorts en total →</div>
                      </div>
                      <div className="kpi-card" style={{alignSelf:'flex-start',minWidth:150}}>
                        <div className="kpi-label">CAMISETAS ENVIADAS</div>
                        <div className="kpi-value">{totalCamTodas}</div>
                        <div className="kpi-sub">en todas las entregas</div>
                      </div>
                      <div className="kpi-card" style={{alignSelf:'flex-start',minWidth:150}}>
                        <div className="kpi-label">SHORTS ENVIADOS</div>
                        <div className="kpi-value">{totalShtTodas}</div>
                        <div className="kpi-sub">en todas las entregas</div>
                      </div>
                      <div className="kpi-card" style={{alignSelf:'flex-start',minWidth:150,cursor:'pointer'}} onClick={()=>setRepTab('plantel')}>
                        <div className="kpi-label">PLANTEL</div>
                        <div className="kpi-value">{(db.plantel||[]).length}</div>
                        <div className="kpi-sub">jugadores registrados →</div>
                      </div>
                      <button className="btn btn-dark" onClick={openRepModal} disabled={!(db.plantel||[]).length} style={{opacity:(db.plantel||[]).length?1:0.5,cursor:(db.plantel||[]).length?'pointer':'not-allowed',alignSelf:'flex-start'}}>+ Nueva reposición</button>
                    </div>
                  )
                })()}
                {!(db.plantel||[]).length && (
                  <div style={{fontSize:13,color:'#7a5800',background:'#FFF8D6',border:'1px solid #FFD200',borderRadius:8,padding:'10px 14px'}}>
                    Configurá el <button onClick={()=>setRepTab('plantel')} style={{background:'none',border:'none',fontWeight:700,color:'#7a5800',cursor:'pointer',padding:0,textDecoration:'underline'}}>plantel</button> primero para poder registrar reposiciones.
                  </div>
                )}
                {(db.reposiciones||[]).length === 0
                  ? <div style={{color:'#8a8a82',fontSize:14,textAlign:'center',padding:'40px 0'}}>No hay reposiciones registradas aún.</div>
                  : (() => {
                    const torneos = [...new Set((db.reposiciones||[]).map(r=>r.torneo).filter(Boolean))]
                    const filtered = (db.reposiciones||[]).filter(r => !repFilterTorneo || r.torneo === repFilterTorneo)
                    return (
                      <>
                        {torneos.length > 0 && (
                          <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                            <button className={`chip${repFilterTorneo===''?' active':''}`} onClick={() => setRepFilterTorneo('')}>Todos</button>
                            {torneos.map(t => (
                              <button key={t} className={`chip${repFilterTorneo===t?' active':''}`} onClick={() => setRepFilterTorneo(t)}>{t}</button>
                            ))}
                          </div>
                        )}
                        <div className="card" style={{padding:0,overflow:'hidden'}}>
                          <div className="table-header" style={{gridTemplateColumns:'100px 1fr 120px 70px 36px'}}>
                            <div>FECHA</div><div>CONCEPTO</div><div>TORNEO</div><div style={{textAlign:'right'}}>JUGADORES</div><div/>
                          </div>
                          {filtered.length === 0
                            ? <div style={{color:'#8a8a82',fontSize:13,textAlign:'center',padding:'24px 0'}}>Sin reposiciones para este torneo.</div>
                            : filtered.map(r => (
                              <div key={r.id} className="table-row" style={{gridTemplateColumns:'100px 1fr 120px 70px 36px',cursor:'pointer'}} onClick={() => setRepDetail(r)}>
                                <div style={{fontFamily:'IBM Plex Mono,monospace',fontSize:12,color:'#6a6a62'}}>{r.fecha}</div>
                                <div>
                                  <div style={{fontWeight:600}}>{r.concepto}</div>
                                  {r.creadoPor && <div style={{fontSize:11.5,color:'#8a8a82'}}>{r.creadoPor}</div>}
                                </div>
                                <div style={{fontSize:12}}>
                                  {r.torneo && <div style={{fontWeight:600}}>{r.torneo}</div>}
                                  {r.fechaTorneo != null && r.fechaTorneo !== '' && <div style={{fontSize:11,color:'#8a8a82'}}>Fecha {r.fechaTorneo}</div>}
                                </div>
                                <div style={{textAlign:'right',fontWeight:700,fontFamily:'IBM Plex Mono,monospace'}}>{(r.jugadores||[]).length}</div>
                                <div style={{textAlign:'right',color:'#8a8a82',fontSize:18,lineHeight:1}}>›</div>
                              </div>
                            ))
                          }
                        </div>
                      </>
                    )
                  })()
                }
              </>)}

              {/* Tab: Plantel */}
              {repTab === 'plantel' && (<>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                  <span style={{fontSize:12.5,color:'#8a8a82'}}>Jugadores con su talle de camiseta y short</span>
                  <button className="btn btn-dark" onClick={() => { setPlantelForm({id:null,numero:'',nombre:'',posicion:'Jugador',talleCamiseta:'L',talleShort:'L'}); setPlantelModal(true) }}>+ Jugador</button>
                </div>
                {(db.plantel||[]).length === 0
                  ? <div style={{color:'#8a8a82',fontSize:14,textAlign:'center',padding:'40px 0'}}>No hay jugadores en el plantel.</div>
                  : (
                    <div className="card" style={{padding:0,overflow:'hidden'}}>
                      <div className="table-header" style={{gridTemplateColumns:'50px 1fr 80px 90px 90px 72px'}}>
                        <div>Nº</div><div>NOMBRE</div><div>POSICIÓN</div><div>CAMISETA</div><div>SHORT</div><div/>
                      </div>
                      {(db.plantel||[]).sort((a,b)=>(Number(a.numero)||0)-(Number(b.numero)||0)).map(j => (
                        <div key={j.id} className="table-row" style={{gridTemplateColumns:'50px 1fr 80px 90px 90px 72px',
                          background: j.nombre.trim().toLowerCase()==='libre' ? '#3a3a3a' : j.posicion==='Golero' ? '#A5D6A7' : undefined}}>
                          <div style={{fontWeight:800,fontSize:15,color:j.nombre.trim().toLowerCase()==='libre'?'#888':'#1a1a1a'}}>{j.numero||'—'}</div>
                          <div style={{fontWeight:700,color:j.nombre.trim().toLowerCase()==='libre'?'#888':'#1a1a1a',fontStyle:j.nombre.trim().toLowerCase()==='libre'?'italic':undefined,textTransform:j.nombre.trim().toLowerCase()==='libre'?undefined:'uppercase'}}>{j.nombre}</div>
                          <div style={{color:'#1a1a1a',fontSize:12}}>{j.posicion||'Jugador'}</div>
                          <div style={{color:'#1a1a1a'}}>{j.talleCamiseta}</div>
                          <div style={{color:'#1a1a1a'}}>{j.talleShort}</div>
                          <div style={{display:'flex',gap:6,justifyContent:'flex-end'}}>
                            <button onClick={() => { setPlantelForm({...j}); setPlantelModal(true) }} style={{background:'none',border:'none',cursor:'pointer',color:'#8a8a82',fontSize:14,padding:'2px 4px'}}>✎</button>
                            <button onClick={() => deletePlantelJugador(j.id)} style={{background:'none',border:'none',cursor:'pointer',color:'#C2473D',fontSize:16,fontWeight:700,padding:'2px 4px'}}>×</button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )
                }
              </>)}
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
                  <div style={{fontSize:12.5,color:'#8a8a82',marginTop:2}}>
                    {d.receptor}
                    {d.receptor === 'Deportes Anexos' && (
                      disciplinaEdit !== null
                        ? <span style={{marginLeft:6,display:'inline-flex',gap:4,alignItems:'center'}}>
                            <select className="field-input" value={disciplinaEdit} onChange={e=>setDisciplinaEdit(e.target.value)}
                              autoFocus style={{fontSize:12,padding:'2px 6px',width:170}}>
                              <option value="">— Seleccioná —</option>
                              {DISCIPLINAS_DEPORTES_ANEXOS.map(d => <option key={d} value={d}>{d}</option>)}
                            </select>
                            <button onClick={()=>saveDisciplinaEdit(d.id)} style={{background:'#FFD200',border:'none',borderRadius:4,padding:'2px 8px',fontWeight:700,fontSize:11,cursor:'pointer'}}>✓</button>
                            <button onClick={()=>setDisciplinaEdit(null)} style={{background:'none',border:'none',cursor:'pointer',color:'#8a8a82',fontSize:13}}>✕</button>
                          </span>
                        : <span>
                            {d.disciplina ? ' · ' + d.disciplina : <span style={{color:'#C2473D',fontSize:11,marginLeft:4}}>sin disciplina</span>}
                            <button onClick={()=>setDisciplinaEdit(d.disciplina||'')} style={{background:'none',border:'none',cursor:'pointer',color:'#8a8a82',fontSize:13,marginLeft:4,padding:'0 2px'}}>✎</button>
                          </span>
                    )}
                    {d.receptor !== 'Deportes Anexos' && d.disciplina ? ' · ' + d.disciplina : ''}
                    {' · '}{d.fecha}
                  </div>
                  {d.creadoPor && <div style={{fontSize:12,color:'#aaa',marginTop:2}}>Registrado por: {d.creadoPor}</div>}
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
                {st === 'rechazado' && d.motivoRechazo && (
                  <div style={{padding:'10px 20px',background:'#FBEAE8',borderBottom:'1px solid #f5c6c3',fontSize:13,color:'#C2473D'}}>
                    <b>Motivo de rechazo:</b> {d.motivoRechazo}
                  </div>
                )}
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
                <label className="field-label">Tipo</label>
                <div style={{display:'flex',gap:8}}>
                  {['JUGADOR','GOLERO'].map(t => (
                    <button key={t} className={`talle-btn${utiForm.tipo===t?' active':''}`}
                      style={{flex:1,padding:'10px 0',fontSize:13,fontWeight:700}}
                      onClick={()=>setUtiForm(p=>({...p,tipo:t,modelo:''}))}>
                      {t}
                    </button>
                  ))}
                </div>
              </div>
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
                    {['S','M','L','XL'].map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
              </div>
              <div className="form-group">
                <label className="field-label">Jugador / Asignado</label>
                <input className="field-input" value={utiForm.jugador} onChange={e=>setUtiForm(p=>({...p,jugador:e.target.value}))} placeholder="Nombre del jugador" />
              </div>
              <div className="form-group">
                <label className="field-label">Estampado</label>
                <input className="field-input" value={utiForm.estampado} onChange={e=>setUtiForm(p=>({...p,estampado:e.target.value}))} placeholder="" />
              </div>
              <div className="form-group">
                <label className="field-label">Temporada</label>
                <select className="field-input" value={utiForm.temporada} onChange={e=>setUtiForm(p=>({...p,temporada:e.target.value}))}>
                  <option value="">Seleccionar…</option>
                  {['2012/2013','2013/2014','2015/2016','2016','2017','2018','2019','2020','2021','2022','2023','2024','2025','2026'].map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
              <div className="form-group">
                <label className="field-label">Modelo</label>
                <select className="field-input" value={utiForm.modelo} onChange={e=>setUtiForm(p=>({...p,modelo:e.target.value}))}>
                  <option value="">Seleccionar…</option>
                  {(utiForm.tipo==='GOLERO' ? MODELOS_GOLERO : MODELOS_JUGADOR).map(m => <option key={m} value={m}>{m}</option>)}
                </select>
              </div>
              <div className="form-group">
                <label className="field-label">Parches</label>
                <input className="field-input" value={utiForm.parches} onChange={e=>setUtiForm(p=>({...p,parches:e.target.value}))} placeholder="" />
              </div>
              <div className="form-group">
                <label className="field-label">Detalle</label>
                <input className="field-input" value={utiForm.detalle} onChange={e=>setUtiForm(p=>({...p,detalle:e.target.value}))} placeholder="" />
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={()=>setUtiModal(false)}>Cancelar</button>
              <button className="btn btn-dark" onClick={saveUti}>Guardar</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Nueva Reposición Camisetas */}
      {repModal && (
        <div className="modal-backdrop" onClick={() => setRepModal(false)}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{maxWidth:600,width:'96%'}}>
            <div className="modal-header">
              <div className="modal-title">{repForm.editId ? 'Editar reposición' : 'Nueva reposición'}</div>
              <button className="modal-close" onClick={() => setRepModal(false)}>×</button>
            </div>
            <div className="modal-body" style={{maxHeight:'70vh',overflowY:'auto'}}>
              <div className="form-group">
                <label className="field-label">Concepto</label>
                <input className="field-input" value={repForm.concepto} onChange={e => setRepForm(p=>({...p,concepto:e.target.value}))} placeholder="Ej. Reposición vs Nacional" autoFocus />
              </div>
              {/* Torneo y Fecha */}
              <div style={{display:'flex',gap:12,marginTop:4,alignItems:'flex-end'}}>
                <div className="form-group" style={{flex:1,marginBottom:0}}>
                  <label className="field-label">Torneo</label>
                  <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                    {['APERTURA','CLAUSURA','INTERMEDIO','COPA AUF'].map(t=>(
                      <button key={t} type="button" onClick={()=>setRepForm(p=>({...p,torneo:t}))}
                        style={{padding:'6px 10px',borderRadius:6,border:'2px solid',fontWeight:700,fontSize:11,cursor:'pointer',
                          borderColor:repForm.torneo===t?'#FFD200':'#ECECE8',
                          background:repForm.torneo===t?'#FFF8D6':'#fff',
                          color:repForm.torneo===t?'#7a5800':'#8a8a82'}}>
                        {t}
                      </button>
                    ))}
                  </div>
                </div>
                {TORNEOS_CON_FECHA.includes(repForm.torneo) ? (
                  <div className="form-group" style={{width:80,marginBottom:0}}>
                    <label className="field-label">Fecha</label>
                    <select className="field-input" value={repForm.fechaTorneo} onChange={e=>setRepForm(p=>({...p,fechaTorneo:e.target.value}))}>
                      {Array.from({length:15},(_,i)=>i+1).map(n=>(
                        <option key={n} value={n}>{n}</option>
                      ))}
                    </select>
                  </div>
                ) : (
                  <div className="form-group" style={{width:140,marginBottom:0}}>
                    <label className="field-label">Fecha</label>
                    <input className="field-input" value={repForm.fechaTorneo||''} onChange={e=>setRepForm(p=>({...p,fechaTorneo:e.target.value}))} placeholder="Ej. Ida / Final" />
                  </div>
                )}
              </div>
              {/* Selector de tipo de camiseta — uno por posición */}
              <div style={{display:'flex',gap:16,marginTop:14}}>
                <div style={{flex:1}}>
                  <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',marginBottom:6}}>EQUIPO JUGADORES</div>
                  <div style={{display:'flex',gap:6}}>
                    {REP_TIPOS_JUGADOR.map(t=>(
                      <button key={t} type="button" onClick={()=>setRepForm(p=>({...p,tipoCamisetaJugador:t}))}
                        style={{flex:1,padding:'6px 4px',borderRadius:6,border:'2px solid',fontWeight:700,fontSize:11,cursor:'pointer',
                          borderColor:repForm.tipoCamisetaJugador===t?'#FFD200':'#ECECE8',
                          background:repForm.tipoCamisetaJugador===t?'#FFF8D6':'#fff',
                          color:repForm.tipoCamisetaJugador===t?'#7a5800':'#8a8a82'}}>
                        {t}
                      </button>
                    ))}
                  </div>
                </div>
                <div style={{flex:1}}>
                  <div style={{fontSize:11,fontWeight:700,color:'#8a8a82',marginBottom:6}}>EQUIPO GOLEROS</div>
                  <div style={{display:'flex',gap:6}}>
                    {REP_TIPOS_GOLERO.map(t=>(
                      <button key={t} type="button" onClick={()=>setRepForm(p=>({...p,tipoCamisetaGolero:t}))}
                        style={{flex:1,padding:'6px 4px',borderRadius:6,border:'2px solid',fontWeight:700,fontSize:11,cursor:'pointer',
                          borderColor:repForm.tipoCamisetaGolero===t?'#FFD200':'#ECECE8',
                          background:repForm.tipoCamisetaGolero===t?'#FFF8D6':'#fff',
                          color:repForm.tipoCamisetaGolero===t?'#7a5800':'#8a8a82'}}>
                        {t}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <div style={{marginTop:16}}>
                <div style={{display:'grid',gridTemplateColumns:'40px 1fr 80px 80px 56px',gap:6,marginBottom:4,fontSize:10,fontWeight:700,color:'#8a8a82',padding:'4px 6px',background:'#F5F5F0',borderRadius:6}}>
                  <div>Nº</div><div>NOMBRE</div><div style={{textAlign:'center'}}>CAMISETA</div><div style={{textAlign:'center'}}>SHORT</div><div style={{textAlign:'center'}}>DESC.</div>
                </div>
                {repForm.rows.map((r, i) => {
                  const hasQty = Number(r.cantCamiseta)>0 || Number(r.cantShort)>0
                  const isLibre = r.nombre.trim().toLowerCase()==='libre'
                  return (
                    <div key={i} style={{display:'grid',gridTemplateColumns:'40px 1fr 80px 80px 56px',gap:6,marginBottom:3,alignItems:'center',padding:'5px 6px',borderRadius:6,
                      background:isLibre?'#3a3a3a':hasQty?'#FFFDF0':'transparent',border:hasQty?'1px solid #FFD200':'1px solid transparent'}}>
                      <div style={{fontFamily:'IBM Plex Mono,monospace',fontWeight:700,fontSize:13,color:isLibre?'#888':undefined}}>{r.numero||'—'}</div>
                      <div>
                        <span style={{fontWeight:600,fontSize:13,color:isLibre?'#888':undefined,fontStyle:isLibre?'italic':undefined}}>{r.nombre}</span>
                        {!isLibre && <span style={{fontSize:10,color:'#aaa',marginLeft:6}}>{r.posicion||'Jugador'}</span>}
                      </div>
                      <input className="field-input mono" type="number" min="0" value={r.cantCamiseta}
                        onChange={e=>setRepForm(p=>({...p,rows:p.rows.map((x,ix)=>ix===i?{...x,cantCamiseta:e.target.value}:x)}))}
                        placeholder="0" style={{textAlign:'center',padding:'5px 4px'}} />
                      <input className="field-input mono" type="number" min="0" value={r.cantShort}
                        onChange={e=>setRepForm(p=>({...p,rows:p.rows.map((x,ix)=>ix===i?{...x,cantShort:e.target.value}:x)}))}
                        placeholder="0" style={{textAlign:'center',padding:'5px 4px'}} />
                      {!isLibre ? (
                        <button type="button"
                          onClick={()=>setRepForm(p=>({...p,rows:p.rows.map((x,ix)=>ix===i?{...x,descuento:!x.descuento}:x)}))}
                          style={{padding:'4px 6px',borderRadius:5,border:'2px solid',fontWeight:700,fontSize:11,cursor:'pointer',width:'100%',
                            borderColor:r.descuento!==false?'#2d6a4f':'#ccc',
                            background:r.descuento!==false?'#d8f3dc':'#f5f5f5',
                            color:r.descuento!==false?'#1b4332':'#999'}}>
                          {r.descuento!==false?'SÍ':'NO'}
                        </button>
                      ) : <div/>}
                    </div>
                  )
                })}
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={() => setRepModal(false)}>Cancelar</button>
              <button className="btn btn-dark" onClick={saveReposicion}>{repForm.editId ? 'Guardar cambios' : 'Guardar reposición'}</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Detalle Reposición */}
      {repDetail && (
        <div className="modal-backdrop" onClick={() => setRepDetail(null)}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{maxWidth:580,width:'96%'}}>
            <div className="modal-header">
              <div style={{flex:1,minWidth:0}}>
                {repConceptoEdit !== null ? (
                  <div style={{display:'flex',gap:6,alignItems:'center'}}>
                    <input className="field-input" value={repConceptoEdit} onChange={e=>setRepConceptoEdit(e.target.value)}
                      onKeyDown={e=>{if(e.key==='Enter')saveRepConcepto();if(e.key==='Escape')setRepConceptoEdit(null)}}
                      autoFocus style={{fontWeight:700,fontSize:16,flex:1}} />
                    <button className="btn btn-dark" style={{padding:'6px 12px',fontSize:13}} onClick={saveRepConcepto}>✓</button>
                    <button className="btn btn-ghost" style={{padding:'6px 10px',fontSize:13}} onClick={()=>setRepConceptoEdit(null)}>✕</button>
                  </div>
                ) : (
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <div className="modal-title">{repDetail.concepto}</div>
                    <button onClick={()=>setRepConceptoEdit(repDetail.concepto)} style={{background:'none',border:'none',cursor:'pointer',color:'#8a8a82',fontSize:14,padding:'2px 4px',lineHeight:1}}>✎</button>
                  </div>
                )}
                <div style={{fontSize:12.5,color:'#8a8a82',marginTop:2}}>
                  {repDetail.fecha}{repDetail.creadoPor ? ' · '+repDetail.creadoPor : ''}
                  {repDetail.torneo && <span style={{marginLeft:8,fontWeight:700,color:'#7a5800',background:'#FFF8D6',border:'1px solid #FFD200',borderRadius:4,padding:'1px 7px',fontSize:11}}>
                    {repDetail.torneo}{repDetail.fechaTorneo ? ' · Fecha '+repDetail.fechaTorneo : ''}
                  </span>}
                </div>
              </div>
              <button className="modal-close" onClick={() => { setRepConceptoEdit(null); setRepDetail(null) }}>×</button>
            </div>
            <div className="modal-body" style={{maxHeight:'60vh',overflowY:'auto'}}>
              {/* Tipos de camiseta usados en esta reposición */}
              {(repDetail.tipoCamisetaJugador||repDetail.tipoCamisetaGolero) && (
                <div style={{display:'flex',gap:8,marginBottom:12,flexWrap:'wrap'}}>
                  {repDetail.tipoCamisetaJugador && (
                    <span style={{fontSize:12,fontWeight:700,background:'#FFF8D6',border:'1px solid #FFD200',borderRadius:5,padding:'3px 10px'}}>
                      Jugadores: {repDetail.tipoCamisetaJugador}
                    </span>
                  )}
                  {repDetail.tipoCamisetaGolero && (
                    <span style={{fontSize:12,fontWeight:700,background:'#F0F0EC',border:'1px solid #ccc',borderRadius:5,padding:'3px 10px'}}>
                      Goleros: {repDetail.tipoCamisetaGolero}
                    </span>
                  )}
                </div>
              )}
              <div style={{display:'grid',gridTemplateColumns:'40px 1fr 80px 80px',gap:6,marginBottom:4,fontSize:10,fontWeight:700,color:'#8a8a82',background:'#F5F5F0',borderRadius:6,padding:'5px 6px'}}>
                <div>Nº</div><div>NOMBRE</div><div style={{textAlign:'center'}}>CAMISETA</div><div style={{textAlign:'center'}}>SHORT</div>
              </div>
              {(repDetail.jugadores||[]).map((j,i) => {
                const goleroRowBg = j.posicion==='Golero'
                  ? j.tipoCamiseta==='NEGRO'   ? {background:'#d0d0d0',borderBottom:'1px solid #bbb'}
                  : j.tipoCamiseta==='NARANJA'  ? {background:'#FFE5CC',borderBottom:'1px solid #FFB870'}
                  : j.tipoCamiseta==='CREMA'    ? {background:'#FFF8E8',borderBottom:'1px solid #EDE0C0'}
                  : {borderBottom:'1px solid #F5F5F0'}
                  : {borderBottom:'1px solid #F5F5F0'}
                return (
                  <div key={i} style={{display:'grid',gridTemplateColumns:'40px 1fr 80px 80px',gap:6,padding:'6px 6px',fontSize:13,alignItems:'center',...goleroRowBg}}>
                    <div style={{fontFamily:'IBM Plex Mono,monospace',fontWeight:700,color:'#6a6a62'}}>{j.numero||'—'}</div>
                    <div>
                      <div style={{fontWeight:500}}>{j.nombre||'—'}</div>
                      {j.talleCamiseta && <div style={{fontSize:10,color:'#8a8a82',marginTop:1}}>CAM {j.talleCamiseta} · SHORT {j.talleShort}</div>}
                    </div>
                    <div style={{textAlign:'center',fontWeight:700,fontFamily:'IBM Plex Mono,monospace',color:j.cantCamiseta>0?'#1a1a1a':'#ccc'}}>
                      {j.cantCamiseta>0?j.cantCamiseta:'—'}
                    </div>
                    <div style={{textAlign:'center',fontWeight:700,fontFamily:'IBM Plex Mono,monospace',color:j.cantShort>0?'#1a1a1a':'#ccc'}}>
                      {j.cantShort>0?j.cantShort:'—'}
                    </div>
                  </div>
                )
              })}
              <div style={{marginTop:10,fontSize:12,color:'#8a8a82',textAlign:'right'}}>{(repDetail.jugadores||[]).length} jugadores</div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={() => setRepDetail(null)}>Cerrar</button>
              <button className="btn btn-ghost" style={{border:'1px solid #2d6a4f',color:'#2d6a4f'}} onClick={() => exportRepToExcel(repDetail)}>↓ Excel</button>
              <button className="btn btn-dark" onClick={() => openRepEdit(repDetail)}>Editar</button>
              <button style={{padding:'8px 16px',borderRadius:7,border:'1px solid #C2473D',background:'#FBEAE8',color:'#C2473D',fontWeight:700,cursor:'pointer'}}
                onClick={() => { if(window.confirm('¿Eliminar esta reposición?')) deleteReposicion(repDetail.id) }}>Eliminar</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Resumen entregas por jugador por mes */}
      {repResumen && (() => {
        const MESES_ES = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
        const esRep = r => /^reposici[oó]n\.?\s+vs/i.test((r.concepto||'').trim())
        const repsValidas = (db.reposiciones||[]).filter(esRep)

        // Agrupar por mes (clave MM/YYYY)
        const mesesMap = {}
        repsValidas.forEach(r => {
          const p = (r.fecha||'').split('/')
          const key = p.length===3 ? p[1]+'/'+p[2] : r.fecha||'?'
          if (!mesesMap[key]) mesesMap[key] = []
          mesesMap[key].push(r)
        })
        const mesesOrdenados = Object.keys(mesesMap).sort((a,b)=>{
          const [ma,ya] = a.split('/').map(Number)
          const [mb,yb] = b.split('/').map(Number)
          return ya!==yb ? ya-yb : ma-mb
        })

        return (
          <div className="modal-backdrop" onClick={()=>setRepResumen(null)}>
            <div className="modal" onClick={e=>e.stopPropagation()} style={{maxWidth:580,width:'96%'}}>
              <div className="modal-header">
                <div className="modal-title">Detalle por jugador</div>
                <button className="modal-close" onClick={()=>setRepResumen(null)}>×</button>
              </div>
              <div className="modal-body" style={{padding:'0 0 8px',maxHeight:'72vh',overflowY:'auto'}}>
                {mesesOrdenados.length === 0 && (
                  <div style={{padding:32,textAlign:'center',color:'#888',fontSize:14}}>Sin datos</div>
                )}
                {mesesOrdenados.map(mesKey => {
                  const [mm, yyyy] = mesKey.split('/')
                  const mesNombre = `${MESES_ES[Number(mm)]||mm} ${yyyy}`
                  const repsDelMes = mesesMap[mesKey]

                  const jugMapMes = {}
                  repsDelMes.forEach(r => (r.jugadores||[]).forEach(j => {
                    if (!jugMapMes[j.nombre]) jugMapMes[j.nombre] = {numero:j.numero||'—', nombre:j.nombre}
                  }))
                  const jugsMes = Object.values(jugMapMes).sort((a,b)=>(Number(a.numero)||0)-(Number(b.numero)||0))

                  const getCam = nombre => repsDelMes.reduce((acc,r)=>{const j=(r.jugadores||[]).find(x=>x.nombre===nombre);return acc+(j?Number(j.cantCamiseta)||0:0)},0)
                  const getSht = nombre => repsDelMes.reduce((acc,r)=>{const j=(r.jugadores||[]).find(x=>x.nombre===nombre);return acc+(j?Number(j.cantShort)||0:0)},0)

                  const filas = jugsMes.map(j=>({...j,cam:getCam(j.nombre),sht:getSht(j.nombre)})).filter(f=>f.cam+f.sht>0)
                  const totCam = filas.reduce((s,f)=>s+f.cam,0)
                  const totSht = filas.reduce((s,f)=>s+f.sht,0)

                  return (
                    <div key={mesKey}>
                      <div style={{background:'#121212',color:'#FFD200',fontWeight:700,fontSize:12,letterSpacing:'.06em',padding:'9px 14px',position:'sticky',top:0,zIndex:2}}>
                        {mesNombre.toUpperCase()}
                      </div>
                      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
                        <thead>
                          <tr style={{background:'#2a2a2a',color:'#ccc'}}>
                            <th style={{padding:'6px 12px',textAlign:'left',fontWeight:600,fontSize:11,letterSpacing:'.04em'}}>JUGADOR</th>
                            <th style={{padding:'6px 14px',textAlign:'center',fontWeight:600,fontSize:11,letterSpacing:'.04em'}}>CAMISETAS</th>
                            <th style={{padding:'6px 14px',textAlign:'center',fontWeight:600,fontSize:11,letterSpacing:'.04em'}}>SHORTS</th>
                            <th style={{padding:'6px 14px',textAlign:'center',fontWeight:600,fontSize:11,letterSpacing:'.04em'}}>TOTAL</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filas.map((f,i)=>(
                            <tr key={i} style={{borderBottom:'1px solid #F0F0EC',background:i%2===0?'#fff':'#FAFAF8'}}>
                              <td style={{padding:'7px 12px',fontWeight:500,whiteSpace:'nowrap'}}>
                                <span style={{fontFamily:'IBM Plex Mono,monospace',color:'#8a8a82',marginRight:8,fontSize:11}}>{f.numero}</span>
                                {f.nombre}
                              </td>
                              <td style={{padding:'7px 14px',textAlign:'center',fontFamily:'IBM Plex Mono,monospace',fontWeight:f.cam>0?700:400,color:f.cam>0?'#1a1a1a':'#ccc'}}>{f.cam>0?f.cam:'—'}</td>
                              <td style={{padding:'7px 14px',textAlign:'center',fontFamily:'IBM Plex Mono,monospace',fontWeight:f.sht>0?700:400,color:f.sht>0?'#1a1a1a':'#ccc'}}>{f.sht>0?f.sht:'—'}</td>
                              <td style={{padding:'7px 14px',textAlign:'center',fontFamily:'IBM Plex Mono,monospace',fontWeight:700,background:'#FFF8D6'}}>{f.cam+f.sht}</td>
                            </tr>
                          ))}
                          <tr style={{background:'#F5F2E8',borderTop:'2px solid #E8E4D8'}}>
                            <td style={{padding:'6px 12px',fontSize:11,fontWeight:700,letterSpacing:'.04em',color:'#555'}}>SUBTOTAL</td>
                            <td style={{padding:'6px 14px',textAlign:'center',fontFamily:'IBM Plex Mono,monospace',fontWeight:700,color:'#333'}}>{totCam||'—'}</td>
                            <td style={{padding:'6px 14px',textAlign:'center',fontFamily:'IBM Plex Mono,monospace',fontWeight:700,color:'#333'}}>{totSht||'—'}</td>
                            <td style={{padding:'6px 14px',textAlign:'center',fontFamily:'IBM Plex Mono,monospace',fontWeight:700,color:'#121212'}}>{totCam+totSht}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  )
                })}
              </div>
              <div className="modal-footer">
                <button className="btn btn-ghost" onClick={()=>setRepResumen(null)}>Cerrar</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* Modal: Plantel — agregar/editar jugador */}
      {plantelModal && (
        <div className="modal-backdrop" onClick={() => setPlantelModal(false)}>
          <div className="modal modal-sm" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">{plantelForm.id !== null ? 'Editar jugador' : 'Agregar jugador'}</div>
              <button className="modal-close" onClick={() => setPlantelModal(false)}>×</button>
            </div>
            <div className="modal-body">
              <div style={{display:'flex',gap:12}}>
                <div className="form-group" style={{width:80}}>
                  <label className="field-label">Número</label>
                  <input className="field-input mono" type="number" min="1" max="99" value={plantelForm.numero}
                    onChange={e => setPlantelForm(p=>({...p,numero:e.target.value}))}
                    placeholder="10" style={{textAlign:'center'}} />
                </div>
                <div className="form-group" style={{flex:1}}>
                  <label className="field-label">Nombre completo</label>
                  <input className="field-input" value={plantelForm.nombre}
                    onChange={e => setPlantelForm(p=>({...p,nombre:e.target.value}))}
                    placeholder="Ej. Maximiliano Olivera" autoFocus />
                </div>
              </div>
              <div style={{display:'flex',gap:8,marginTop:4}}>
                {['Jugador','Golero'].map(pos => (
                  <button key={pos} type="button"
                    onClick={() => setPlantelForm(p=>({...p,posicion:pos}))}
                    style={{flex:1,padding:'8px 0',borderRadius:6,border:'2px solid',fontWeight:700,fontSize:13,cursor:'pointer',
                      borderColor: plantelForm.posicion===pos ? '#FFD200' : '#ECECE8',
                      background: plantelForm.posicion===pos ? '#FFF8D6' : '#fff',
                      color: plantelForm.posicion===pos ? '#7a5800' : '#8a8a82'}}>
                    {pos}
                  </button>
                ))}
              </div>
              <div style={{display:'flex',gap:12,marginTop:4}}>
                <div className="form-group" style={{flex:1}}>
                  <label className="field-label">Talle Camiseta</label>
                  <select className="field-input" value={plantelForm.talleCamiseta} onChange={e => setPlantelForm(p=>({...p,talleCamiseta:e.target.value}))}>
                    {TALLES_ADULTO.map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
                <div className="form-group" style={{flex:1}}>
                  <label className="field-label">Talle Short</label>
                  <select className="field-input" value={plantelForm.talleShort} onChange={e => setPlantelForm(p=>({...p,talleShort:e.target.value}))}>
                    {TALLES_ADULTO.map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={() => setPlantelModal(false)}>Cancelar</button>
              <button className="btn btn-dark" onClick={savePlantelJugador}>Guardar</button>
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
                <label className="field-label">Fecha <span style={{fontSize:11,color:'#8a8a82',fontWeight:400}}>(opcional — por defecto hoy)</span></label>
                <input type="date" className="field-input" value={nd.fecha
                  ? nd.fecha.split('/').reverse().join('-')
                  : new Date().toISOString().slice(0,10)}
                  onChange={e => {
                    const [y,m,d] = e.target.value.split('-')
                    setNd(p=>({...p, fecha: d+'/'+m+'/'+y}))
                  }} />
              </div>
              <div className="form-group">
                <label className="field-label">Grupo / Plantel</label>
                <select className="field-input" value={nd.receptor} onChange={e => setNd(p=>({...p,receptor:e.target.value,paga:null,disciplina:''}))}>
                  <option value="">Seleccionar grupo…</option>
                  {RECEPTORES.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
              </div>
              {nd.receptor === 'Deportes Anexos' && (
                <div className="form-group">
                  <label className="field-label">Disciplina</label>
                  <select className="field-input" value={nd.disciplina} onChange={e => setNd(p=>({...p,disciplina:e.target.value}))}>
                    <option value="">— Seleccioná una disciplina —</option>
                    {DISCIPLINAS_DEPORTES_ANEXOS.map(d => <option key={d} value={d}>{d}</option>)}
                  </select>
                </div>
              )}
              {!ndIsDev && (
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
                  {receptorUsers.length === 0 && <div style={{marginTop:6,fontSize:12,color:'#8a8a82'}}>No hay usuarios receptores registrados aún.</div>}
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
                  {(() => {
                    const existing = na.code ? db.articles.find(a => a.code === na.code) : null
                    if (!existing?.ubic) return null
                    return (
                      <div style={{marginTop:6,fontSize:12,color:'#7a5800',background:'#FFF8D6',border:'1px solid #FFD200',borderRadius:6,padding:'6px 10px',display:'flex',justifyContent:'space-between',alignItems:'center',gap:8}}>
                        <span>Ya existe en <b>{existing.ubic}</b></span>
                        <button type="button" onClick={() => {
                          const ubic = existing.ubic
                          const isT = ubic === 'TRANSITO'
                          setNa(p => ({...p, estante: isT ? 'TRANSITO' : ubic.slice(0,-1), altura: isT ? 'A' : ubic.slice(-1)}))
                        }} style={{padding:'3px 10px',borderRadius:5,border:'1px solid #e6be00',background:'#FFD200',color:'#121212',fontWeight:700,fontSize:11.5,cursor:'pointer',whiteSpace:'nowrap'}}>
                          Usar ubicación
                        </button>
                      </div>
                    )
                  })()}
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
                  {na.estante !== 'TRANSITO' && (
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <span style={{fontSize:12,color:'#8a8a82',whiteSpace:'nowrap'}}>Altura</span>
                    <select className="field-input" style={{flex:1}} value={na.altura} onChange={e => setNa(p=>({...p,altura:e.target.value}))}>
                      {ALTURAS.map(o => <option key={o} value={o}>{o}</option>)}
                    </select>
                  </div>
                  )}
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
                  {mv.estante !== 'TRANSITO' && <>
                  <span style={{fontSize:12,color:'#8a8a82',whiteSpace:'nowrap'}}>Altura</span>
                  <select className="field-input" style={{flex:1}} value={mv.altura} onChange={e => setMv(p=>({...p,altura:e.target.value}))}>
                    {ALTURAS.map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                  </>}
                </div>
                <div style={{marginTop:6,fontSize:12,color:'#8a8a82'}}>Destino: <b style={{color:'#1a1a1a'}}>{mv.estante === 'TRANSITO' ? 'TRANSITO' : mv.estante + mv.altura}</b></div>
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
                      <span style={{background:u.role==='admin'?'#121212':u.role==='solo-vista'?'#FFF4E6':'#EDF7F2',color:u.role==='admin'?'#FFD200':u.role==='solo-vista'?'#c2560a':'#2e9b5e',border:'1px solid '+(u.role==='admin'?'#3a3a3a':u.role==='solo-vista'?'#e8834a':'#2e9b5e'),borderRadius:5,padding:'2px 8px',fontSize:11,fontWeight:700,flexShrink:0}}>{u.role==='admin'?'Admin':u.role==='solo-vista'?'Solo Vista':'Receptor'}</span>
                      {u.username === session && <span className="badge gray">Vos</span>}
                      {currentUser?.role === 'admin' && u.username !== session && <button className="btn-del" onClick={()=>deleteUser(u.username)}>✕</button>}
                    </div>
                    {currentUser?.role === 'admin' && u.username !== session && (
                      <div style={{display:'flex',gap:6,marginTop:8,paddingLeft:42}}>
                        {[['admin','Administrador'],['solo-vista','Solo Vista'],['receptor','Receptor']].map(([v,label]) => (
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
                  {currentUser?.role === 'admin' && (
                  <div className="form-group">
                    <label className="field-label">Rol</label>
                    <div style={{display:'flex',gap:8}}>
                      {[['admin','Administrador'],['solo-vista','Solo Vista'],['receptor','Receptor']].map(([v,label]) => (
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

      {rechazarModal.delId !== null && (
        <div className="modal-overlay" onClick={() => setRechazarModal({ delId: null, motivo: '' })}>
          <div className="modal-box" onClick={e => e.stopPropagation()} style={{maxWidth:420}}>
            <div className="modal-header">
              <div className="modal-title">Motivo de rechazo</div>
              <button className="modal-close" onClick={() => setRechazarModal({ delId: null, motivo: '' })}>×</button>
            </div>
            <div className="modal-body" style={{display:'flex',flexDirection:'column',gap:14}}>
              <p style={{margin:0,fontSize:13.5,color:'#6a6a62'}}>Explicá brevemente por qué rechazás esta entrega.</p>
              <div className="form-group">
                <label className="field-label">Motivo</label>
                <textarea
                  className="field-input"
                  rows={3}
                  style={{resize:'vertical',fontFamily:'inherit'}}
                  placeholder="Ej: Talle incorrecto, artículo dañado…"
                  value={rechazarModal.motivo}
                  onChange={e => setRechazarModal(p => ({...p, motivo: e.target.value}))}
                />
              </div>
              {!rechazarModal.motivo.trim() && (
                <div style={{fontSize:12,color:'#8a8a82'}}>El motivo es obligatorio para rechazar.</div>
              )}
              <div style={{display:'flex',gap:10}}>
                <button className="btn btn-ghost" style={{flex:1}} onClick={() => setRechazarModal({ delId: null, motivo: '' })}>Cancelar</button>
                <button
                  style={{flex:1,padding:'10px 0',borderRadius:8,border:'none',cursor: rechazarModal.motivo.trim() ? 'pointer':'not-allowed',fontWeight:700,fontSize:14,background: rechazarModal.motivo.trim() ? '#C2473D':'#e0a09a',color:'#fff'}}
                  onClick={() => { if(rechazarModal.motivo.trim()) receptorRechazar(rechazarModal.delId, rechazarModal.motivo.trim()) }}
                >✕ Rechazar entrega</button>
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
