import { useState, useEffect, useCallback, useRef } from 'react'
import { supabase } from './supabase.js'
import * as XLSX from 'xlsx'

const TALLE_ORDER = ['2','4','6','8','10','12','14','Único','S','M','L','XL','XXL','XXXL']
const TALLES_ADULTO = ['S','M','L','XL','XXL','XXXL','Único']
const TALLES_NINO   = ['2','4','6','8','10','12','14']
const RECEPTORES = ['1° División','3° División','Juveniles','Captación','Femenino','Juveniles Femenino','Fútbol Sala Masculino','Fútbol Sala Femenino','Basket','Deportes Anexos','Funcionarios','Protocolo']
const CATEGORIAS = ['Entrenamiento','Juego','Casual']
const ESTANTES = ['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20']
const ALTURAS = ['A','B','C','D','E','O']

const EMPTY_DB = { articles:[], deliveries:[], movimientos:[], nextId:1, nextDel:1, nextMov:1 }

async function loadFromSupabase() {
  const { data, error } = await supabase
    .from('deposito_state')
    .select('*')
    .eq('id', 1)
    .single()
  if (error || !data) return EMPTY_DB
  return {
    articles: (data.articles || []).map(a => ({
      ...a, sizes: (a.sizes || []).map(s => ({ talle: s.talle, qty: Number(s.qty)||0, min: Number(s.min)||0 }))
    })),
    deliveries: data.deliveries || [],
    movimientos: data.movimientos || [],
    nextId: data.next_id || 1,
    nextDel: data.next_del || 1,
    nextMov: data.next_mov || 1,
  }
}

async function saveToSupabase(db) {
  await supabase.from('deposito_state').upsert({
    id: 1,
    articles: db.articles,
    deliveries: db.deliveries,
    movimientos: db.movimientos,
    next_id: db.nextId,
    next_del: db.nextDel,
    next_mov: db.nextMov,
    updated_at: new Date().toISOString(),
  })
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
  const [search, setSearch] = useState('')
  const [cat, setCat] = useState('Todas')
  const [modal, setModal] = useState(null)
  const [confirm, setConfirm] = useState(null)
  const [editing, setEditing] = useState(null)
  const [movFilter, setMovFilter] = useState('Todos')
  const [delFilterReceptor, setDelFilterReceptor] = useState('')
  const [delFilterPersona, setDelFilterPersona] = useState('')
  const [toast, setToast] = useState('')
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const toastTimer = useRef(null)
  const saveTimer = useRef(null)

  // delivery/devolución form
  const [nd, setNd] = useState({ mode:'entrega', persona:'', receptor:'', cCode:'', cSearch:'', cTalle:'', cQty:'', paga:null, lines:[] })
  // new article form
  const [na, setNa] = useState({ code:'', name:'', cat:'Entrenamiento', tipo:'adulto', precio:'', tallesArr:[], tallesMins:{}, tallesQty:{}, estante:'1', altura:'A' })
  // reponer form
  const [rep, setRep] = useState({ qtys:{} })
  // ajuste form
  const [aj, setAj] = useState({ talle:'', cantidad:'' })
  // mover form
  const [mv, setMv] = useState({ tallesArr:[], estante:'1', altura:'A' })

  // Load from Supabase on mount (filter out articles with no stock)
  useEffect(() => {
    loadFromSupabase().then(data => {
      setDb({...data, articles: data.articles.filter(a => total(a) > 0)})
      setLoading(false)
    })
  }, [])

  // Save to Supabase whenever data changes (debounced 800ms)
  useEffect(() => {
    if (loading) return
    clearTimeout(saveTimer.current)
    saveTimer.current = setTimeout(() => saveToSupabase(db), 800)
  }, [db, loading])

  // Redirect to inventario if selected article was deleted (reached stock 0)
  useEffect(() => {
    if (view === 'detalle' && selectedId !== null && !db.articles.find(a => a.id === selectedId)) {
      setView('inventario')
    }
  }, [db.articles, view, selectedId])

  const showToast = useCallback((msg) => {
    setToast(msg)
    clearTimeout(toastTimer.current)
    toastTimer.current = setTimeout(() => setToast(''), 2600)
  }, [])

  const closeModal = () => setModal(null)
  const byCode = (code) => db.articles.find(a => a.code === code)
  const curCode = () => { const a = db.articles.find(x => x.id === selectedId); return a ? a.code : '' }

  const goView = (v) => { setView(v); setSearch(''); setSidebarOpen(false) }
  const openDetail = (id) => { setSelectedId(id); setView('detalle'); setSidebarOpen(false) }

  // ---- Entregas / Devoluciones ----
  const openEntrega = () => { setNd({ mode:'entrega', persona:'', receptor:'', cCode:'', cSearch:'', cTalle:'', cQty:'', paga:null, lines:[] }); setModal('entrega') }
  const openDevolucion = () => { setNd({ mode:'devolucion', persona:'', receptor:'', cCode:'', cSearch:'', cTalle:'', cQty:'', paga:null, lines:[] }); setModal('entrega') }
  const openEntregaFromDetail = () => { const a = byCode(curCode()); setNd({ mode:'entrega', persona:'', receptor:'', cCode:a?a.code:'', cSearch:'', cTalle:'', cQty:'', lines:[] }); setModal('entrega') }
  const openDevolucionFromDetail = () => { const a = byCode(curCode()); setNd({ mode:'devolucion', persona:'', receptor:'', cCode:a?a.code:'', cSearch:'', cTalle:'', cQty:'', lines:[] }); setModal('entrega') }

  const ndAddLine = () => {
    const qty = parseInt(nd.cQty, 10)
    if(!nd.cCode || !nd.cTalle || !qty || qty <= 0) { showToast('Completá artículo, talle y cantidad.'); return }
    if(nd.mode !== 'devolucion') {
      const a = byCode(nd.cCode); const sz = a && a.sizes.find(s => s.talle === nd.cTalle)
      const already = nd.lines.filter(l => l.code === nd.cCode && l.talle === nd.cTalle).reduce((s,l) => s+l.qty, 0)
      if(!sz || qty + already > sz.qty) { showToast('Stock insuficiente ('+(sz?sz.qty-already:0)+' disp.).'); return }
    }
    setNd(p => ({...p, lines:[...p.lines,{code:nd.cCode,talle:nd.cTalle,qty}], cCode:'', cSearch:'', cTalle:'', cQty:''}))
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
        const a = articles.find(x => x.code === l.code); const z = a && a.sizes.find(x => x.talle === l.talle)
        if(z) z.qty = esDev ? z.qty + l.qty : Math.max(0, z.qty - l.qty)
        if(esDev) {
          movimientos.unshift({id:mid++, code:l.code, name:a?.name||l.code, tipo:'entrada', fecha, talle:l.talle, qty:l.qty, detalle:'Devolución de '+nd.persona+' ('+nd.receptor+')'})
        } else {
          movimientos.unshift({id:mid++, code:l.code, name:a?.name||l.code, tipo:'salida', fecha, talle:l.talle, qty:l.qty, detalle:'Entrega a '+nd.persona+' ('+nd.receptor+')', delId:s.nextDel})
        }
      })
      const activeArticles = articles.filter(a => total(a) > 0)
      if(esDev) return { ...s, articles:activeArticles, movimientos, modal:null, nextMov:mid }
      const deliveries = [{id:s.nextDel, fecha, persona:nd.persona.trim(), receptor:nd.receptor, paga:nd.receptor==='Protocolo'?nd.paga:null, monto:nd.receptor==='Protocolo'&&nd.paga==='si'?ndMonto:null, lines:[...nd.lines]}, ...s.deliveries]
      return { ...s, articles:activeArticles, movimientos, deliveries, nextDel:s.nextDel+1, nextMov:mid }
    })
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
    const code = curCode(); const fecha = today()
    setDb(s => {
      let nextMov = s.nextMov
      const artName = (s.articles.find(a => a.code === code)||{}).name || code
      const articles = s.articles.map(a => {
        if(a.code !== code) return a
        return {...a, sizes: a.sizes.map(z => { const e=entries.find(e=>e.talle===z.talle); return e ? {...z, qty:z.qty+e.q} : z })}
      })
      const newMovs = entries.map(e => ({id:nextMov++, code, name:artName, tipo:'entrada', fecha, talle:e.talle, qty:e.q, detalle:'Ingreso de stock'}))
      return { ...s, articles, movimientos:[...newMovs,...s.movimientos], nextMov }
    })
    setModal(null)
    const total = entries.reduce((s,e)=>s+e.q, 0)
    showToast('Entrada registrada: +'+total+' u. en '+entries.length+' talle'+(entries.length>1?'s':'')+'.')
  }

  // ---- Ajuste ----
  const openAjuste = () => { setAj({ talle:'', cantidad:'' }); setModal('ajuste') }
  const ajConfirm = () => {
    if(!aj.talle || aj.cantidad === '') { showToast('Elegí talle e ingresá la cantidad contada.'); return }
    const q = parseInt(aj.cantidad, 10)
    if(isNaN(q) || q < 0) { showToast('Cantidad inválida.'); return }
    const code = curCode(); const fecha = today()
    const art = byCode(code); const z0 = art && art.sizes.find(z => z.talle === aj.talle); const cur = z0 ? z0.qty : 0
    const delta = q - cur
    if(delta === 0) { showToast('Sin cambios: el stock ya es '+q+'.'); setModal(null); return }
    setDb(s => {
      const artName = (s.articles.find(a => a.code === code)||{}).name || code
      const articles = s.articles.map(a => { if(a.code!==code) return a; return {...a, sizes:a.sizes.map(z => z.talle===aj.talle?{...z,qty:Math.max(0,q)}:z)} })
      const activeArticles = articles.filter(a => total(a) > 0)
      const movimientos = [{id:s.nextMov, code, name:artName, tipo:(delta>0?'entrada':'salida'), fecha, talle:aj.talle, qty:Math.abs(delta), detalle:'Ajuste por recuento (de '+cur+' a '+q+')'}, ...s.movimientos]
      return { ...s, articles:activeArticles, movimientos, nextMov:s.nextMov+1 }
    })
    setModal(null)
    showToast('Stock ajustado: '+aj.talle+' = '+q+' ('+(delta>0?'+':'')+delta+').')
  }

  // ---- Mover talle ----
  const openMover = () => { setMv({ tallesArr:[], estante:'1', altura:'A' }); setModal('mover') }

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
    if(mv.tallesArr.length === 0) { showToast('Seleccioná al menos un talle.'); return }
    const newUbic = mv.estante + mv.altura
    if(newUbic === selA.ubic) { showToast('La ubicación destino es la misma que la actual.'); return }
    setDb(prev => {
      const code = curCode()
      let arts = prev.articles
      let nextId = prev.nextId
      const targetEntry = arts.find(a => a.code === code && a.ubic === newUbic && a.id !== selA.id)
      if(targetEntry) {
        arts = arts.map(a => {
          if(a.id !== targetEntry.id) return a
          const newSizes = [...a.sizes]
          mv.tallesArr.forEach(t => {
            const src = selA.sizes.find(sz => sz.talle === t)
            const idx = newSizes.findIndex(sz => sz.talle === t)
            if(idx >= 0) newSizes[idx] = {...newSizes[idx], qty: newSizes[idx].qty + (src?.qty||0)}
            else newSizes.push({...src})
          })
          return {...a, sizes: newSizes}
        })
      } else {
        const movedSizes = selA.sizes.filter(sz => mv.tallesArr.includes(sz.talle))
        arts = [...arts, {id:nextId++, code, name:selA.name, cat:selA.cat, ubic:newUbic, sizes:movedSizes}]
      }
      const remaining = selA.sizes.filter(sz => !mv.tallesArr.includes(sz.talle))
      arts = remaining.length === 0
        ? arts.filter(a => a.id !== selA.id)
        : arts.map(a => a.id === selA.id ? {...a, sizes:remaining} : a)
      return {...prev, articles:arts, nextId}
    })
    setModal(null); setView('inventario')
    showToast('Talle(s) movido(s) a ' + mv.estante + mv.altura + '.')
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

  // ---- Derived data ----
  const { articles, deliveries, movimientos } = db
  const codeName = Object.fromEntries(articles.map(a => [a.code, a.name]))

  const kpis = {
    articulos: articles.length,
    unidades: fmt(articles.reduce((s,a) => s + total(a), 0)),
    valorStock: articles.reduce((s,a) => s + (a.precio||0) * total(a), 0),
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

  const codeTalleCounts = {}
  articles.forEach(a => a.sizes.forEach(s => {
    const k = a.code + ':' + s.talle
    codeTalleCounts[k] = (codeTalleCounts[k]||0)+1
  }))
  const dupArticleIds = new Set(
    articles.filter(a => a.sizes.some(s => codeTalleCounts[a.code+':'+s.talle] > 1)).map(a => a.id)
  )
  const dupCodes = new Set(articles.filter(a => dupArticleIds.has(a.id)).map(a => a.code))
  const dupList = [...dupCodes].map(code => {
    const entries = articles.filter(a => a.code === code)
    const talleCounts = {}
    entries.forEach(a => a.sizes.forEach(s => { talleCounts[s.talle] = (talleCounts[s.talle]||0)+1 }))
    const tallesDup = Object.keys(talleCounts).filter(t => talleCounts[t] > 1)
    return { code, name: entries[0].name, entries, tallesDup }
  })

  const invRows = filtered.map(a => ({
    ...a,
    totalFmt: fmt(total(a)),
    sizesLabel: sizesLabel(a),
    low: isLow(a),
    ubic: a.ubic||'—',
    dupUbic: dupArticleIds.has(a.id),
    precio: a.precio||0,
  }))

  const lowList = articles.filter(isLow).map(a => ({
    ...a,
    tallesEnMin: a.sizes.filter(s => (s.min||0) > 0 && s.qty === s.min).length,
    tallesBajo:  a.sizes.filter(s => (s.min||0) > 0 && s.qty < s.min).length,
  }))

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

  const receptorCards = RECEPTORES.map(name => {
    const ds = deliveries.filter(d => d.receptor===name)
    const unidades = ds.reduce((s,d) => s+d.lines.reduce((x,l)=>x+l.qty,0),0)
    return { name, ini:ini(name), count:ds.length, unidades }
  })

  const selA = articles.find(a => a.id === selectedId)
  let detail = null
  if(selA) {
    const tot = total(selA); const low = isLow(selA)
    const maxQ = Math.max(1, ...selA.sizes.map(s=>s.qty))
    const ordered = [...selA.sizes].sort((a,b) => TALLE_ORDER.indexOf(a.talle)-TALLE_ORDER.indexOf(b.talle))
    const sizes = ordered.map(s => {
      const sLow = (s.min||0)>0 && s.qty<=(s.min||0)
      return {...s, isLow:sLow, pct:Math.round(s.qty/maxQ*100)}
    })
    const movs = movimientos.filter(m => m.code===selA.code)
    detail = { ...selA, total:tot, totalFmt:fmt(tot), low, sizesLabel:sizesLabel(selA), ubic:selA.ubic||'—', sizes, movs, noMovs:movs.length===0 }
  }

  // nd derived
  const ndA = byCode(nd.cCode)
  const ndTalleOptions = ndA ? ndA.sizes.map(s => ({value:s.talle, label:s.talle+' ('+s.qty+' disp.)'})) : []
  let stockHint = ''
  if(nd.cCode && nd.cTalle && ndA) { const z=ndA.sizes.find(s=>s.talle===nd.cTalle); if(z) stockHint='Disponible: '+z.qty+' u. en talle '+nd.cTalle }
  const ndTotal = nd.lines.reduce((s,l) => s+l.qty, 0)
  const ndMonto = nd.receptor === 'Protocolo' && nd.paga === 'si'
    ? nd.lines.reduce((s,l) => { const art=articles.find(a=>a.code===l.code); return s+(art?.precio||0)*l.qty }, 0) * 0.5
    : 0
  const ndOk = nd.persona && nd.persona.trim() && nd.receptor && nd.lines.length > 0

  const repTalleOptions = selA ? selA.sizes.map(s => ({value:s.talle, label:s.talle})) : []
  const ajTalleOptions = selA ? selA.sizes.map(s => ({value:s.talle, label:s.talle+' (sistema: '+s.qty+')'})) : []

  const ndIsDev = nd.mode === 'devolucion'

  if (loading) return (
    <div style={{display:'flex',alignItems:'center',justifyContent:'center',height:'100dvh',flexDirection:'column',gap:16,background:'#121212'}}>
      <img src="/escudo.png" alt="Peñarol" style={{height:64,opacity:.9}} />
      <div style={{color:'#FFD200',fontFamily:'Archivo Black,sans-serif',fontSize:14,letterSpacing:'.1em'}}>CARGANDO…</div>
    </div>
  )

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
          {[['panel','PANEL'],['inventario','INVENTARIO'],['entregas','ENTREGAS'],['movimientos','MOVIMIENTOS'],['receptores','RECEPTORES']].map(([key,label]) => (
            <button key={key} className={`nav-item${view===key||(key==='inventario'&&view==='detalle')?' active':''}`} onClick={() => goView(key)}>
              <span className="nav-dot" />
              {label}
            </button>
          ))}
        </nav>
        <div className="sidebar-user">
          <div className="user-avatar">CP</div>
          <div>
            <div className="user-name">COMPRAS PEÑAROL</div>
            <div className="user-role">Gestión de depósito</div>
          </div>
        </div>
      </aside>

      {/* Main */}
      <div className="main-area">
        <header className="topbar">
          <button className="hamburger" onClick={() => setSidebarOpen(o=>!o)} aria-label="Menú">
            <span/><span/><span/>
          </button>
          <div className="topbar-title">
            {{panel:'PANEL',inventario:'INVENTARIO',detalle:'DETALLE',entregas:'ENTREGAS',movimientos:'MOVIMIENTOS',receptores:'RECEPTORES'}[view]}
          </div>
          <div className="topbar-spacer" />
          <div className="search-box">
            <span className="search-icon" />
            <input value={search} onChange={e => { setSearch(e.target.value); if((view==='panel'||view==='detalle')&&e.target.value) setView('inventario') }} placeholder="Buscar…" />
          </div>
          <button className="btn btn-ghost" onClick={openArticulo}><span>+</span><span> Artículo</span></button>
          <button className="btn btn-ghost" onClick={openDevolucion}><span>↩</span><span> Dev.</span></button>
          <button className="btn btn-yellow" onClick={openEntrega}><span>+</span><span> Entrega</span></button>
        </header>

        <div className="content">
          {/* PANEL */}
          {view === 'panel' && (
            <>
              <div className="kpi-grid">
                <div className="kpi-card"><div className="kpi-label">ARTÍCULOS</div><div className="kpi-value">{kpis.articulos}</div><div className="kpi-sub">referencias activas</div></div>
                <div className="kpi-card"><div className="kpi-label">UNIDADES EN STOCK</div><div className="kpi-value">{kpis.unidades}</div><div className="kpi-sub">suma de todos los talles</div>{kpis.valorStock > 0 && <div className="kpi-sub" style={{marginTop:6,fontWeight:700}}>$ {kpis.valorStock.toLocaleString('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2})} valor total</div>}</div>
                <div className="kpi-card"><div className="kpi-label">ENTREGAS</div><div className="kpi-value">{kpis.entregas}</div><div className="kpi-sub">en el historial</div></div>
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
                    <div key={a.id} className="table-row clickable" style={{gridTemplateColumns:'1fr auto'}} onClick={() => openDetail(a.id)}>
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
                    <div className="card-title">Entregas recientes</div>
                    <div className="card-spacer"/>
                    <button className="back-link" style={{color:'#9a7d00',margin:0}} onClick={() => goView('entregas')}>Ver todas →</button>
                  </div>
                  {recentDeliveries.length === 0 && <div className="empty">Sin entregas registradas.</div>}
                  {recentDeliveries.map(d => (
                    <div key={d.id} className="table-row" style={{gridTemplateColumns:'34px 1fr auto'}}>
                      <div className="avatar">{d.ini}</div>
                      <div style={{minWidth:0}}>
                        <div style={{fontWeight:600,fontSize:13.5,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.persona}</div>
                        <div style={{fontSize:11.5,color:'#8a8a82',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.receptor} · {d.resumen}</div>
                      </div>
                      <div style={{textAlign:'right',flexShrink:0}}>
                        <div style={{fontWeight:700,fontSize:14}}>{d.totalUd}</div>
                        <div style={{fontSize:11,color:'#8a8a82'}}>{d.fecha}</div>
                      </div>
                    </div>
                  ))}
                </div>
                {dupList.length > 0 && (
                <div className="card">
                  <div className="card-header">
                    <div className="card-title">Artículos en ubicaciones múltiples</div>
                    <div className="card-spacer"/>
                    <span className="badge" style={{background:'#FFF0C2',color:'#7a5800',border:'1px solid #FFD200'}}>{dupList.length} artículo{dupList.length>1?'s':''}</span>
                  </div>
                  {dupList.map(d => (
                    <div key={d.code} className="table-row clickable" style={{gridTemplateColumns:'1fr auto'}} onClick={() => openDetail(d.entries[0].id)}>
                      <div>
                        <div style={{fontWeight:600,fontSize:13.5}}>{d.name}</div>
                        <div style={{fontSize:11.5,color:'#8a8a82',fontFamily:'IBM Plex Mono,monospace'}}>{d.code}</div>
                        <div style={{fontSize:12,color:'#6a6a62',marginTop:3}}>Ubicaciones: <b>{d.entries.map(a=>a.ubic||'—').join(' · ')}</b></div>
                      </div>
                      {d.tallesDup.length > 0 && (
                        <div style={{textAlign:'right',flexShrink:0}}>
                          <div style={{fontSize:11,color:'#8a8a82',marginBottom:2}}>Talle{d.tallesDup.length>1?'s':''} duplicado{d.tallesDup.length>1?'s':''}</div>
                          <div style={{fontWeight:700,fontSize:13,color:'#C2473D'}}>{d.tallesDup.join(', ')}</div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
                )}
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
                  <div key={r.id} className="table-row clickable inv-cols" onClick={() => openDetail(r.id)}>
                    <div className="mono" style={{fontSize:12.5,color:'#6a6a62',fontWeight:500}}>{r.code}</div>
                    <div style={{fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.name}</div>
                    <div className="inv-col-ubic"><span className="ubic-badge">{r.ubic}</span></div>
                    <div className="inv-col-cat" style={{color:'#6a6a62'}}>{r.cat}</div>
                    <div className="inv-col-sizes" style={{color:'#6a6a62'}}>{r.sizesLabel}</div>
                    <div style={{textAlign:'right',fontWeight:700,fontFamily:'IBM Plex Mono,monospace'}}>{r.totalFmt}</div>
                    <div style={{textAlign:'right',display:'flex',gap:4,justifyContent:'flex-end',flexWrap:'wrap'}}>
                      {r.dupUbic && <span className="badge" style={{background:'#FFF0C2',color:'#7a5800',border:'1px solid #FFD200'}}>⚠ Art. duplicado</span>}
                      {r.low && <span className="badge low">Bajo mín.</span>}
                    </div>
                    <div className="inv-col-precio mono" style={{textAlign:'right',fontSize:12.5,color:'#6a6a62'}}>
                      {r.precio > 0 ? '$ '+r.precio.toLocaleString('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2}) : '—'}
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
                  <div style={{padding:'22px 24px',borderBottom:'1px solid #E7E7E3'}}>
                    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',gap:16}}>
                      <div style={{flex:1}}>
                        <div className="mono" style={{fontSize:12.5,color:'#8a8a82'}}>{detail.code}</div>
                        <div style={{fontWeight:800,fontSize:22,marginTop:3}}>{detail.name}</div>
                        <div style={{display:'flex',gap:8,marginTop:9,alignItems:'center',flexWrap:'wrap'}}>
                          <span className="badge gray">{detail.cat}</span>
                          {detail.low && <span className="badge low">Bajo mínimo</span>}
                          <span className="ubic-badge"><span style={{fontSize:11,color:'#9a7d00',fontFamily:'Archivo,sans-serif'}}>UBIC. </span>{detail.ubic}</span>
                        </div>
                      </div>
                      <div style={{textAlign:'right',flexShrink:0}}>
                        <div style={{fontFamily:'Archivo Black,sans-serif',fontSize:30,lineHeight:1}}>{detail.totalFmt}</div>
                        <div style={{fontSize:11.5,color:'#8a8a82',marginTop:4}}>unidades totales</div>
                        {detail.precio > 0 && <div style={{marginTop:8,fontSize:13,fontWeight:700,color:'#1a1a1a'}}>$ {detail.precio.toLocaleString('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2})}</div>}
                      </div>
                    </div>
                  </div>
                  {dupArticleIds.has(detail.id) && (() => {
                    const otros = articles.filter(a => a.code===detail.code && a.id!==detail.id)
                    const thisTalles = new Set(detail.sizes.map(s => s.talle))
                    const lineas = otros.map(a => ({
                      ubic: a.ubic||'sin ubic.',
                      tallesDup: a.sizes.map(s => s.talle).filter(t => thisTalles.has(t)),
                      tallesSolo: a.sizes.map(s => s.talle).filter(t => !thisTalles.has(t)),
                    }))
                    return (
                      <div style={{margin:'0 24px 0',padding:'10px 14px',background:'#FFF8E1',borderBottom:'1px solid #FFE57A',fontSize:12.5,color:'#7a5800'}}>
                        <div style={{display:'flex',gap:8,alignItems:'center',marginBottom:6}}><span>⚠</span><b>Artículo registrado en múltiples ubicaciones</b></div>
                        {lineas.map((l,i) => (
                          <div key={i} style={{paddingLeft:22,marginBottom:3}}>
                            <b>{l.ubic}</b>
                            {l.tallesDup.length > 0 && <span style={{color:'#C2473D'}}> — talles duplicados: <b>{l.tallesDup.join(', ')}</b></span>}
                            {l.tallesSolo.length > 0 && <span style={{color:'#7a5800'}}> — talles exclusivos allí: {l.tallesSolo.join(', ')}</span>}
                          </div>
                        ))}
                      </div>
                    )
                  })()}
                  <div style={{padding:'18px 24px'}}>
                    <div style={{fontSize:12,color:'#8a8a82',fontWeight:700,letterSpacing:'.04em',marginBottom:14}}>STOCK POR TALLE</div>
                    {detail.sizes.map(s => (
                      <div key={s.talle} className="bar-row">
                        <div style={{width:46,fontWeight:700,fontSize:13.5}}>{s.talle}</div>
                        <div className="bar-track"><div className="bar-fill" style={{width:s.pct+'%',background:s.isLow||s.qty<=0?'#C2473D':'#FFD200'}} /></div>
                        <div style={{textAlign:'right',flexShrink:0}}>
                          <div className="mono" style={{fontWeight:600,fontSize:13.5}}>{s.qty}</div>
                          <div style={{fontSize:10.5,color:'#8a8a82'}}>mín {s.min}</div>
                        </div>
                      </div>
                    ))}
                    <div className="detail-actions">
                      <button className="btn btn-yellow" onClick={openReponer}>+ Registrar entrada</button>
                      <button className="btn btn-ghost" onClick={openEntregaFromDetail}>Registrar entrega</button>
                      <button className="btn btn-ghost" onClick={openDevolucionFromDetail}>↩ Devolución</button>
                      <button className="btn btn-dark" onClick={openAjuste}>Ajustar stock</button>
                      <button className="btn btn-ghost btn-full" onClick={openMover}>⇄ Cambiar de ubicación</button>
                      <button className="btn btn-ghost btn-full" onClick={openEdit}>✎ Editar artículo</button>
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
                <div key={d.id} className="table-row del-cols">
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
                  <div style={{display:'flex',justifyContent:'flex-end'}}><button className="btn-del" onClick={() => askDeleteDelivery(d.id)}>✕</button></div>
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
          {view === 'receptores' && (
            <div className="receptor-grid">
              {receptorCards.map(r => (
                <div key={r.name} className="card" style={{padding:20,display:'flex',gap:14,alignItems:'center'}}>
                  <div className="avatar xl">{r.ini}</div>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:700,fontSize:15}}>{r.name}</div>
                    <div style={{fontSize:12.5,color:'#8a8a82',marginTop:3}}>{r.count} entregas · {r.unidades} unidades</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* ===== MODALES ===== */}

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
                        {articles.filter(a => a.name.toLowerCase().includes(nd.cSearch.toLowerCase()) || a.code.toLowerCase().includes(nd.cSearch.toLowerCase())).length === 0
                          ? <div style={{padding:'10px 14px',fontSize:13,color:'#8a8a82'}}>Sin resultados</div>
                          : articles.filter(a => a.name.toLowerCase().includes(nd.cSearch.toLowerCase()) || a.code.toLowerCase().includes(nd.cSearch.toLowerCase())).map(a => (
                            <div key={a.code} style={{padding:'9px 14px',cursor:'pointer',borderBottom:'1px solid #F2F2EE',fontSize:13}} onClick={() => setNd(p=>({...p, cCode:a.code, cSearch:'', cTalle:'', cQty:''}))}>
                              <span style={{fontWeight:600}}>{a.name}</span>
                              <span style={{color:'#8a8a82',fontSize:11.5,marginLeft:8}}>{a.code}</span>
                            </div>
                          ))
                        }
                      </div>
                    )}
                  </div>
                  <div style={{display:'grid',gridTemplateColumns:'1fr 1fr auto',gap:10}}>
                    <select className="field-input" value={nd.cTalle} onChange={e => setNd(p=>({...p,cTalle:e.target.value}))}>
                      <option value="">Talle…</option>
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
              <div style={{display:'grid',gridTemplateColumns:'1fr 1.4fr',gap:12}}>
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
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
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
              <div style={{fontSize:13,color:'#8a8a82',marginBottom:16}}>{selA.name} <span className="mono">· {selA.code}</span></div>
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
              <div style={{fontSize:13,color:'#8a8a82',marginBottom:6}}>{selA.name} <span className="mono">· {selA.code}</span></div>
              <div style={{fontSize:12,color:'#9a7d00',background:'#FBF7E3',padding:'8px 12px',borderRadius:6,marginBottom:16}}>
                Corrección por recuento: ingresá la cantidad física real contada. El sistema registra la diferencia.
              </div>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
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
              <div style={{display:'grid',gridTemplateColumns:'1fr 1.4fr',gap:12}}>
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
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
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
                <label className="field-label">Seleccioná los talles a mover</label>
                <div className="talle-grid">
                  {selA.sizes.map(s => (
                    <button key={s.talle} className={`talle-btn${mv.tallesArr.includes(s.talle)?' active':''}`}
                      onClick={() => setMv(p => ({...p, tallesArr: p.tallesArr.includes(s.talle) ? p.tallesArr.filter(t=>t!==s.talle) : [...p.tallesArr, s.talle]}))}>
                      {s.talle}
                    </button>
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
