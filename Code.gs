/*****************************************************************
 * SALÓN DE BELLEZA · Backend (Google Apps Script + Google Sheets)
 * Versión 1.0
 *
 * 1) Crea una Hoja de cálculo nueva → Extensiones → Apps Script
 * 2) Pega este archivo como Code.gs, guarda y ejecuta instalar()
 * 3) Implementar → Nueva implementación → App web
 *    Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario
 * 4) Copia la URL /exec y pégala en index.html (API_URL_DEFAULT)
 *
 * Usuario inicial: admin · Contraseña: admin123 (cámbiala)
 *****************************************************************/

const TZ = 'America/Tegucigalpa';
const SESION_SEG = 21600; // 6 horas (máximo permitido por CacheService)

const SCHEMA = {
  Usuarios:          ['id','usuario','nombre','rol','color','hash','salt','activo','creado'],
  Clientes:          ['id','nombre','telefono','correo','cumpleanos','notas','creado'],
  Servicios:         ['id','nombre','categoria','duracion','precio','color','activo'],
  Tratamientos:      ['id','nombre','descripcion','sesiones','intervaloDias','duracion','precio','color','activo'],
  PlanesTratamiento: ['id','clienteId','tratamientoId','sesionesTotales','sesionesUsadas','precio','pagado','estado','fechaInicio','notas'],
  Productos:         ['id','codigo','nombre','categoria','precio','costo','stock','stockMin','activo'],
  Citas:             ['id','clienteId','estilistaId','tipo','refId','planId','titulo','inicio','fin','estado','precio','notas','creadoPor','creado','servicios'],
  Ventas:            ['id','fecha','clienteId','usuarioId','subtotal','descuento','isv','total','metodo','citaId','anulada','nota'],
  VentaDetalle:      ['id','ventaId','tipo','refId','descripcion','cantidad','precio','total'],
  Gastos:            ['id','fechaGasto','categoria','detalle','proveedor','monto','metodo','usuarioId','nota','anulado','creado'],
  Distribuciones:    ['id','mes','tipo','destino','detalle','monto','usuarioId','nota','anulado','creado'],
  Config:            ['clave','valor']
};
const PREFIJO = { Usuarios:'U', Clientes:'C', Servicios:'S', Tratamientos:'T', PlanesTratamiento:'P', Productos:'R', Citas:'A', Ventas:'V', VentaDetalle:'D', Gastos:'G', Distribuciones:'Q' };
const NUMS = new Set(['duracion','precio','sesiones','intervaloDias','sesionesTotales','sesionesUsadas','pagado','costo','stock','stockMin','subtotal','descuento','isv','total','cantidad','monto']);
const BOOLS = new Set(['activo','anulada','anulado']);
const SOLO_FECHA = new Set(['cumpleanos','fechaInicio','fechaGasto']);
const TEXTO = new Set(['telefono','codigo','inicio','fin','fecha','fechaInicio','cumpleanos','creado','valor','fechaGasto','mes']);
const ESTADOS = ['pendiente','confirmada','en_proceso','completada','cancelada','no_asistio'];
const NO_OCUPA = ['cancelada','no_asistio'];
const ROLES = ['admin','recepcion','estilista'];

const CONFIG_DEF = {
  nombre: 'Mi Salón de Belleza', telefono: '', direccion: '', rtn: '',
  moneda: 'L', isv: '0.15', preciosConISV: 'si',
  apertura: '08:00', cierre: '19:00', intervalo: '15',
  mensajeRecibo: '¡Gracias por tu visita!',
  mensajeWA: 'Hola {cliente} 👋, te recordamos tu cita en {salon} el {fecha} a las {hora} para {servicio}. ¡Te esperamos!'
};

const ESCRITURA = {
  Clientes: ['admin','recepcion'],
  PlanesTratamiento: ['admin','recepcion'],
  Servicios: ['admin'],
  Tratamientos: ['admin'],
  Productos: ['admin']
};

/* ===================== ENTRADA HTTP ===================== */

function doGet() {
  return json_({ ok: true, app: 'Salón de Belleza', version: '1.0' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    out = { ok: true, data: route_(req) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  const a = req.action, d = req.data || {};
  if (a === 'login') return login_(d);
  const u = auth_(req.token);
  switch (a) {
    case 'logout':          CacheService.getScriptCache().remove('S_' + req.token); return true;
    case 'bootstrap':       return bootstrap_(u);
    case 'citas':           return citas_(u, d);
    case 'guardarCita':     return lock_(() => guardarCita_(u, d));
    case 'estadoCita':      return lock_(() => estadoCita_(u, d));
    case 'guardar':         return lock_(() => guardar_(u, d));
    case 'eliminar':        return lock_(() => eliminar_(u, d));
    case 'guardarUsuario':  return lock_(() => guardarUsuario_(u, d));
    case 'cambiarClave':    return lock_(() => cambiarClave_(u, d));
    case 'venta':           return lock_(() => venta_(u, d));
    case 'ventas':          return ventas_(u, d);
    case 'anularVenta':     return lock_(() => anularVenta_(u, d));
    case 'dashboard':       return dashboard_(u, d);
    case 'historialCliente':return historialCliente_(u, d);
    case 'gastos':          return gastos_(u, d);
    case 'guardarGasto':    return lock_(() => guardarGasto_(u, d));
    case 'anularGasto':     return lock_(() => anularGasto_(u, d));
    case 'movimientos':     return movimientos_(u, d);
    case 'historial':       return historial_(u);
    case 'guardarDistribucion': return lock_(() => guardarDistribucion_(u, d));
    case 'anularDistribucion':  return lock_(() => anularDistribucion_(u, d));
    default: throw new Error('Acción no válida: ' + a);
  }
}

function lock_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ===================== HOJAS (BASE DE DATOS) ===================== */

const MEMO = {};

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sh_(name) {
  const ss = ss_();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    const h = SCHEMA[name];
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#8e3b63').setFontColor('#ffffff');
    s.setFrozenRows(1);
    h.forEach((k, i) => { if (TEXTO.has(k)) s.getRange(2, i + 1, s.getMaxRows() - 1, 1).setNumberFormat('@'); });
  }
  return s;
}

function obj_(h, r) {
  const o = {};
  h.forEach((k, i) => {
    let x = r[i];
    if (x instanceof Date) x = Utilities.formatDate(x, TZ, k === 'mes' ? 'yyyy-MM' : SOLO_FECHA.has(k) ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm");
    if (NUMS.has(k)) x = Number(x) || 0;
    else if (BOOLS.has(k)) x = (x === true || x === 'true' || x === 'TRUE' || x === 1 || x === 'SI');
    else x = (x === null || x === undefined) ? '' : String(x);
    o[k] = x;
  });
  return o;
}

function rows_(name) {
  if (!MEMO[name]) {
    const s = sh_(name), n = s.getLastRow(), h = SCHEMA[name];
    MEMO[name] = n < 2 ? [] : s.getRange(2, 1, n - 1, h.length).getValues()
      .filter(r => r[0] !== '' && r[0] !== null)
      .map(r => obj_(h, r));
  }
  return MEMO[name].map(o => Object.assign({}, o));
}

function cell_(v) { return (v === undefined || v === null) ? '' : v; }

function findRow_(s, id) {
  const n = s.getLastRow();
  if (n < 2 || !id) return 0;
  const f = s.getRange(2, 1, n - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}

let SEQ_ = 0;
function newId_(name) {
  SEQ_ = (SEQ_ + 1) % 1296;
  return (PREFIJO[name] || 'X') + Date.now().toString(36).toUpperCase() + SEQ_.toString(36).toUpperCase().padStart(2, '0') +
    Math.floor(Math.random() * 1296).toString(36).toUpperCase().padStart(2, '0');
}

function upsert_(name, rec) {
  const s = sh_(name), h = SCHEMA[name];
  const r = rec.id ? findRow_(s, rec.id) : 0;
  let full = rec;
  if (r) {
    const prev = obj_(h, s.getRange(r, 1, 1, h.length).getValues()[0]);
    full = Object.assign(prev, rec);
  } else if (!full.id) {
    full.id = newId_(name);
  }
  const row = h.map(k => cell_(full[k]));
  if (r) s.getRange(r, 1, 1, h.length).setValues([row]);
  else s.appendRow(row);
  delete MEMO[name];
  return obj_(h, row);
}

function appendRows_(name, objs) {
  if (!objs.length) return;
  const s = sh_(name), h = SCHEMA[name];
  const vals = objs.map(o => h.map(k => cell_(o[k])));
  s.getRange(s.getLastRow() + 1, 1, vals.length, h.length).setValues(vals);
  delete MEMO[name];
}

function deleteRow_(name, id) {
  const s = sh_(name), r = findRow_(s, id);
  if (!r) throw new Error('Registro no encontrado');
  s.deleteRow(r);
  delete MEMO[name];
}

function ahora_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm"); }
function hoy_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function r2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }

/* ===================== SEGURIDAD ===================== */

function hash_(clave, salt) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + clave, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(b);
}

function publico_(u) { const o = Object.assign({}, u); delete o.hash; delete o.salt; return o; }

function req_(u, roles) {
  if (roles.indexOf(u.rol) < 0) throw new Error('No tienes permiso para esta acción');
}

function login_(d) {
  const c = CacheService.getScriptCache();
  const usuario = String(d.usuario || '').trim().toLowerCase();
  const kf = 'F_' + usuario;
  const fallos = Number(c.get(kf) || 0);
  if (fallos >= 5) throw new Error('Demasiados intentos fallidos. Espera 10 minutos.');
  const u = rows_('Usuarios').find(x => String(x.usuario).toLowerCase() === usuario);
  if (!u || !u.activo || hash_(String(d.clave || ''), u.salt) !== u.hash) {
    c.put(kf, String(fallos + 1), 600);
    throw new Error('Usuario o contraseña incorrectos');
  }
  c.remove(kf);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  c.put('S_' + token, u.id, SESION_SEG);
  return { token: token, usuario: publico_(u) };
}

function auth_(token) {
  if (!token) throw new Error('SESION');
  const c = CacheService.getScriptCache();
  const uid = c.get('S_' + token);
  if (!uid) throw new Error('SESION');
  const u = rows_('Usuarios').find(x => x.id === uid);
  if (!u || !u.activo) throw new Error('SESION');
  c.put('S_' + token, uid, SESION_SEG); // sesión deslizante
  return u;
}

function cambiarClave_(u, d) {
  if (hash_(String(d.actual || ''), u.salt) !== u.hash) throw new Error('La contraseña actual no es correcta');
  if (String(d.nueva || '').length < 6) throw new Error('La nueva contraseña debe tener al menos 6 caracteres');
  const salt = Utilities.getUuid();
  upsert_('Usuarios', { id: u.id, salt: salt, hash: hash_(d.nueva, salt) });
  return true;
}

function guardarUsuario_(u, d) {
  req_(u, ['admin']);
  const r = Object.assign({}, d.registro || {});
  const todos = rows_('Usuarios');
  r.usuario = String(r.usuario || '').trim().toLowerCase();
  if (!r.usuario || !r.nombre) throw new Error('Usuario y nombre son obligatorios');
  if (/\s/.test(r.usuario)) throw new Error('El usuario no debe tener espacios');
  if (todos.some(x => x.usuario === r.usuario && x.id !== r.id)) throw new Error('Ese nombre de usuario ya existe');
  if (ROLES.indexOf(r.rol) < 0) throw new Error('Rol no válido');
  if (r.id === u.id && (r.rol !== 'admin' || r.activo === false)) throw new Error('No puedes quitarte el rol de administrador ni desactivarte');
  delete r.hash; delete r.salt;
  if (d.clave) {
    if (String(d.clave).length < 6) throw new Error('La contraseña debe tener al menos 6 caracteres');
    r.salt = Utilities.getUuid();
    r.hash = hash_(d.clave, r.salt);
  } else if (!r.id) {
    throw new Error('Asigna una contraseña al nuevo usuario');
  }
  if (!r.id) r.creado = ahora_();
  return publico_(upsert_('Usuarios', r));
}

/* ===================== DATOS GENERALES ===================== */

function config_() {
  const o = Object.assign({}, CONFIG_DEF);
  rows_('Config').forEach(r => { o[r.clave] = String(r.valor); });
  return o;
}

function guardarConfig_(reg) {
  const s = sh_('Config');
  const actual = Object.assign(config_(), reg || {});
  const vals = Object.keys(actual).map(k => [k, String(actual[k])]);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).clearContent();
  s.getRange(2, 1, vals.length, 2).setValues(vals);
  delete MEMO.Config;
  return actual;
}

function bootstrap_(u) {
  return {
    usuario: publico_(u),
    config: config_(),
    usuarios: rows_('Usuarios').map(publico_),
    clientes: rows_('Clientes'),
    servicios: rows_('Servicios'),
    tratamientos: rows_('Tratamientos'),
    productos: rows_('Productos'),
    planes: rows_('PlanesTratamiento'),
    gastoCats: ['admin','recepcion'].indexOf(u.rol) >= 0 ? catalogoGastos_() : []
  };
}

function guardar_(u, d) {
  const ent = d.entidad;
  if (ent === 'Config') { req_(u, ['admin']); return guardarConfig_(d.registro); }
  if (!ESCRITURA[ent]) throw new Error('Entidad no válida');
  req_(u, ESCRITURA[ent]);
  const r = Object.assign({}, d.registro || {});
  if (SCHEMA[ent].indexOf('nombre') >= 0 && !r.id && !String(r.nombre || '').trim()) throw new Error('El nombre es obligatorio');
  if (!r.id) {
    if (SCHEMA[ent].indexOf('creado') >= 0) r.creado = ahora_();
    if (SCHEMA[ent].indexOf('activo') >= 0 && r.activo === undefined) r.activo = true;
  }
  if (ent === 'Productos' && r.codigo) {
    if (rows_('Productos').some(p => p.codigo && String(p.codigo) === String(r.codigo) && p.id !== r.id)) throw new Error('Ya existe un producto con ese código');
  }
  if (ent === 'PlanesTratamiento' && !r.id) {
    const t = rows_('Tratamientos').find(x => x.id === r.tratamientoId);
    if (!t) throw new Error('Selecciona un tratamiento');
    if (!r.clienteId) throw new Error('Selecciona un cliente');
    r.sesionesTotales = Number(r.sesionesTotales) || t.sesiones || 1;
    r.sesionesUsadas = Number(r.sesionesUsadas) || 0;
    r.precio = (r.precio === undefined || r.precio === '') ? t.precio : Number(r.precio);
    r.pagado = Number(r.pagado) || 0;
    r.estado = r.estado || 'activo';
    r.fechaInicio = r.fechaInicio || hoy_();
  }
  return upsert_(ent, r);
}

function eliminar_(u, d) {
  req_(u, ['admin']);
  const permitidas = ['Clientes','Servicios','Tratamientos','Productos','PlanesTratamiento','Citas'];
  if (permitidas.indexOf(d.entidad) < 0) throw new Error('No se puede eliminar este tipo de registro');
  if (d.entidad === 'Citas') {
    const c = rows_('Citas').find(x => x.id === d.id);
    if (c && c.estado === 'completada') throw new Error('No se puede eliminar una cita completada');
  }
  deleteRow_(d.entidad, d.id);
  return true;
}

/* ===================== CITAS ===================== */

function nombreRef_(tipo, refId) {
  if (!refId) return '';
  const l = rows_(tipo === 'tratamiento' ? 'Tratamientos' : 'Servicios');
  const x = l.find(i => i.id === refId);
  return x ? x.nombre : '';
}

function citas_(u, d) {
  let l = rows_('Citas');
  if (d.desde) l = l.filter(c => c.fin >= d.desde);
  if (d.hasta) l = l.filter(c => c.inicio <= d.hasta);
  if (u.rol === 'estilista') l = l.filter(c => c.estilistaId === u.id);
  return l;
}

function guardarCita_(u, d) {
  const r = Object.assign({}, d.registro || {});
  const todas = rows_('Citas');
  const prev = r.id ? todas.find(c => c.id === r.id) : null;
  if (r.id && !prev) throw new Error('La cita ya no existe');
  if (u.rol === 'estilista') {
    if (prev && prev.estilistaId !== u.id) throw new Error('Solo puedes modificar tus propias citas');
    r.estilistaId = u.id;
  }
  if (!r.clienteId) throw new Error('Selecciona un cliente');
  if (!r.estilistaId) throw new Error('Selecciona un estilista');
  if (!r.inicio || !r.fin || r.fin <= r.inicio) throw new Error('El horario de la cita no es válido');
  if (prev && (prev.estado === 'completada')) throw new Error('No se puede modificar una cita completada');

  const choque = todas.find(c => c.id !== r.id && c.estilistaId === r.estilistaId &&
    NO_OCUPA.indexOf(c.estado) < 0 && r.inicio < c.fin && r.fin > c.inicio);
  if (choque && !d.forzar) {
    throw new Error('CHOQUE|' + choque.titulo + ' (' + choque.inicio.slice(11, 16) + '–' + choque.fin.slice(11, 16) + ')');
  }

  if (r.planId) {
    const p = rows_('PlanesTratamiento').find(x => x.id === r.planId);
    if (!p) throw new Error('El plan de tratamiento no existe');
    if (p.clienteId !== r.clienteId) throw new Error('El plan no pertenece a este cliente');
    if (p.estado !== 'activo' && (!prev || prev.planId !== r.planId)) throw new Error('El plan de tratamiento no está activo');
    r.tipo = 'tratamiento';
    r.refId = p.tratamientoId;
  }

  let nombreServicio = nombreRef_(r.tipo, r.refId);
  if (r.planId || r.tipo === 'tratamiento') {
    r.servicios = '';
  } else {
    // Una cita puede llevar varios servicios
    let lista = r.servicios;
    if (typeof lista === 'string') { try { lista = lista ? JSON.parse(lista) : []; } catch (e) { lista = []; } }
    if (!Array.isArray(lista) || !lista.length) lista = r.refId ? [{ id: r.refId }] : [];
    const svs = rows_('Servicios');
    lista = lista.map(z => {
      const id = typeof z === 'string' ? z : z.id, x = svs.find(v => v.id === id);
      if (!x) throw new Error('Uno de los servicios ya no existe');
      return { id: x.id, nombre: x.nombre, duracion: Number(z.duracion) || x.duracion,
        precio: z.precio === undefined || z.precio === '' ? x.precio : r2_(z.precio) };
    });
    if (!lista.length) throw new Error('Agrega al menos un servicio');
    r.tipo = 'servicio'; r.refId = lista[0].id; r.servicios = JSON.stringify(lista);
    nombreServicio = lista.map(z => z.nombre).join(' + ');
  }
  const cli = rows_('Clientes').find(x => x.id === r.clienteId);
  r.titulo = (cli ? cli.nombre : 'Cliente') + ' · ' + (nombreServicio || 'Cita');
  r.precio = Number(r.precio) || 0;
  r.estado = prev ? prev.estado : 'pendiente';
  if (!prev) { r.creadoPor = u.id; r.creado = ahora_(); }
  return upsert_('Citas', r);
}

function estadoCita_(u, d) {
  if (ESTADOS.indexOf(d.estado) < 0) throw new Error('Estado no válido');
  const c = rows_('Citas').find(x => x.id === d.id);
  if (!c) throw new Error('La cita no existe');
  if (u.rol === 'estilista' && c.estilistaId !== u.id) throw new Error('Solo puedes modificar tus propias citas');
  return aplicarEstado_(c, d.estado);
}

function aplicarEstado_(c, estado) {
  if (c.estado === estado) return c;
  if (c.planId) {
    const delta = (estado === 'completada' ? 1 : 0) - (c.estado === 'completada' ? 1 : 0);
    if (delta) {
      const p = rows_('PlanesTratamiento').find(x => x.id === c.planId);
      if (p) {
        p.sesionesUsadas = Math.max(0, (p.sesionesUsadas || 0) + delta);
        if (p.estado !== 'cancelado') p.estado = p.sesionesUsadas >= p.sesionesTotales ? 'completado' : 'activo';
        upsert_('PlanesTratamiento', p);
      }
    }
  }
  c.estado = estado;
  return upsert_('Citas', c);
}

/* ===================== VENTAS ===================== */

function venta_(u, d) {
  const items = (d.items || []).map(i => ({
    tipo: String(i.tipo || 'servicio'), refId: String(i.refId || ''), descripcion: String(i.descripcion || ''),
    cantidad: Math.max(1, Number(i.cantidad) || 1), precio: Math.max(0, Number(i.precio) || 0)
  }));
  if (!items.length) throw new Error('El ticket está vacío');
  if (items.some(i => i.tipo === 'tratamiento' || i.tipo === 'abono') && !d.clienteId) {
    throw new Error('Para vender tratamientos o abonar a un plan selecciona un cliente registrado');
  }
  const cfg = config_();
  const tasa = Number(cfg.isv) || 0;
  const incl = cfg.preciosConISV !== 'no';

  // Validar stock
  const prods = rows_('Productos'), pm = {};
  prods.forEach(p => { pm[p.id] = p; });
  const pedido = {};
  items.forEach(i => { if (i.tipo === 'producto') pedido[i.refId] = (pedido[i.refId] || 0) + i.cantidad; });
  Object.keys(pedido).forEach(id => {
    const p = pm[id];
    if (!p) throw new Error('Producto no encontrado');
    if (p.stock < pedido[id]) throw new Error('Stock insuficiente de ' + p.nombre + ' (disponible: ' + p.stock + ')');
  });

  const subtotal = items.reduce((a, i) => a + i.cantidad * i.precio, 0);
  const descuento = Math.min(Math.max(0, Number(d.descuento) || 0), subtotal);
  const base = subtotal - descuento;
  let isv, total;
  if (incl) { total = base; isv = tasa ? base - base / (1 + tasa) : 0; }
  else { isv = base * tasa; total = base + isv; }
  const factor = subtotal ? base / subtotal : 1;

  const venta = {
    id: newId_('Ventas'), fecha: ahora_(), clienteId: d.clienteId || '', usuarioId: u.id,
    subtotal: r2_(subtotal), descuento: r2_(descuento), isv: r2_(isv), total: r2_(total),
    metodo: d.metodo || 'Efectivo', citaId: d.citaId || '', anulada: false, nota: d.nota || ''
  };
  const detalle = items.map(i => ({
    id: newId_('VentaDetalle') + Math.floor(Math.random() * 99), ventaId: venta.id, tipo: i.tipo, refId: i.refId,
    descripcion: i.descripcion, cantidad: i.cantidad, precio: r2_(i.precio), total: r2_(i.cantidad * i.precio)
  }));
  appendRows_('Ventas', [venta]);
  appendRows_('VentaDetalle', detalle);

  // Descontar inventario
  Object.keys(pedido).forEach(id => { upsert_('Productos', { id: id, stock: pm[id].stock - pedido[id] }); });

  // Paquetes de tratamiento y abonos
  const trats = rows_('Tratamientos');
  items.forEach(i => {
    if (i.tipo === 'tratamiento') {
      const t = trats.find(x => x.id === i.refId);
      if (!t) return;
      for (let n = 0; n < i.cantidad; n++) {
        upsert_('PlanesTratamiento', {
          clienteId: d.clienteId, tratamientoId: t.id, sesionesTotales: t.sesiones || 1, sesionesUsadas: 0,
          precio: r2_(i.precio), pagado: r2_(i.precio * factor), estado: 'activo', fechaInicio: hoy_(), notas: 'Venta ' + venta.id
        });
      }
    } else if (i.tipo === 'abono') {
      const p = rows_('PlanesTratamiento').find(x => x.id === i.refId);
      if (p) upsert_('PlanesTratamiento', { id: p.id, pagado: r2_(p.pagado + i.precio * i.cantidad * factor) });
    }
  });

  // Cerrar la cita asociada
  if (d.citaId) {
    const c = rows_('Citas').find(x => x.id === d.citaId);
    if (c && c.estado !== 'completada' && NO_OCUPA.indexOf(c.estado) < 0) aplicarEstado_(c, 'completada');
  }
  return { venta: venta, detalle: detalle };
}

function rango_(d) {
  const desde = d.desde ? (d.desde.length === 10 ? d.desde + 'T00:00' : d.desde) : '';
  const hasta = d.hasta ? (d.hasta.length === 10 ? d.hasta + 'T23:59' : d.hasta) : '';
  return { desde: desde, hasta: hasta };
}

function ventas_(u, d) {
  const r = rango_(d);
  let v = rows_('Ventas').filter(x => (!r.desde || x.fecha >= r.desde) && (!r.hasta || x.fecha <= r.hasta));
  if (u.rol === 'estilista') v = v.filter(x => x.usuarioId === u.id);
  const ids = {};
  v.forEach(x => { x.detalle = []; ids[x.id] = x; });
  rows_('VentaDetalle').forEach(det => { if (ids[det.ventaId]) ids[det.ventaId].detalle.push(det); });
  return v.sort((a, b) => a.fecha < b.fecha ? 1 : -1);
}

function anularVenta_(u, d) {
  req_(u, ['admin']);
  const v = rows_('Ventas').find(x => x.id === d.id);
  if (!v) throw new Error('La venta no existe');
  if (v.anulada) throw new Error('La venta ya está anulada');
  const det = rows_('VentaDetalle').filter(x => x.ventaId === v.id);
  const prods = rows_('Productos');
  det.forEach(i => {
    if (i.tipo === 'producto') {
      const p = prods.find(x => x.id === i.refId);
      if (p) { p.stock += i.cantidad; upsert_('Productos', { id: p.id, stock: p.stock }); }
    } else if (i.tipo === 'abono') {
      const p = rows_('PlanesTratamiento').find(x => x.id === i.refId);
      if (p) upsert_('PlanesTratamiento', { id: p.id, pagado: Math.max(0, r2_(p.pagado - i.total)) });
    }
  });
  rows_('PlanesTratamiento').filter(p => p.notas === 'Venta ' + v.id && p.sesionesUsadas === 0)
    .forEach(p => upsert_('PlanesTratamiento', { id: p.id, estado: 'cancelado', notas: 'Anulado (venta ' + v.id + ')' }));
  upsert_('Ventas', { id: v.id, anulada: true, nota: (v.nota ? v.nota + ' · ' : '') + 'Anulada por ' + u.nombre + ' ' + ahora_() });
  return true;
}

/* ===================== GASTOS ===================== */

// Categorías, detalles y proveedores ya usados, ordenados por frecuencia (para sugerirlos)
function catalogoGastos_() {
  const m = {};
  rows_('Gastos').filter(g => !g.anulado).forEach(g => {
    const c = String(g.categoria || '').trim();
    if (!c) return;
    const x = m[c] || (m[c] = { usos: 0, ultimo: '', det: {}, prov: {}, montos: {}, fdet: {} });
    x.usos++;
    if (g.fechaGasto > x.ultimo) x.ultimo = g.fechaGasto;
    const d = String(g.detalle || '').trim();
    if (d) {
      x.det[d] = (x.det[d] || 0) + 1;
      if (!x.fdet[d] || g.fechaGasto >= x.fdet[d]) { x.fdet[d] = g.fechaGasto; x.montos[d] = g.monto; }
    }
    const p = String(g.proveedor || '').trim();
    if (p) x.prov[p] = (x.prov[p] || 0) + 1;
  });
  const orden = o => Object.keys(o).sort((a, b) => o[b] - o[a] || a.localeCompare(b));
  return Object.keys(m).map(k => ({ categoria: k, usos: m[k].usos, ultimo: m[k].ultimo,
    detalles: orden(m[k].det), proveedores: orden(m[k].prov), montos: m[k].montos }))
    .sort((a, b) => b.usos - a.usos || a.categoria.localeCompare(b.categoria));
}

function gastos_(u, d) {
  req_(u, ['admin','recepcion']);
  const desde = String(d.desde || '').slice(0, 10), hasta = String(d.hasta || '').slice(0, 10);
  return rows_('Gastos').filter(g => (!desde || g.fechaGasto >= desde) && (!hasta || g.fechaGasto <= hasta))
    .sort((a, b) => (a.fechaGasto + a.creado) < (b.fechaGasto + b.creado) ? 1 : -1);
}

function guardarGasto_(u, d) {
  req_(u, ['admin','recepcion']);
  const r = Object.assign({}, d.registro || {});
  const prev = r.id ? rows_('Gastos').find(g => g.id === r.id) : null;
  if (r.id && !prev) throw new Error('El gasto ya no existe');
  if (prev && prev.anulado) throw new Error('No se puede modificar un gasto anulado');
  if (prev && u.rol !== 'admin' && prev.usuarioId !== u.id) throw new Error('Solo puedes modificar los gastos que registraste');
  ['categoria','detalle','proveedor','nota','metodo'].forEach(k => { r[k] = String(r[k] || '').trim().replace(/\s+/g, ' '); });
  if (!r.categoria) throw new Error('Selecciona o escribe una categoría');
  if (!r.detalle) throw new Error('Escribe el detalle del gasto');
  r.monto = r2_(r.monto);
  if (!(r.monto > 0)) throw new Error('El monto debe ser mayor a cero');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(r.fechaGasto || ''))) r.fechaGasto = hoy_();
  if (r.fechaGasto > hoy_()) throw new Error('La fecha del gasto no puede ser futura');
  // Unificar con lo ya usado (evita "Agua" y "agua" como cosas distintas)
  const cats = catalogoGastos_();
  const cat = cats.find(c => c.categoria.toLowerCase() === r.categoria.toLowerCase());
  if (cat) {
    r.categoria = cat.categoria;
    const det = cat.detalles.find(x => x.toLowerCase() === r.detalle.toLowerCase());
    if (det) r.detalle = det;
  }
  const prov = [].concat.apply([], cats.map(c => c.proveedores)).find(x => r.proveedor && x.toLowerCase() === r.proveedor.toLowerCase());
  if (prov) r.proveedor = prov;
  r.metodo = r.metodo || 'Efectivo';
  if (!prev) { r.usuarioId = u.id; r.creado = ahora_(); r.anulado = false; }
  else { delete r.usuarioId; delete r.creado; delete r.anulado; }
  const g = upsert_('Gastos', r);
  return { gasto: g, cats: catalogoGastos_() };
}

function anularGasto_(u, d) {
  req_(u, ['admin']);
  const g = rows_('Gastos').find(x => x.id === d.id);
  if (!g) throw new Error('El gasto no existe');
  if (g.anulado) throw new Error('El gasto ya está anulado');
  upsert_('Gastos', { id: g.id, anulado: true, nota: (g.nota ? g.nota + ' · ' : '') + 'Anulado por ' + u.nombre + ' ' + ahora_() });
  return { cats: catalogoGastos_() };
}

/* ===================== DIARIO, HISTORIAL Y UTILIDADES ===================== */

// Resumen por mes: ingresos (ventas), ISV, gastos, utilidad y cuánto de ella ya tiene destino
function resumenMeses_() {
  const m = {};
  const mes = k => m[k] || (m[k] = { mes: k, ingresos: 0, isv: 0, gastos: 0, ventas: 0, nGastos: 0, asignado: 0, cuentas: 0, pagos: 0 });
  rows_('Ventas').forEach(v => { if (v.anulada || !v.fecha) return; const x = mes(v.fecha.slice(0, 7)); x.ingresos += v.total; x.isv += v.isv; x.ventas++; });
  rows_('Gastos').forEach(g => { if (g.anulado || !g.fechaGasto) return; const x = mes(g.fechaGasto.slice(0, 7)); x.gastos += g.monto; x.nGastos++; });
  rows_('Distribuciones').forEach(d => { if (d.anulado) return; const x = mes(d.mes); x.asignado += d.monto; if (d.tipo === 'pago') x.pagos += d.monto; else x.cuentas += d.monto; });
  const actual = hoy_().slice(0, 7);
  return Object.keys(m).sort().reverse().map(k => {
    const x = m[k];
    ['ingresos','isv','gastos','asignado','cuentas','pagos'].forEach(f => { x[f] = r2_(x[f]); });
    x.utilidad = r2_(x.ingresos - x.isv - x.gastos);
    x.pendiente = r2_(Math.max(0, x.utilidad - x.asignado));
    x.enCurso = k >= actual;
    return x;
  });
}

function historial_(u) {
  req_(u, ['admin','recepcion']);
  const dist = rows_('Distribuciones');
  const activos = dist.filter(d => !d.anulado);
  const uso = t => { const n = {}; activos.filter(d => d.tipo === t).forEach(d => { n[d.destino] = (n[d.destino] || 0) + 1; }); return Object.keys(n).sort((a, b) => n[b] - n[a]); };
  const total = t => { const n = {}; activos.filter(d => d.tipo === t).forEach(d => { n[d.destino] = r2_((n[d.destino] || 0) + d.monto); }); return Object.keys(n).map(k => ({ destino: k, total: n[k] })).sort((a, b) => b.total - a.total); };
  return {
    meses: resumenMeses_(),
    distribuciones: dist.sort((a, b) => a.creado < b.creado ? 1 : -1),
    cuentas: total('cuenta'), pagos: total('pago'),
    destinos: { cuenta: uso('cuenta'), pago: uso('pago') }
  };
}

function movimientos_(u, d) {
  req_(u, ['admin','recepcion']);
  const mes = /^\d{4}-\d{2}$/.test(String(d.mes || '')) ? d.mes : hoy_().slice(0, 7);
  const usuarios = {}, clientes = {};
  rows_('Usuarios').forEach(x => { usuarios[x.id] = x.nombre; });
  rows_('Clientes').forEach(x => { clientes[x.id] = x.nombre; });
  const ventas = rows_('Ventas').filter(v => String(v.fecha).slice(0, 7) === mes);
  const det = {};
  ventas.forEach(v => { det[v.id] = []; });
  rows_('VentaDetalle').forEach(x => { if (det[x.ventaId]) det[x.ventaId].push(x.descripcion + (x.cantidad > 1 ? ' ×' + x.cantidad : '')); });
  const mov = ventas.map(v => ({
    id: v.id, fecha: v.fecha, tipo: 'ingreso', concepto: det[v.id].join(', ') || 'Venta', categoria: 'Venta',
    tercero: v.clienteId ? (clientes[v.clienteId] || 'Cliente') : 'Consumidor final', metodo: v.metodo,
    usuario: usuarios[v.usuarioId] || '', monto: v.total, isv: v.isv, anulado: v.anulada, nota: v.nota
  }));
  rows_('Gastos').filter(g => String(g.fechaGasto).slice(0, 7) === mes).forEach(g => mov.push({
    id: g.id, fecha: g.fechaGasto + 'T' + (String(g.creado).slice(0, 10) === g.fechaGasto ? String(g.creado).slice(11, 16) : '00:00'),
    tipo: 'gasto', concepto: g.detalle, categoria: g.categoria, tercero: g.proveedor, metodo: g.metodo,
    usuario: usuarios[g.usuarioId] || '', monto: g.monto, isv: 0, anulado: g.anulado, nota: g.nota
  }));
  mov.sort((a, b) => a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0);
  const resumen = resumenMeses_().find(x => x.mes === mes) ||
    { mes: mes, ingresos: 0, isv: 0, gastos: 0, ventas: 0, nGastos: 0, asignado: 0, cuentas: 0, pagos: 0, utilidad: 0, pendiente: 0, enCurso: mes >= hoy_().slice(0, 7) };
  return { mes: mes, movimientos: mov, resumen: resumen, distribuciones: rows_('Distribuciones').filter(x => x.mes === mes && !x.anulado) };
}

function guardarDistribucion_(u, d) {
  req_(u, ['admin']);
  const r = Object.assign({}, d.registro || {});
  delete r.id;
  r.mes = String(r.mes || '');
  if (!/^\d{4}-\d{2}$/.test(r.mes)) throw new Error('Mes no válido');
  if (r.mes >= hoy_().slice(0, 7)) throw new Error('Solo se puede decidir la utilidad de meses ya cerrados');
  r.tipo = r.tipo === 'pago' ? 'pago' : 'cuenta';
  ['destino','detalle','nota'].forEach(k => { r[k] = String(r[k] || '').trim().replace(/\s+/g, ' '); });
  if (!r.destino) throw new Error(r.tipo === 'pago' ? 'Indica qué se va a pagar' : 'Indica a qué cuenta va el dinero');
  r.monto = r2_(r.monto);
  if (!(r.monto > 0)) throw new Error('El monto debe ser mayor a cero');
  const m = resumenMeses_().find(x => x.mes === r.mes);
  if (!m || m.utilidad <= 0) throw new Error('Ese mes no dejó utilidad para asignar');
  if (r.monto > m.pendiente + 0.001) throw new Error('El monto supera la utilidad disponible de ese mes (' + m.pendiente.toFixed(2) + ')');
  const igual = rows_('Distribuciones').find(x => !x.anulado && x.tipo === r.tipo && x.destino.toLowerCase() === r.destino.toLowerCase());
  if (igual) r.destino = igual.destino;
  r.usuarioId = u.id; r.creado = ahora_(); r.anulado = false;
  upsert_('Distribuciones', r);
  return historial_(u);
}

function anularDistribucion_(u, d) {
  req_(u, ['admin']);
  const x = rows_('Distribuciones').find(z => z.id === d.id);
  if (!x) throw new Error('El registro no existe');
  if (x.anulado) throw new Error('Ya estaba anulado');
  upsert_('Distribuciones', { id: x.id, anulado: true, nota: (x.nota ? x.nota + ' · ' : '') + 'Anulado por ' + u.nombre + ' ' + ahora_() });
  return historial_(u);
}

/* ===================== REPORTES ===================== */

function dashboard_(u, d) {
  req_(u, ['admin','recepcion']);
  const r = rango_({ desde: d.desde || hoy_(), hasta: d.hasta || hoy_() });
  const ventas = rows_('Ventas').filter(v => !v.anulada && v.fecha >= r.desde && v.fecha <= r.hasta);
  const ids = new Set(ventas.map(v => v.id));
  const det = rows_('VentaDetalle').filter(x => ids.has(x.ventaId));
  const todas = rows_('Citas');
  const citaMap = {};
  todas.forEach(c => { citaMap[c.id] = c; });
  const citas = todas.filter(c => c.inicio >= r.desde && c.inicio <= r.hasta);
  const usuarios = rows_('Usuarios');

  const k = { ingresos: 0, ventas: ventas.length, ticket: 0, isv: 0, descuentos: 0, citas: citas.length,
    completadas: 0, canceladas: 0, noAsistio: 0, pendientes: 0, servicios: 0, productos: 0, tratamientos: 0, abonos: 0, clientesNuevos: 0,
    gastos: 0, utilidad: 0 };
  const porDia = {}, porMetodo = {}, porEst = {}, top = {}, porEstado = {};
  const est = id => {
    if (!porEst[id]) {
      const x = usuarios.find(z => z.id === id);
      porEst[id] = { id: id, nombre: x ? x.nombre : '—', color: x ? x.color : '', citas: 0, completadas: 0, ingresos: 0 };
    }
    return porEst[id];
  };

  ventas.forEach(v => {
    k.ingresos += v.total; k.isv += v.isv; k.descuentos += v.descuento;
    const dia = v.fecha.slice(0, 10);
    porDia[dia] = (porDia[dia] || 0) + v.total;
    porMetodo[v.metodo] = (porMetodo[v.metodo] || 0) + v.total;
    const c = v.citaId && citaMap[v.citaId];
    est(c ? c.estilistaId : v.usuarioId).ingresos += v.total;
  });
  det.forEach(i => {
    const campo = { servicio: 'servicios', producto: 'productos', tratamiento: 'tratamientos', abono: 'abonos' }[i.tipo] || 'servicios';
    k[campo] += i.total;
    const key = i.tipo === 'abono' ? 'abono' : i.tipo + ':' + (i.refId || i.descripcion);
    if (!top[key]) top[key] = { tipo: i.tipo, descripcion: i.tipo === 'abono' ? 'Abonos a planes' : i.descripcion, cantidad: 0, total: 0 };
    top[key].cantidad += i.cantidad; top[key].total += i.total;
  });
  citas.forEach(c => {
    porEstado[c.estado] = (porEstado[c.estado] || 0) + 1;
    const e = est(c.estilistaId);
    e.citas++;
    if (c.estado === 'completada') { e.completadas++; k.completadas++; }
    if (c.estado === 'cancelada') k.canceladas++;
    if (c.estado === 'no_asistio') k.noAsistio++;
    if (c.estado === 'pendiente') k.pendientes++;
  });
  k.clientesNuevos = rows_('Clientes').filter(c => c.creado >= r.desde && c.creado <= r.hasta).length;
  k.ticket = ventas.length ? k.ingresos / ventas.length : 0;
  const gastosPorDia = {}, gastosPorCategoria = {};
  rows_('Gastos').filter(g => !g.anulado && g.fechaGasto >= r.desde.slice(0, 10) && g.fechaGasto <= r.hasta.slice(0, 10)).forEach(g => {
    k.gastos += g.monto;
    gastosPorDia[g.fechaGasto] = (gastosPorDia[g.fechaGasto] || 0) + g.monto;
    gastosPorCategoria[g.categoria] = (gastosPorCategoria[g.categoria] || 0) + g.monto;
  });
  k.utilidad = k.ingresos - k.isv - k.gastos; // ventas sin ISV menos gastos
  Object.keys(k).forEach(x => { k[x] = r2_(k[x]); });

  return {
    kpi: k, porDia: porDia, porMetodo: porMetodo, porEstado: porEstado, gastosPorDia: gastosPorDia,
    gastosPorCategoria: Object.keys(gastosPorCategoria).map(c => ({ categoria: c, total: r2_(gastosPorCategoria[c]) })).sort((a, b) => b.total - a.total),
    top: Object.keys(top).map(x => top[x]).sort((a, b) => b.total - a.total).slice(0, 10),
    estilistas: Object.keys(porEst).map(x => porEst[x]).sort((a, b) => b.ingresos - a.ingresos)
  };
}

function historialCliente_(u, d) {
  req_(u, ['admin','recepcion']);
  const citas = rows_('Citas').filter(c => c.clienteId === d.id).sort((a, b) => a.inicio < b.inicio ? 1 : -1);
  const ventas = rows_('Ventas').filter(v => v.clienteId === d.id).sort((a, b) => a.fecha < b.fecha ? 1 : -1);
  const ids = {};
  ventas.forEach(v => { v.detalle = []; ids[v.id] = v; });
  rows_('VentaDetalle').forEach(x => { if (ids[x.ventaId]) ids[x.ventaId].detalle.push(x); });
  return { citas: citas, ventas: ventas, planes: rows_('PlanesTratamiento').filter(p => p.clienteId === d.id) };
}

/* ===================== INSTALACIÓN ===================== */

function instalar() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(SCHEMA).forEach(n => {
    // Crea las hojas que falten y actualiza encabezados si la versión nueva agregó columnas
    const sh = sh_(n), h = SCHEMA[n];
    const actual = sh.getRange(1, 1, 1, h.length).getValues()[0];
    if (h.some((k, i) => actual[i] !== k)) sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#8e3b63').setFontColor('#ffffff');
  });
  if (rows_('Config').length === 0) guardarConfig_({});
  if (rows_('Usuarios').length === 0) {
    const salt = Utilities.getUuid();
    upsert_('Usuarios', { usuario: 'admin', nombre: 'Administrador', rol: 'admin', color: '#8e3b63',
      salt: salt, hash: hash_('admin123', salt), activo: true, creado: ahora_() });
  }
  if (rows_('Servicios').length === 0) {
    appendRows_('Servicios', [
      ['Corte de dama','Cabello',45,300,'#c9798f'], ['Corte de caballero','Cabello',30,150,'#7a9cc6'],
      ['Tinte completo','Color',120,1200,'#a35d8a'], ['Mechas / Balayage','Color',180,2500,'#d4a24c'],
      ['Peinado','Cabello',60,500,'#e3a18b'], ['Manicure','Uñas',45,250,'#e07a9b'],
      ['Pedicure','Uñas',60,350,'#5fb3a1'], ['Maquillaje social','Maquillaje',60,800,'#b07cc6']
    ].map(x => ({ id: newId_('Servicios') + x[0].length, nombre: x[0], categoria: x[1], duracion: x[2], precio: x[3], color: x[4], activo: true })));
  }
  if (rows_('Tratamientos').length === 0) {
    appendRows_('Tratamientos', [
      ['Hidratación capilar profunda','Paquete de 4 sesiones semanales',4,7,60,1600,'#4f8fc0'],
      ['Limpieza facial profunda','Paquete de 3 sesiones cada 15 días',3,15,60,1800,'#5fae7d'],
      ['Alisado con keratina','Aplicación + 1 retoque',2,30,180,3500,'#c08a4f']
    ].map(x => ({ id: newId_('Tratamientos') + x[0].length, nombre: x[0], descripcion: x[1], sesiones: x[2], intervaloDias: x[3], duracion: x[4], precio: x[5], color: x[6], activo: true })));
  }
  if (rows_('Productos').length === 0) {
    appendRows_('Productos', [
      ['P001','Shampoo hidratante 500 ml','Cuidado capilar',380,220,12,3],
      ['P002','Acondicionador reparador 500 ml','Cuidado capilar',390,225,10,3],
      ['P003','Mascarilla capilar 250 g','Cuidado capilar',450,260,6,2],
      ['P004','Esmalte de uñas','Uñas',120,60,30,5]
    ].map(x => ({ id: newId_('Productos') + x[0], codigo: x[0], nombre: x[1], categoria: x[2], precio: x[3], costo: x[4], stock: x[5], stockMin: x[6], activo: true })));
  }
  ['Hoja 1','Sheet1','Hoja1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  Logger.log('Instalación lista. Usuario: admin · Contraseña: admin123');
}
