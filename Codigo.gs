// ===================== COMEDOR SODEXO – Desayuno, almuerzo y cena (v5 · local y recojo) =====================
// • Hoja "Menú": se programa el menú de toda la semana (Fecha · Servicio · Plato · Stock).
// • Panel (J:L): horario de cada servicio. Cada uno abre y cierra solo.
// • Hoja "Precios": opciones por servicio y modalidad (Local / Recojo).
// • Hoja "Platos": ingredientes, kcal, proteínas y grasas de cada plato (se llena una vez por plato).
// • HTML "index": consulta por fecha; pedidos de hoy dentro del horario de cada servicio.
// • Un pedido por persona por servicio por día. Para anular: "Anulado" en la columna Estado de "Pedidos".

const HOJA_MENU = 'Menú';
const HOJA_PRECIOS = 'Precios';
const HOJA_PEDIDOS = 'Pedidos';
const HOJA_PLATOS = 'Platos';                   // Ingredientes y valores nutricionales de cada plato
const MAX_POR_PERSONA = 10;          // máximo por opción en un pedido
const FILAS_MENU = 300;              // filas disponibles para programar el menú
const AZUL = '#1f3864', AZUL_CLARO = '#dde5f3', GRIS = '#f3f3f3', AMARILLO = '#fffbea';
const PERSONAL_DEFECTO = ['SPCC', 'Funcionario', 'Empleado', 'Contratista'];
const COMEDOR = 'Staff Ilo';                    // Casilla marcada en el vale: 'Staff Ilo' u 'Hospital'
const CARPETA_VALES = 'Vales de consumo';       // Carpeta (junto a la hoja) donde se guarda cada vale en PDF
const SERVICIOS_DEF = [['Desayuno', 5, 7], ['Almuerzo', 9, 11], ['Cena', 16, 18]];
// Servicio, opción, precio con IGV, modalidad, activo. Vacío = pendiente de confirmar.
const PRECIOS_DEF = [
  ['Desayuno', 'Completo', 3.9, 'Local', 'Sí'],
  ['Desayuno', 'Completo', '', 'Recojo', 'No'],
  ['Almuerzo', 'Completo', 10.5, 'Local', 'Sí'],
  ['Almuerzo', 'Simple', 8.6, 'Local', 'Sí'],
  ['Almuerzo', 'Solo segundo', 7.4, 'Local', 'Sí'],
  ['Almuerzo', 'Completo', 13, 'Recojo', 'Sí'],
  ['Almuerzo', 'Económico', 10.8, 'Recojo', 'Sí'],
  ['Almuerzo', 'Solo segundo', 9, 'Recojo', 'Sí'],
  ['Cena', 'Completo', 9.3, 'Local', 'Sí'],
  ['Cena', 'Simple', 8, 'Local', 'Sí'],
  ['Cena', 'Completo', 11.9, 'Recojo', 'Sí'],
  ['Cena', 'Económico', 10.3, 'Recojo', 'Sí'],
  ['Cena', 'Solo segundo', '', 'Local', 'No'],
  ['Cena', 'Solo segundo', '', 'Recojo', 'No'],
  ['Rancho caliente', 'Completo', 13.2, 'Local', 'Sí']
];
const CAB_PED = ['Fecha y hora', 'Código', 'Fecha consumo', 'Servicio', 'Nombre completo', 'Tipo de personal',
  'Registro / DNI', 'Plato', 'Opción', 'Cantidad', 'Precio unit.', 'Subtotal', 'Descuento planilla', 'Observaciones', 'Estado', 'Modalidad', 'Solicitud ID',
  'Correo', 'Empresa', 'Dpto. / Área', 'Vale PDF', 'Método de pago'];
const METODOS_PAGO = ['Efectivo', 'Yape', 'Plin', 'Tarjeta'];   // Se pide cuando NO es descuento por planilla
const COL_CORREO = 18, COL_VALE = 21;           // columnas "Correo" y "Vale PDF" en Pedidos
const MAX_COPIAS = 5;                           // copias por correo que se pueden pedir de un mismo vale

// Panel de control en la hoja Menú (columnas J:L)
const P = {
  LINK: 'K2', SRV_INI: 4, SRV_FIN: 8, VENTA: 'K9', ESTADO: 'K10', CARPETA: 'K12'
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🍽️ Comedor')
    .addItem('📂 Abrir carpeta de vales (PDF)', 'abrirCarpetaVales_')
    .addItem('🔗 Ver link de pedidos', 'verLink_')
    .addItem('🥗 Agregar platos nuevos a "Platos"', 'agregarPlatos_')
    .addItem('🧪 Cargar menú de ejemplo (21–27/09)', 'cargarMenuEjemplo')
    .addSeparator()
    .addItem('Recalcular pedidos ahora', 'actualizarStock_')
    .addItem('Reparar diseño y activadores', 'configurarTodo_')
    .addToUi();
}

// ---------------------------------------------------------------- Utilidades
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('COMEDOR_SS_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
  if (!ss) throw new Error('Primero ejecuta configurarTodo desde Extensiones → Apps Script de la hoja de cálculo.');
  return ss;
}
let TZ = null;
function tz_() { return TZ || (TZ = ss_().getSpreadsheetTimeZone()); }
function aviso_(msg) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); } }
function props_() { return PropertiesService.getDocumentProperties(); }
function hhmm_(min) { return Math.floor(min / 60) + ':' + ('0' + min % 60).slice(-2); }
function pad_(n) { return ('0' + n).slice(-2); }
function norm_(s) { return String(s || '').replace(/\s+/g, '').toUpperCase(); }
function keyS_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase(); }
function hoyKey_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function hoyTxt_() { return Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy'); }

function limpiar_(v, max) {
  let s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// Convierte fecha (Date o texto dd/mm/aaaa) a "aaaa-mm-dd"
function fechaKey_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v,tz_(),'yyyy-MM-dd');
  const s=String(v || '').trim();
  let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/), y, mo, d;
  if(m){y=Number(m[1]);mo=Number(m[2]);d=Number(m[3]);}
  else {m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2}|\d{4})$/);if(!m)return '';y=Number(m[3])+(m[3].length===2?2000:0);mo=Number(m[2]);d=Number(m[1]);}
  const check=new Date(Date.UTC(y,mo-1,d));
  return check.getUTCFullYear()===y && check.getUTCMonth()===mo-1 && check.getUTCDate()===d ? y+'-'+pad_(mo)+'-'+pad_(d) : '';
}

function horaEnMinutos_(v, porDefecto) {
  if (v === '' || v == null) return porDefecto;
  if (v instanceof Date) v = Utilities.formatDate(v, tz_(), 'HH:mm');
  if (typeof v === 'number') {
    const n = Math.round(v < 1 && v > 0 ? v * 1440 : v * 60);
    return Number.isFinite(n) && n >= 0 && n <= 1440 ? n : porDefecto;
  }
  const m = String(v).trim().match(/^(\d{1,2})(?::(\d{2}))?$/);
  if (!m || Number(m[2] || 0) > 59) return porDefecto;
  const n = Number(m[1]) * 60 + Number(m[2] || 0);
  return n <= 1440 ? n : porDefecto;
}

function minutosAhora_() {
  const s = Utilities.formatDate(new Date(), tz_(), 'HH:mm').split(':');
  return Number(s[0]) * 60 + Number(s[1]);
}

function fechaLarga_(d) {
  const dias = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo'];
  const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const dia = dias[Number(Utilities.formatDate(d, tz_(), 'u')) - 1];
  return dia.charAt(0).toUpperCase() + dia.slice(1) + ' ' + Utilities.formatDate(d, tz_(), 'd') +
    ' de ' + meses[Number(Utilities.formatDate(d, tz_(), 'M')) - 1];
}

// ---------------------------------------------------------------- Formulario web
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Pedido de Comedor – SODEXO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Lo llama la página: servicios de hoy con su estado, platos y opciones
function datosFormulario(fecha) {
  const actual = hoyKey_(), hoy = fecha ? fechaKey_(fecha) : actual, ahora = minutosAhora_();
  if (!hoy) throw new Error('Fecha no válida.');
  const opciones = leerOpciones_(), menu = leerMenuTodo_(), servs = leerServicios_(), fichas = leerPlatos_();
  const menuHoy = unicos_(menu.filter(m => m.fecha === hoy));
  const usados = usados_(leerLineas_(hoy, false));

  const servicios = servs.map(s => {
    const platos = menuHoy.filter(m => m.sKey === s.key).map(m => {
      const u = usados[claveUso_(hoy, s.key, m.plato)] || 0;
      const f = fichas[keyS_(m.plato)] || {};
      return { plato: m.plato, stock: m.stock, quedan: m.stock == null ? null : Math.max(0, m.stock - u),
        ingredientes: f.ingredientes || '', kcal: f.kcal, proteinas: f.proteinas, grasas: f.grasas };
    });
    const ops = opciones[s.key] || [];
    const estado = ahora < s.abre ? 'pronto' : (ahora < s.cierra ? 'abierto' : 'cerrado');
    let motivo = '';
    if (!platos.length) motivo = 'Sin menú programado';
    else if (!ops.length) motivo = 'Precios por confirmar';
    else if (!platos.some(p => p.quedan == null || p.quedan > 0)) motivo = 'Agotado';
    return {
      nombre: s.nombre, key: s.key, abre: hhmm_(s.abre), cierra: hhmm_(s.cierra),
      estado: estado, motivo: motivo, disponible: hoy === actual && estado === 'abierto' && !motivo,
      opciones: ops, platos: platos
    };
  });

  // Menú programado de los próximos días (solo consulta)
  const semana = Array.from(new Set(menu.map(m => m.fecha).filter(f => f >= actual))).sort().slice(0, 7).map(f => {
    const delDia = unicos_(menu.filter(m => m.fecha === f));
    return { fecha: f, servicios: servs.map(s => ({ nombre: s.nombre,
      platos: delDia.filter(m => m.sKey === s.key).map(m => m.plato) })).filter(s => s.platos.length) };
  }).filter(d => d.servicios.length);

  return { fecha: hoy.split('-').reverse().join('/'), fechaKey: hoy, hoy: actual, ahora: ahora,
    semana: semana, servicios: servicios, personal: leerPersonal_(), pagos: METODOS_PAGO, max: MAX_POR_PERSONA };
}

// Recibe el pedido, valida todo con bloqueo y lo guarda
// Registra el pedido y guarda su vale en PDF en Drive (copia del administrador).
function enviarPedido(p) {
  p = p || {};
  const firma = String(p.firma || '');
  if (!/^data:image\/png;base64,[A-Za-z0-9+\/=]+$/.test(firma) || firma.length > 300000) throw new Error('Firma tu vale antes de enviar.');
  const v = registrarPedido_(p);
  v.firma = firma;
  try { guardarVale_(v); } catch (e) { Logger.log('Vale ' + v.codigo + ': ' + e); }
  return { codigo: v.codigo, servicio: v.servicio, modalidad: v.modalidad, hora: v.hora, nombre: v.nombre,
    registro: v.registro, items: v.items, total: v.total };
}

// Lo llama la pantalla final: envía una copia del vale al correo que escriba el trabajador.
function enviarCopiaVale(p) {
  p = p || {};
  const solicitud = String(p.solicitud || ''), correo = limpiar_(p.correo, 120).toLowerCase();
  if (!/^[a-zA-Z0-9-]{16,80}$/.test(solicitud)) throw new Error('No se encontró tu pedido. Pide tu vale en el comedor.');
  if (!correoValido_(correo)) throw new Error('Escribe un correo válido.');
  let lineas = leerLineas_(hoyKey_(), false).filter(l => l.solicitud === solicitud);
  if (!lineas.length) lineas = leerLineas_(null, false).filter(l => l.solicitud === solicitud);
  if (!lineas.length) throw new Error('No se encontró tu pedido. Pide tu vale en el comedor.');
  const cache = CacheService.getScriptCache(), clave = 'copias-' + solicitud, n = Number(cache.get(clave) || 0);
  if (n >= MAX_COPIAS) throw new Error('Ya se enviaron ' + MAX_COPIAS + ' copias de este vale.');
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Hoy ya no se pueden enviar más correos. Pide tu vale en el comedor.');

  const v = comprobante_(lineas);
  const id = (v.vale.match(/\/d\/([-\w]+)/) || [])[1];
  let pdf = null;
  if (id) try { pdf = DriveApp.getFileById(id).getBlob(); } catch (e) { Logger.log(e); }
  if (!pdf) pdf = guardarVale_(v) || valePdf_(v);
  correoVale_(v, correo, pdf);
  cache.put(clave, String(n + 1), 21600);

  const lista = v.correo.split(/,\s*/).filter(String);
  if (lista.indexOf(correo) < 0) lista.push(correo);
  const sh = hojaPedidos_();
  v.filas.forEach(f => sh.getRange(f, COL_CORREO).setValue(lista.join(', ')));
  return { correo: correo };
}

function correoValido_(c) { return /^[^\s@']+@[^\s@]+\.[^\s@]{2,}$/.test(c); }

function registrarPedido_(p) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Hay muchos pedidos a la vez. Intenta de nuevo en unos segundos.');
  try {
    const nombre = limpiar_(p.nombre, 80);
    const registro = limpiar_(p.registro, 20).toUpperCase();
    const personal = String(p.personal || '');
    const descuento = p.descuento === 'No' ? 'No' : (p.descuento === 'Sí' ? 'Sí' : '');
    const obs = limpiar_(p.obs, 200);
    const empresa = limpiar_(p.empresa, 60), area = limpiar_(p.area, 60);

    if (nombre.split(/\s+/).length < 2 || nombre.length < 5) throw new Error('Escribe tu nombre completo.');
    if (!/^[A-Z0-9][A-Z0-9 .-]{3,19}$/.test(registro)) throw new Error('Escribe tu Registro o DNI.');
    if (leerPersonal_().indexOf(personal) < 0) throw new Error('Elige tu tipo de personal.');
    if (!descuento) throw new Error('Indica si el consumo es con descuento por planilla.');
    const pago = descuento === 'No' ? String(p.pago || '') : '';
    if (descuento === 'No' && METODOS_PAGO.indexOf(pago) < 0) throw new Error('Elige cómo vas a pagar.');

    const s = leerServicios_().filter(x => x.key === keyS_(p.servicio))[0];
    if (!s) throw new Error('Elige un servicio.');
    const modalidad = String(p.modalidad || '');
    if (!['Local', 'Recojo'].includes(modalidad)) throw new Error('Elige consumo en local o recojo.');
    if (p.fecha !== hoyKey_()) throw new Error('Solo se reciben pedidos para hoy. Actualiza el menú.');
    const solicitud = String(p.solicitud || '');
    if (!/^[a-zA-Z0-9-]{16,80}$/.test(solicitud)) throw new Error('Actualiza la página e intenta de nuevo.');
    const hoy = hoyKey_();
    const menu = unicos_(leerMenuTodo_().filter(m => m.fecha === hoy && m.sKey === s.key));
    const previas = leerLineas_(hoy, true).filter(l => l.sKey === s.key);
    const vigentes = previas.filter(l => !l.anulado);
    const repetidas = vigentes.filter(l => l.solicitud === solicitud && norm_(l.registro) === norm_(registro));
    if (repetidas.length) return comprobante_(repetidas);
    const ahora = minutosAhora_();
    if (ahora < s.abre || ahora >= s.cierra)
      throw new Error('Los pedidos de ' + s.nombre + ' están cerrados. Horario: ' + hhmm_(s.abre) + ' a ' + hhmm_(s.cierra) + '.');

    if (vigentes.some(l => norm_(l.registro) === norm_(registro)))
      throw new Error('Ya tienes un pedido de ' + s.nombre.toLowerCase() + ' hoy (Registro/DNI ' + registro + '). Si necesitas cambiarlo, avisa al comedor.');

    const precio = Object.create(null);
    (leerOpciones_()[s.key] || []).filter(o => o.modalidad === modalidad).forEach(o => precio[o.tipo] = o.precio);
    if (!Array.isArray(p.items) || !p.items.length || p.items.length > 100) throw new Error('Elige entre 1 y 100 opciones.');
    const items = p.items.map(i => {
      if (!i || !Number.isInteger(i.cant) || i.cant < 1 || i.cant > MAX_POR_PERSONA)
        throw new Error('Cantidad no válida. Máximo ' + MAX_POR_PERSONA + ' por opción.');
      return { plato: String(i.plato), tipo: String(i.tipo), cant: i.cant, precioMostrado: i.precio };
    });
    const combinaciones = new Set();
    const pide = Object.create(null);
    items.forEach(i => {
      const m = menu.filter(x => x.plato.toLowerCase() === i.plato.toLowerCase())[0];
      if (!m) throw new Error(i.plato + ' ya no está en el menú de hoy.');
      i.plato = m.plato;
      const combinacion = JSON.stringify([i.plato, i.tipo]);
      if (combinaciones.has(combinacion)) throw new Error('Opción repetida. Actualiza el pedido.');
      combinaciones.add(combinacion);
      if (!(i.tipo in precio)) throw new Error('La opción "' + i.tipo + '" ya no está disponible.');
      if (typeof i.precioMostrado !== 'number' || Math.abs(i.precioMostrado - precio[i.tipo]) > 0.001)
        throw new Error('El precio cambió. Revisa el menú y confirma nuevamente.');
      if (i.cant > MAX_POR_PERSONA) throw new Error('Máximo ' + MAX_POR_PERSONA + ' por opción.');
      pide[m.plato] = (pide[m.plato] || 0) + i.cant;
    });
    const usados = usados_(vigentes);
    Object.keys(pide).forEach(pl => {
      const m = menu.filter(x => x.plato === pl)[0];
      if (m.stock == null) return;
      const quedan = Math.max(0, m.stock - (usados[claveUso_(hoy, s.key, pl)] || 0));
      if (pide[pl] > quedan)
        throw new Error((quedan ? 'Solo quedan ' + quedan + ' de ' + pl : pl + ' se acaba de agotar') + '. Ajusta tu pedido.');
    });

    // Código por servicio y día: D-001, A-001, C-001…
    const secuencia = previas.reduce((max, l) => Math.max(max, Number(l.codigo.split('-').pop()) || 0), 0) + 1;
    const codigo = s.nombre.charAt(0).toUpperCase() + '-' + String(secuencia).padStart(3, '0');

    const fechaHora = new Date(), fTxt = hoyTxt_();
    const filas = items.map(i => [fechaHora, codigo, fTxt, s.nombre, nombre, personal, registro, i.plato, i.tipo, i.cant,
      precio[i.tipo], Math.round(i.cant * precio[i.tipo] * 100) / 100, descuento, obs, '', modalidad, solicitud,
      '', empresa, area, '', pago]);
    const sh = hojaPedidos_(), fila = sh.getLastRow() + 1;
    sh.getRange(fila, 1, filas.length, CAB_PED.length).setValues(filas);
    SpreadsheetApp.flush();
    try { actualizarPanel_(); } catch (e) { Logger.log(e); }

    return comprobante_(items.map((i, n) => ({ hora: fechaHora, codigo: codigo, fecha: hoy, servicio: s.nombre, sKey: s.key,
      nombre: nombre, personal: personal, registro: registro, plato: i.plato, tipo: i.tipo, cant: i.cant, precio: precio[i.tipo],
      descuento: descuento, obs: obs, modalidad: modalidad, solicitud: solicitud, correo: '', empresa: empresa, area: area, pago: pago,
      vale: '', fila: fila + n })));
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- Lectura de hojas
function hojaMenu_() {
  const ss = ss_();
  return hoja_(HOJA_MENU) || ss.insertSheet(HOJA_MENU, 0);
}

function leerServicios_() {
  const v = hojaMenu_().getRange(P.SRV_INI, 10, P.SRV_FIN - P.SRV_INI + 1, 3).getValues();
  return v.filter(r => String(r[0]).trim())
    .map(r => ({ nombre: String(r[0]).trim(), key: keyS_(r[0]), abre: horaEnMinutos_(r[1], -1), cierra: horaEnMinutos_(r[2], -1) }))
    .filter(s => s.abre >= 0 && s.cierra > s.abre);
}

function leerMenuTodo_() {
  const v = hojaMenu_().getRange(2, 1, FILAS_MENU, 4).getValues();
  const out = [];
  v.forEach((r, i) => {
    const plato = String(r[2]).trim(), fecha = fechaKey_(r[0]), sv = String(r[1]).trim();
    if (!plato || !fecha || !sv) return;
    const stock = (r[3] === '' || r[3] === null) ? null : Math.max(0, Math.floor(Number(r[3]) || 0));
    out.push({ fila: i + 2, fecha: fecha, servicio: sv, sKey: keyS_(sv), plato: plato, stock: stock });
  });
  return out;
}

function claveUso_(fecha, sKey, plato) { return fecha + '|' + sKey + '|' + String(plato).trim().toLowerCase(); }

function unicos_(menu) {
  const vistos = {};
  return menu.filter(m => {
    const k = claveUso_(m.fecha, m.sKey, m.plato);
    if (vistos[k]) return false;
    vistos[k] = 1;
    return true;
  });
}

// Busca una hoja sin distinguir mayúsculas, tildes ni espacios ("PRECIOS", "Precios ").
function hoja_(nombre) {
  const ss = ss_(), k = keyS_(nombre).replace(/\s+/g, '');
  return ss.getSheetByName(nombre) || ss.getSheets().filter(h => keyS_(h.getName()).replace(/\s+/g, '') === k)[0] || null;
}

function hojaPrecios_() {
  let sh = hoja_(HOJA_PRECIOS);
  if (!sh) { configurarPrecios_(); sh = hoja_(HOJA_PRECIOS); }   // Si falta, se crea con los precios por defecto
  if (!sh) throw new Error('No se encontró la hoja "Precios" en "' + ss_().getName() + '". Ejecuta configurarTodo desde esa hoja.');
  return sh;
}

function opcionBase_(tipo) {
  const k = keyS_(tipo);
  return k.includes('segundo') ? 'Solo segundo' : k.includes('econom') ? 'Económico' :
    k.includes('simple') ? 'Simple' : k.includes('complet') ? 'Completo' : tipo;
}

function configurarPrecios_() {
  const ss = ss_();
  let sh = hoja_(HOJA_PRECIOS);
  if (!sh) sh = ss.insertSheet(HOJA_PRECIOS);
  else if (sh.getName() !== HOJA_PRECIOS) sh.setName(HOJA_PRECIOS);
  if (sh.getLastRow() && String(sh.getRange('A1').getValue()).trim() !== 'Servicio') {
    // Conserva íntegra cualquier estructura más antigua.
    let name = 'Precios (anterior)', i = 2;
    while (ss.getSheetByName(name)) name = 'Precios (anterior ' + i++ + ')';
    sh.setName(name); sh = ss.insertSheet(HOJA_PRECIOS);
  }
  const migrar = sh.getRange('D1').getValue() !== 'Modalidad';
  if (migrar) {
    const existentes = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow()-1, 3).getValues().filter(r => r[0] && r[1]) : [];
    const filas = existentes.map(r => [r[0], opcionBase_(String(r[1])), r[2],
      /llevar|recojo/i.test(r[1]) ? 'Recojo' : 'Local', '', 'Sí']);
    PRECIOS_DEF.forEach(r => {
      if (!filas.some(f => keyS_(f[0]) === keyS_(r[0]) && f[1] === r[1] && f[3] === r[3]))
        filas.push([r[0], r[1], r[2], r[3], '', r[4]]);
    });
    // Columna E (personal) se conserva durante la migración.
    filas.forEach((r,i) => { sh.getRange(i+2,1,1,4).setValues([r.slice(0,4)]); sh.getRange(i+2,6).setValue(r[5]); });
  }
  sh.getRange('A1:G1').setValues([['Servicio','Opción','Precio (S/)','Modalidad','Tipo de personal','Activo','Incluye (opcional)']])
    .setFontWeight('bold').setBackground(AZUL).setFontColor('#ffffff');
  if (!sh.getRange(2,5,Math.max(1,sh.getLastRow()-1),1).getValues().some(r => r[0]))
    sh.getRange(2,5,PERSONAL_DEFECTO.length,1).setValues(PERSONAL_DEFECTO.map(p => [p]));
  sh.getRange('C2:C').setNumberFormat('"S/ "0.00');
  sh.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Local','Recojo'],true).setAllowInvalid(false).build());
  sh.getRange('F2:F').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Sí','No'],true).setAllowInvalid(false).build());
  sh.getRange('G1').setNote('Escribe qué incluye cada opción. No se presume el contenido del completo o económico.');
  [120,180,110,110,180,80,300].forEach((w,i) => sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1);
}

function leerOpciones_() {
  const sh = hojaPrecios_(), n = sh.getLastRow()-1, o = Object.create(null), vistos = new Set();
  if (n < 1) return o;
  sh.getRange(2,1,n,7).getValues().forEach(r => {
    const k = keyS_(r[0]), tipo = String(r[1]).trim(), modalidad = modalidad_(r[3]), precio = precio_(r[2]);
    if (!k || !tipo || !activo_(r[5]) || !modalidad || precio === null) return;
    const id = JSON.stringify([k,modalidad,keyS_(tipo)]);
    if (vistos.has(id)) throw new Error('Opción duplicada en Precios: ' + r[0] + ' / ' + modalidad + ' / ' + tipo);
    vistos.add(id);
    (o[k] = o[k] || []).push({tipo:tipo,modalidad:modalidad,precio:precio,incluye:String(r[6] || '')});
  });
  return o;
}

// Tolera precios escritos como texto ("S/ 10,50"), modalidad en minúsculas y Activo como casilla.
function precio_(v) {
  if (typeof v === 'string') {
    const t = v.replace(/s\/|\s/gi, '');
    v = /^\d+(,\d{1,2})?$/.test(t) ? Number(t.replace(',', '.')) : (/^\d+(\.\d+)?$/.test(t) ? Number(t) : NaN);
  }
  return typeof v === 'number' && Number.isFinite(v) && v >= 0 ? Math.round(v * 100) / 100 : null;
}
function modalidad_(v) {
  const k = keyS_(v);
  return k === 'local' ? 'Local' : (k === 'recojo' || k.includes('llevar')) ? 'Recojo' : '';
}
function activo_(v) { return v === true || ['si', 'true', 'x'].includes(keyS_(v)); }

// ---------------------------------------------------------------- Hoja Platos (ficha nutricional)
const CAB_PLATOS = ['Plato', 'Ingredientes / acompañamiento', 'Kcal', 'Proteínas (g)', 'Grasas (g)'];

function hojaPlatos_() {
  const ss = ss_();
  let sh = hoja_(HOJA_PLATOS);
  if (sh) return sh;
  sh = ss.insertSheet(HOJA_PLATOS);
  sh.getRange(1, 1, 1, CAB_PLATOS.length).setValues([CAB_PLATOS]).setFontWeight('bold')
    .setBackground(AZUL).setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  [230, 380, 70, 100, 90].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('B2:B').setWrap(true);
  sh.getRange('C2:E').setHorizontalAlignment('center').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0).setHelpText('Número por porción (ej. 520). Déjalo vacío si no lo sabes.').setAllowInvalid(false).build());
  sh.getRange('A1').setNote('Los platos del menú se agregan solos. Completa ingredientes y valores por porción: se muestran en el formulario. Lo que quede vacío no se muestra.');
  return sh;
}

// Clave del plato → ficha. Se busca sin distinguir mayúsculas ni tildes.
function leerPlatos_() {
  const sh = hoja_(HOJA_PLATOS), out = Object.create(null);
  if (!sh || sh.getLastRow() < 2) return out;
  const num = v => (v === '' || v == null || !Number.isFinite(Number(v))) ? null : Math.round(Number(v));
  sh.getRange(2, 1, sh.getLastRow() - 1, CAB_PLATOS.length).getValues().forEach(r => {
    const k = keyS_(r[0]);
    if (k) out[k] = { ingredientes: String(r[1] || '').trim(), kcal: num(r[2]), proteinas: num(r[3]), grasas: num(r[4]) };
  });
  return out;
}

// Agrega a "Platos" los platos del menú que aún no están.
function sincronizarPlatos_() {
  const sh = hojaPlatos_(), existentes = leerPlatos_(), nuevos = [];
  leerMenuTodo_().forEach(m => {
    const k = keyS_(m.plato);
    if (!(k in existentes)) { existentes[k] = {}; nuevos.push([m.plato]); }
  });
  if (nuevos.length) sh.getRange(sh.getLastRow() + 1, 1, nuevos.length, 1).setValues(nuevos);
  return nuevos.length;
}

function agregarPlatos_() {
  const n = sincronizarPlatos_();
  ss_().setActiveSheet(hojaPlatos_());
  aviso_(n ? '✅ Se agregaron ' + n + ' platos a la hoja "Platos". Completa sus ingredientes y valores.' : 'La hoja "Platos" ya tiene todos los platos del menú.');
}

function leerPersonal_() {
  const sh = hojaPrecios_();
  const n = Math.max(sh.getLastRow() - 1, 1);
  const lista = sh.getRange(2, 5, n, 1).getValues().map(r => String(r[0]).trim()).filter(String);
  return lista.length ? lista : PERSONAL_DEFECTO;
}

function hojaPedidos_() {
  const ss = ss_();
  let sh = hoja_(HOJA_PEDIDOS);
  if (sh && String(sh.getRange(1, 3).getValue()) !== CAB_PED[2]) {
    // Formato anterior: se guarda aparte y se crea la hoja nueva
    let nombre = 'Pedidos (anterior)', i = 2;
    while (ss.getSheetByName(nombre)) nombre = 'Pedidos (anterior ' + (i++) + ')';
    sh.setName(nombre);
    sh = null;
  }
  if (!sh) {
    sh = ss.insertSheet(HOJA_PEDIDOS);
    sh.getRange(1, 1, 1, CAB_PED.length).setValues([CAB_PED]).setFontWeight('bold')
      .setBackground(AZUL).setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange('C:C').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@');
    sh.getRange('K:L').setNumberFormat('"S/ "#,##0.00');
    sh.setColumnWidth(5, 220); sh.setColumnWidth(8, 200); sh.setColumnWidth(9, 220);
    sh.getRange('O1').setNote('Escribe "Anulado" en esta columna para anular una línea: el stock se devuelve.');
  }
  if (String(sh.getRange(1, CAB_PED.length).getValue()) !== CAB_PED[CAB_PED.length - 1])
    sh.getRange(1, 16, 1, CAB_PED.length - 15).setValues([CAB_PED.slice(15)]).setFontWeight('bold')
      .setBackground(AZUL).setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  return sh;
}

// Líneas de pedido (de una fecha "aaaa-mm-dd" o de todas si fecha es null)
function leerLineas_(fecha, incluirAnulados) {
  const sh = hojaPedidos_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, CAB_PED.length).getValues().map((r, i) => ({
    hora: r[0], codigo: String(r[1]), fecha: fechaKey_(r[2]), servicio: String(r[3]), sKey: keyS_(r[3]),
    nombre: String(r[4]), personal: String(r[5]), registro: String(r[6]), plato: String(r[7]), tipo: String(r[8]),
    cant: Number(r[9]) || 0, precio: Number(r[10]) || 0, descuento: String(r[12]), obs: String(r[13]),
    modalidad: String(r[15] || (/llevar/i.test(r[8]) ? 'Recojo' : 'Sin especificar')), solicitud: String(r[16] || ''),
    anulado: String(r[14]).toLowerCase().indexOf('anul') >= 0,
    correo: String(r[17] || ''), empresa: String(r[18] || ''), area: String(r[19] || ''), vale: String(r[20] || ''), pago: String(r[21] || ''), fila: i + 2
  })).filter(l => l.fecha && (!fecha || l.fecha === fecha) && (incluirAnulados || !l.anulado));
}

function comprobante_(lineas) {
  const l = lineas[0];
  return {codigo:l.codigo, servicio:l.servicio, sKey:l.sKey, modalidad:l.modalidad, nombre:l.nombre, registro:l.registro,
    hora:Utilities.formatDate(l.hora,tz_(),'dd/MM/yyyy HH:mm'), fechaKey:l.fecha, personal:l.personal, descuento:l.descuento, pago:l.pago,
    obs:l.obs, correo:l.correo, empresa:l.empresa, area:l.area, solicitud:l.solicitud, vale:l.vale, filas:lineas.map(i => i.fila),
    items:lineas.map(i => ({plato:i.plato,tipo:i.tipo,cant:i.cant,precio:i.precio,subtotal:Math.round(i.cant*i.precio*100)/100})),
    total:Math.round(lineas.reduce((t,i) => t+i.cant*i.precio,0)*100)/100};
}

function usados_(lineas) {
  const u = {};
  lineas.forEach(l => { const k = claveUso_(l.fecha, l.sKey, l.plato); u[k] = (u[k] || 0) + l.cant; });
  return u;
}

// ---------------------------------------------------------------- Vale de consumo (PDF)
function html_(s) { return String(s == null ? '' : s).replace(/^'/, '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function numeroVale_(v) { return v.fechaKey.replace(/-/g, '') + '-' + v.codigo; }
function fechaVale_(v) { return v.fechaKey.split('-').reverse().join('/'); }

// Genera el vale una sola vez por pedido, lo guarda en Drive y anota el link en Pedidos.
// Devuelve el PDF, o null si ya estaba guardado.
function guardarVale_(v) {
  const cache = CacheService.getScriptCache(), clave = 'vale-' + v.solicitud;
  if (v.vale || cache.get(clave)) return null;
  cache.put(clave, '1', 21600);
  const pdf = valePdf_(v);
  const archivo = carpetaVales_(v.fechaKey).createFile(pdf);
  const sh = hojaPedidos_();
  v.filas.forEach(f => sh.getRange(f, COL_VALE).setValue(archivo.getUrl()));
  v.vale = archivo.getUrl();
  return pdf;
}

function correoVale_(v, correo, pdf) {
  const modo = v.modalidad === 'Recojo' ? 'para llevar' : 'consumo en local';
  MailApp.sendEmail({
    to: correo, name: 'Comedor Sodexo', attachments: [pdf],
    subject: 'Vale de consumo N° ' + numeroVale_(v) + ' · ' + v.servicio + ' ' + fechaVale_(v),
    htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1b2333">' +
      '<p>Hola ' + html_(v.nombre) + ',</p>' +
      '<p>Registramos tu pedido de <b>' + html_(v.servicio) + '</b> (' + modo + ') del ' + fechaVale_(v) + '.</p>' +
      '<table style="border-collapse:collapse;font-size:14px">' + v.items.map(i =>
        '<tr><td style="padding:4px 12px 4px 0">' + i.cant + ' × ' + html_(i.plato) + ' · ' + html_(i.tipo) + '</td>' +
        '<td style="padding:4px 0;text-align:right">S/ ' + i.subtotal.toFixed(2) + '</td></tr>').join('') +
      '<tr><td style="padding:8px 12px 0 0;border-top:1px solid #999"><b>Total</b></td>' +
      '<td style="padding:8px 0 0;border-top:1px solid #999;text-align:right"><b>S/ ' + v.total.toFixed(2) + '</b></td></tr></table>' +
      '<p>Código para el comedor: <b style="font-size:18px;color:#1f3864">' + html_(v.codigo) + '</b></p>' +
      '<p>Adjuntamos la copia de tu vale de consumo en PDF.</p></div>'
  });
}

function buscarCarpeta_(dentro, nombre) { const it = dentro.getFoldersByName(nombre); return it.hasNext() ? it.next() : dentro.createFolder(nombre); }

// Carpeta "Vales de consumo" junto a la hoja de cálculo.
function carpetaValesRaiz_() {
  const padres = DriveApp.getFileById(ss_().getId()).getParents();
  return buscarCarpeta_(padres.hasNext() ? padres.next() : DriveApp.getRootFolder(), CARPETA_VALES);
}

function carpetaVales_(fechaKey) { return buscarCarpeta_(carpetaValesRaiz_(), fechaKey.slice(0, 7)); }   // subcarpeta por mes: 2026-09

// Link "Abrir carpeta" en el panel de la hoja Menú.
function linkCarpeta_(sh) {
  const url = carpetaValesRaiz_().getUrl();
  sh.getRange(P.CARPETA).setRichTextValue(SpreadsheetApp.newRichTextValue().setText('Abrir carpeta ↗').setLinkUrl(url).build())
    .setFontWeight('bold');
  return url;
}

function abrirCarpetaVales_() {
  const url = linkCarpeta_(hojaMenu_());
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:6px"><p style="margin:0 0 12px">Aquí se guardan los vales en PDF, en una subcarpeta por mes.</p>' +
    '<a href="' + url + '" target="_blank" style="display:inline-block;background:#1f3864;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none;font-weight:bold">📂 Abrir carpeta de vales</a></div>' +
    '<script>window.open(' + JSON.stringify(url) + ',"_blank");</script>').setWidth(380).setHeight(130), 'Vales de consumo');
}

// Vale en una hoja horizontal de 15 × 10.5 cm (se arma en Google Docs y se exporta a PDF).
const VALE_ANCHO = 425, VALE_ALTO = 298, VALE_MARGEN = 14;   // puntos (1 cm = 28.35 pt)

function valePdf_(v) {
  const DA = DocumentApp, A = DA.Attribute;
  const doc = DA.create('tmp-vale-' + numeroVale_(v));
  const body = doc.getBody();
  body.setPageWidth(VALE_ANCHO).setPageHeight(VALE_ALTO)
    .setMarginTop(VALE_MARGEN).setMarginBottom(VALE_MARGEN).setMarginLeft(VALE_MARGEN + 4).setMarginRight(VALE_MARGEN + 4);
  const util = VALE_ANCHO - 2 * (VALE_MARGEN + 4);
  const txt = s => String(s == null ? '' : s).replace(/^'/, '');
  const box = on => on ? '☒' : '☐';

  // Da formato a todas las celdas de una tabla: tamaño de letra, sin espacios extra y relleno mínimo.
  const formato = (t, tam, anchos, borde) => {
    t.setBorderWidth(borde ? 0.75 : 0).setBorderColor('#555555');
    anchos.forEach((w, i) => t.setColumnWidth(i, w));
    for (let r = 0; r < t.getNumRows(); r++) {
      const fila = t.getRow(r);
      for (let c = 0; c < fila.getNumCells(); c++) {
        const cel = fila.getCell(c);
        cel.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(3).setPaddingRight(3).setVerticalAlignment(DA.VerticalAlignment.CENTER);
        for (let k = 0; k < cel.getNumChildren(); k++) {
          const p = cel.getChild(k).asParagraph();
          p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
          p.setAttributes({ [A.FONT_SIZE]: tam, [A.FONT_FAMILY]: 'Arial' });
        }
      }
    }
    return t;
  };
  const estilo = (cel, o) => {
    const t = cel.editAsText();
    if (!t.getText().length) return;
    if (o.b) t.setBold(true);
    if (o.i) t.setItalic(true);
    if (o.tam) t.setFontSize(o.tam);
    if (o.color) t.setForegroundColor(o.color);
    if (o.al) cel.getChild(0).asParagraph().setAlignment(o.al);
  };
  const sep = alto => body.appendParagraph('').setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1)
    .setAttributes({ [A.FONT_SIZE]: alto });
  const H = DA.HorizontalAlignment;

  // Encabezado
  const p0 = body.getParagraphs()[0];
  p0.setSpacingBefore(0).setSpacingAfter(0).setAttributes({ [A.FONT_SIZE]: 2 });
  const cab = formato(body.appendTable([
    ['sodexo', 'VALE DE CONSUMO', 'N° ' + numeroVale_(v)],
    ['Comedor Staff Ilo ' + box(COMEDOR === 'Staff Ilo') + '    Comedor Hospital ' + box(COMEDOR === 'Hospital'), '', 'Fecha: ' + fechaVale_(v)]
  ]), 9, [110, 160, util - 270], false);
  estilo(cab.getCell(0, 0), { b: true, i: true, tam: 17, color: '#2a295c' });
  estilo(cab.getCell(0, 1), { i: true, tam: 13, al: H.CENTER });
  estilo(cab.getCell(0, 2), { tam: 12, al: H.RIGHT });
  estilo(cab.getCell(1, 2), { b: true, al: H.RIGHT });
  sep(3);

  // Datos de la persona
  const tipos = ['SPCC', 'Funcionario', 'Empleado', 'Contratista'];
  let tipoTxt = tipos.map(p => p + ' ' + box(keyS_(p) === keyS_(v.personal))).join('     ');
  if (!tipos.some(p => keyS_(p) === keyS_(v.personal))) tipoTxt += '     ' + txt(v.personal) + ' ☒';
  const datos = formato(body.appendTable([
    ['Nombre:', txt(v.nombre), 'Registro/DNI:', txt(v.registro)],
    ['Personal:', tipoTxt, '', ''],
    ['Empresa:', txt(v.empresa), 'Dpto./Área:', txt(v.area)],
    ['Cuenta:', '', '', '']
  ]), 8.5, [52, 190, 62, util - 304], true);
  [[0, 1], [0, 3], [2, 1], [2, 3]].forEach(rc => estilo(datos.getCell(rc[0], rc[1]), { b: true }));
  sep(3);

  // Servicios: cantidad, precio y firma
  const filasSrv = [['Desayuno', /^desayuno/], ['Almuerzo', /^almuerzo/], ['Cena', /^cena/], ['Rancho', /rancho/]];
  let usada = filasSrv.findIndex(f => f[1].test(v.sKey));
  const etiquetas = filasSrv.map(f => f[0]);
  if (usada < 0) { etiquetas.push(txt(v.servicio)); usada = etiquetas.length - 1; }
  const cant = v.items.reduce((t, i) => t + i.cant, 0);
  const grid = formato(body.appendTable([['', 'Cant.', 'Precio S/', 'Firma']].concat(etiquetas.map((et, i) =>
    [et, i === usada ? String(cant) : '', 'S/ ' + (i === usada ? v.total.toFixed(2) : ''), i === usada && !v.firma ? 'Pedido web · ' + v.codigo : '']))),
    8.5, [80, 55, 80, util - 215], true);
  for (let c = 0; c < 4; c++) { estilo(grid.getCell(0, c), { b: true, al: H.CENTER }); grid.getCell(0, c).setBackgroundColor('#eef2f9'); }
  estilo(grid.getCell(usada + 1, 1), { b: true, al: H.CENTER });
  estilo(grid.getCell(usada + 1, 2), { b: true });
  if (v.firma) {
    const img = grid.getCell(usada + 1, 3).getChild(0).asParagraph().setAlignment(H.CENTER)
      .appendInlineImage(Utilities.newBlob(Utilities.base64Decode(v.firma.split(',')[1]), 'image/png', 'firma.png'));
    const alto = 26;
    img.setWidth(Math.round(img.getWidth() * alto / img.getHeight())).setHeight(alto);
  }
  sep(3);

  // Total y descuento por planilla
  const pie = formato(body.appendTable([
    [box(v.descuento === 'Sí') + '  Sujeto a descuento por Planilla', 'CONSUMO TOTAL'],
    [box(v.descuento === 'No') + '  NO sujeto a descuento por Planilla' + (v.pago ? '   ·   Pago: ' + v.pago : ''), 'S/ ' + v.total.toFixed(2)]
  ]), 9, [util - 120, 120], false);
  estilo(pie.getCell(0, 1), { b: true, al: H.RIGHT });
  estilo(pie.getCell(1, 1), { b: true, tam: 13, al: H.RIGHT });
  sep(3);

  // Detalle
  const det = body.appendParagraph('Detalle: ' + txt(v.servicio) + ' · ' + (v.modalidad === 'Recojo' ? 'Para llevar' : 'Consumo en local') +
    ' · ' + v.items.map(i => i.cant + '× ' + txt(i.plato) + ' ' + txt(i.tipo) + ' (S/ ' + i.precio.toFixed(2) + ')').join(', ') +
    (v.obs ? ' · Obs.: ' + txt(v.obs) : '') + ' · Registrado ' + v.hora);
  det.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  det.editAsText().setFontSize(7).setForegroundColor('#444444').setFontFamily('Arial');

  doc.saveAndClose();
  const archivo = DriveApp.getFileById(doc.getId());
  const pdf = archivo.getAs(MimeType.PDF)
    .setName('Vale ' + numeroVale_(v) + ' - ' + txt(v.nombre).replace(/[\\/:*?"<>|]/g, '') + '.pdf');
  archivo.setTrashed(true);
  return pdf;
}

// ---------------------------------------------------------------- Link
function linkWeb_() {
  const sh = hojaMenu_();
  const k2 = String(sh.getRange(P.LINK).getValue()).trim();
  if (/\/exec$/.test(k2)) return k2;
  let u = '';
  try { u = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  if (/\/exec$/.test(u)) { sh.getRange(P.LINK).setValue(u); return u; }
  return '';
}

function verLink_() {
  const link = linkWeb_();
  aviso_(link ? 'Link para WhatsApp:\n' + link
    : 'Aún no hay link. Implementar → Nueva implementación → Aplicación web, y pega la URL que termina en /exec en la celda K2 de la hoja Menú.');
}

// ---------------------------------------------------------------- Interacción en la hoja
function alEditar_(e) {
  if (!e || !e.range) return;
  const nombre = e.range.getSheet().getName(), r = e.range.getRow(), c = e.range.getColumn();
  const enMenu = nombre === HOJA_MENU && e.range.getLastRow() > 1 &&
    (c <= 4 || (e.range.getLastColumn() >= 10 && c <= 12 && e.range.getLastRow() >= P.SRV_INI && r <= P.SRV_FIN));
  const enPedidos = nombre === HOJA_PEDIDOS && e.range.getLastColumn() >= 15;
  if (enMenu || enPedidos || nombre === HOJA_PRECIOS) {
    try { actualizarStock_(); } catch (err) { Logger.log(err); }
  }
  if (nombre === HOJA_MENU && c <= 3 && e.range.getLastColumn() >= 3 && e.range.getLastRow() > 1) {
    try { sincronizarPlatos_(); } catch (err) { Logger.log(err); }
  }
}

function controlHorario_() { actualizarStock_(); }

// ---------------------------------------------------------------- Panel y conteos
function actualizarStock_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { actualizarPanel_(); } finally { lock.releaseLock(); }
}

function actualizarPanel_() {
  const sh = hojaMenu_(), hoy = hoyKey_();
  const lineas = leerLineas_(null, false);

  // Pedidos por fecha|servicio|plato y detalle por opción
  const cnt = {}, det = {};
  lineas.forEach(l => {
    const k = claveUso_(l.fecha, l.sKey, l.plato);
    cnt[k] = (cnt[k] || 0) + l.cant;
    det[k] = det[k] || {};
    const etiqueta = l.modalidad + ' · ' + l.tipo;
    det[k][etiqueta] = (det[k][etiqueta] || 0) + l.cant;
  });

  const porFila = {};
  leerMenuTodo_().forEach(m => porFila[m.fila] = m);
  const vistos = {}, out = [];
  for (let f = 2; f < 2 + FILAS_MENU; f++) {
    const m = porFila[f];
    if (!m) { out.push(['', '', '']); continue; }
    const k = claveUso_(m.fecha, m.sKey, m.plato);
    if (vistos[k]) { out.push(['', '', '⚠️ Repetido: se usa la primera fila']); continue; }
    vistos[k] = 1;
    const ped = cnt[k] || 0;
    const detalle = det[k] ? Object.keys(det[k]).map(t => det[k][t] + ' ' + t).join(' · ') : '';
    out.push([ped, m.stock == null ? 'Sin límite' : Math.max(0, m.stock - ped), detalle]);
  }
  sh.getRange(2, 5, FILAS_MENU, 3).setValues(out);

  // Venta de hoy
  const venta = lineas.filter(l => l.fecha === hoy).reduce((a, l) => a + l.cant * l.precio, 0);
  sh.getRange(P.VENTA).setValue(Math.round(venta * 100) / 100).setNumberFormat('"S/ "#,##0.00');

  // Estado según horario
  const ahora = minutosAhora_(), servicios = leerServicios_();
  const abierto = servicios.filter(s => ahora >= s.abre && ahora < s.cierra)[0];
  const proximo = servicios.filter(s => ahora < s.abre).sort((a, b) => a.abre - b.abre)[0];
  const estado = sh.getRange(P.ESTADO);
  if (abierto) estado.setValue('🟢 ' + abierto.nombre.toUpperCase() + ' abierto hasta las ' + hhmm_(abierto.cierra))
    .setBackground('#d9ead3').setFontColor('#274e13');
  else if (proximo) estado.setValue('🔴 CERRADO — próximo: ' + proximo.nombre + ' a las ' + hhmm_(proximo.abre))
    .setBackground('#f4cccc').setFontColor('#990000');
  else estado.setValue('🔴 CERRADO por hoy').setBackground('#f4cccc').setFontColor('#990000');
}

// Separador de fórmulas según el idioma de la hoja (en español es ";")
function fx_(s) {
  if (/^en/.test(ss_().getSpreadsheetLocale())) return s;
  let out = '', q = false;
  for (const ch of s) { if (ch === '"') q = !q; out += (!q && ch === ',') ? ';' : ch; }
  return out;
}

// ---------------------------------------------------------------- Diseño de la hoja Menú
function disenarHoja_() {
  const sh = hojaMenu_();
  const F = FILAS_MENU;

  // Conserva link y horarios actuales
  const k2 = String(sh.getRange(P.LINK).getValue()).trim();
  let srv = sh.getRange(P.SRV_INI, 10, P.SRV_FIN - P.SRV_INI + 1, 3).getValues();
  if (!srv.some(r => String(r[0]).trim())) srv = SERVICIOS_DEF.concat([['', '', ''], ['', '', '']]);

  // ---- Tabla del menú (A:H)
  sh.getRange(1, 1, F + 1, 8).clearFormat().clearDataValidations();
  sh.getRange(1, 1, 1, 8).setValues([['Fecha', 'Servicio', 'Plato', 'Stock', 'Pedidos', 'Quedan', 'Detalle de pedidos', 'Avance']])
    .setFontWeight('bold').setBackground(AZUL).setFontColor('#ffffff').setWrap(true)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  const cuerpo = sh.getRange(2, 1, F, 8);
  cuerpo.setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID).setVerticalAlignment('middle');
  sh.getRange(2, 1, F, 4).setBackground(AMARILLO);
  sh.getRange(2, 1, F, 1).setNumberFormat('ddd dd/mm/yyyy').setHorizontalAlignment('left')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setHelpText('Escribe la fecha (ej. 29/09/2026).').setAllowInvalid(false).build());
  sh.getRange(2, 2, F, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getRange(P.SRV_INI, 10, P.SRV_FIN - P.SRV_INI + 1, 1), true)
    .setHelpText('Elige el servicio (se definen en el panel, columna J).').setAllowInvalid(false).build());
  sh.getRange(2, 3, F, 1).setFontWeight('bold');
  sh.getRange(2, 4, F, 1).setHorizontalAlignment('center').setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0)
      .setHelpText('Porciones disponibles. Déjalo vacío si no tiene límite.').setAllowInvalid(false).build());
  sh.getRange(2, 5, F, 2).setHorizontalAlignment('center');
  sh.getRange(2, 7, F, 1).setFontSize(9).setFontColor('#555555');
  sh.getRange('A1').setNote('Celdas amarillas: programa aquí el menú de la semana (fecha, servicio, plato y stock). Las columnas E a H se llenan solas.');

  const f = [];
  for (let r = 2; r <= F + 1; r++) {
    const pct = 'MIN(10,ROUND(10*$E' + r + '/$D' + r + '))';
    f.push([fx_('=IF(OR($C' + r + '="",$D' + r + '="",$D' + r + '=0),"",REPT("█",' + pct + ')&REPT("░",10-' + pct + ')&"  "&ROUND(100*$E' + r + '/$D' + r + ')&"%")')]);
  }
  sh.getRange(2, 8, F, 1).setFormulas(f).setFontColor('#6aa84f').setFontFamily('Roboto Mono');

  [120, 100, 250, 65, 70, 80, 280, 130, 24, 190, 100, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const rE = sh.getRange(2, 6, F, 1), rH = sh.getRange(2, 8, F, 1), rTodo = sh.getRange(2, 1, F, 8);
  const reglas = [
    [fx_('=AND($C2<>"",ISNUMBER($F2),$F2=0)'), rE, '#f4cccc', '#990000', true],
    [fx_('=AND($C2<>"",ISNUMBER($F2),$F2>0,$F2<=5)'), rE, '#fff2cc', '#7f6000', true],
    [fx_('=AND($C2<>"",ISNUMBER($F2),$F2>5)'), rE, '#d9ead3', '#274e13', true],
    [fx_('=AND($C2<>"",ISNUMBER($D2),$D2>0,$E2>=$D2)'), rH, null, '#cc0000', false],
    [fx_('=AND($C2<>"",ISNUMBER($D2),$D2>0,$E2>=0.8*$D2)'), rH, null, '#bf9000', false],
    ['=$A2=TODAY()', rTodo, '#dbe8fb', null, false],
    [fx_('=AND($A2<>"",$A2<TODAY())'), rTodo, null, '#a0a0a0', false]
  ].map(x => {
    const b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(x[0]).setRanges([x[1]]);
    if (x[2]) b.setBackground(x[2]);
    if (x[3]) b.setFontColor(x[3]);
    if (x[4]) b.setBold(true);
    return b.build();
  });
  sh.setConditionalFormatRules(reglas);

  // ---- Panel de control (J:L)
  sh.getRange('I1:M30').breakApart().clear().clearDataValidations().clearNote();
  sh.setRowHeights(2, 28, 21);
  sh.getRange('J1:L1').merge().setValue('⚙️ PANEL DE CONTROL').setBackground(AZUL).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.getRange('J2').setValue('🔗 Link del formulario');
  sh.getRange('K2:L2').merge();
  if (/^https?:\/\//.test(k2) && !/forms\.gle|docs\.google\.com\/forms/.test(k2)) sh.getRange(P.LINK).setValue(k2);
  else sh.getRange(P.LINK).setNote('Pega aquí la URL de la aplicación web (termina en /exec).');

  sh.getRange('J3:L3').setValues([['🍽️ Servicio', '🕖 Abre (h)', '🕘 Cierra (h)']])
    .setBackground(AZUL_CLARO).setFontWeight('bold').setHorizontalAlignment('center');
  const rSrv = sh.getRange(P.SRV_INI, 10, P.SRV_FIN - P.SRV_INI + 1, 3);
  rSrv.setValues(srv).setBackground(AMARILLO);
  sh.getRange(P.SRV_INI, 10, P.SRV_FIN - P.SRV_INI + 1, 1).setFontWeight('bold');
  sh.getRange(P.SRV_INI, 11, P.SRV_FIN - P.SRV_INI + 1, 2).setHorizontalAlignment('center').setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 24)
      .setHelpText('Hora en formato 24 h (ej. 5 = 5:00 a. m., 16 = 4:00 p. m., 9.5 = 9:30).').setAllowInvalid(false).build());
  sh.getRange('J' + P.SRV_INI).setNote('Puedes agregar otro servicio en las filas vacías (ej. Rancho caliente) y sus opciones en la hoja Precios.');

  sh.getRange('J9').setValue('💰 Venta estimada hoy');
  sh.getRange('K9:L9').merge().setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange('J10').setValue('📶 Estado del formulario');
  sh.getRange('K10:L10').merge().setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange('J2:J10').setFontWeight('bold');
  sh.getRange('J2').setBackground(AZUL_CLARO);
  sh.getRange('J9:J10').setBackground(AZUL_CLARO);
  sh.getRange('J2:L10').setBorder(true, true, true, true, true, true, '#b7c4de', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange('J12').setValue('📂 Vales de consumo (PDF)').setFontWeight('bold').setBackground(AZUL_CLARO);
  sh.getRange('K12:L12').merge();
  try { linkCarpeta_(sh); } catch (e) { Logger.log(e); }
  sh.getRange('J12:L12').setBorder(true, true, true, true, true, true, '#b7c4de', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange('J19:L19').merge().setValue('ℹ️ CÓMO USARLO').setFontWeight('bold').setBackground(GRIS);
  sh.getRange('J20:L27').merge().setWrap(true).setVerticalAlignment('top').setBackground(GRIS).setValue(
    '1. En las celdas amarillas (A a D) programa el menú de la semana: fecha, servicio, plato y stock. Stock vacío = sin límite.\n' +
    '2. Cada servicio abre y cierra solo según su horario (tabla de arriba). El formulario muestra los platos de hoy.\n' +
    '3. Las opciones y precios de cada servicio (completo, económico, segundo) están en la hoja "Precios".\n' +
    '4. Modalidad Local/Recojo, Activo e Incluye se editan también en Precios. Los pedidos llegan a la hoja "Pedidos". Para anular uno, escribe "Anulado" en la columna Estado.\n' +
    '5. Cada pedido genera su vale en PDF. Ábrelos con el link "Abrir carpeta" (K12) o en Menú 🍽️ Comedor → Abrir carpeta de vales.');
}

// ---------------------------------------------------------------- Menú de ejemplo (semana del 21 al 27/09/2026)
// Platos de fondo del menú semanal impreso. Cada día: [platos..., acompañamiento].
const MENU_EJEMPLO = {
  'Desayuno': [
    [['Pollo a la cacerola'], 'Papa sancochada'],
    [['Arroz a la jardinera con cerdo'], 'Ensalada fresca'],
    [['Sarza de atún'], 'Papa sancochada'],
    [['Papa arrebozada'], ''],
    [['Revuelto de verduras'], 'Camote sancochado'],
    [['Lomito al jugo'], 'Papa sancochada'],
    [['Omelette'], '']
  ],
  'Almuerzo': [
    [['Arroz con pollo', 'Lomo saltado'], 'Sarza criolla, papas fritas y arroz blanco'],
    [['Ají de pollo', 'Asado de res'], 'Puré de papa y arroz blanco'],
    [['Carapulcra de cerdo', 'Pollo guisado'], 'Camote sancochado y arroz blanco'],
    [['Pollo oriental', 'Picante a la tacneña'], 'Sarza criolla y arroz blanco'],
    [['Pollo al perejil', 'Albóndigas a la boloñesa'], 'Yuca sancochada, tallarín rojo y arroz blanco'],
    [['Pollo al romero', 'Pescado frito'], 'Papa sancochada, ensalada fresca y arroz blanco'],
    [['Pollo al horno', 'Cerdo al horno'], 'Papa al horno, frejoles y arroz blanco']
  ],
  'Cena': [
    [['Pollo a la naranja', 'Tallarín saltado con res'], 'Camote sancochado y arroz blanco'],
    [['Saltado de mollejitas', 'Pollo al sillao'], 'Papas fritas y arroz con perejil'],
    [['Asado de pollo', 'Pescado al horno'], 'Papa al horno, camote sancochado y arroz blanco'],
    [['Cerdo agridulce', 'Pollo a la mostaza'], 'Puré de camote, papa sancochada y arroz blanco'],
    [['Pollo tipo parrilla', 'Guiso de fideo con res'], 'Papa sancochada y arroz blanco'],
    [['Cerdo al horno', 'Revuelto de verduras'], 'Camote sancochado y arroz blanco'],
    [['Pollo a la plancha', 'Arroz tapado'], 'Papa sancochada y arroz blanco']
  ],
  'Rancho caliente': [
    [['Lentejita guisada con cerdo'], 'Arroz blanco · fruta mandarina'],
    [['Pollo al horno'], 'Papa sancochada, arroz blanco · pudín de fruta'],
    [['Cerdo al horno'], 'Camote sancochado, arroz blanco · mazamorra morada'],
    [['Pollo arvejado'], 'Papa sancochada, arroz blanco · compota de fruta'],
    [['Matasquita de res'], 'Arroz blanco · fruta manzana delicia'],
    [['Pollo al romero'], 'Papa sancochada, arroz blanco · fruta mandarina'],
    [['Chuleta de cerdo'], 'Papas fritas, arroz blanco · fruta manzana israel']
  ]
};

// Ejecuta esta desde el editor (Correr → cargarMenuEjemplo). No duplica filas si se ejecuta dos veces.
function cargarMenuEjemplo() {
  const sh = hojaMenu_(), servicios = leerServicios_();
  const existentes = new Set(leerMenuTodo_().map(m => claveUso_(m.fecha, m.sKey, m.plato)));
  const col = sh.getRange(2, 3, FILAS_MENU, 1).getValues();
  let libre = col.length; while (libre > 0 && !String(col[libre - 1][0]).trim()) libre--;
  const filas = [], faltan = [], ingredientes = {};
  Object.keys(MENU_EJEMPLO).forEach(nombre => {
    const s = servicios.filter(x => x.key === keyS_(nombre))[0];
    if (!s) { faltan.push(nombre); return; }
    MENU_EJEMPLO[nombre].forEach((dia, i) => {
      const fecha = new Date(2026, 8, 21 + i, 12), fk = '2026-09-' + (21 + i);
      dia[0].forEach(plato => {
        if (!ingredientes[keyS_(plato)] && dia[1]) ingredientes[keyS_(plato)] = dia[1];
        if (!existentes.has(claveUso_(fk, s.key, plato))) filas.push([fecha, s.nombre, plato, 30]);
      });
    });
  });
  if (libre + filas.length > FILAS_MENU) throw new Error('No hay filas libres suficientes en la hoja Menú.');
  if (filas.length) sh.getRange(libre + 2, 1, filas.length, 4).setValues(filas);

  // Acompañamientos en la hoja Platos (solo donde está vacío)
  sincronizarPlatos_();
  const hp = hojaPlatos_(), n = hp.getLastRow() - 1;
  if (n > 0) {
    const v = hp.getRange(2, 1, n, 2).getValues();
    v.forEach(r => { if (!String(r[1]).trim() && ingredientes[keyS_(r[0])]) r[1] = ingredientes[keyS_(r[0])]; });
    hp.getRange(2, 2, n, 1).setValues(v.map(r => [r[1]]));
  }
  actualizarStock_();
  aviso_('✅ Se agregaron ' + filas.length + ' platos del 21 al 27/09/2026 (stock 30 cada uno).' +
    (faltan.length ? '\n\n⚠️ No se cargó: ' + faltan.join(', ') + '. Agrégalo en el panel (columna J de Menú) con su horario y vuelve a ejecutar.' : ''));
}

// ---------------------------------------------------------------- Configuración
// Ejecuta esta desde el editor (Correr → configurarTodo). Las funciones que terminan en "_" no aparecen en la lista.
function configurarTodo() { configurarTodo_(); }

function configurarTodo_() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Este proyecto no está vinculado a la hoja. Abre tu Google Sheet → Extensiones → Apps Script, pega ahí el código y ejecuta configurarTodo.');
  PropertiesService.getScriptProperties().setProperty('COMEDOR_SS_ID', ss.getId());
  ss.setSpreadsheetTimeZone('America/Lima'); TZ = null;
  ScriptApp.getProjectTriggers().filter(t => ['alEditar','controlHorario','alEditar_','controlHorario_'].includes(t.getHandlerFunction())).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alEditar_').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('controlHorario_').timeBased().everyMinutes(5).create();

  // Si existía la hoja "Stock" (versión anterior), trae su link
  const viejo = ss.getSheetByName('Stock');
  const sh = hojaMenu_();
  if (viejo && !/\/exec$/.test(String(sh.getRange(P.LINK).getValue()))) {
    const l = String(viejo.getRange('K2').getValue()).trim();
    if (/\/exec$/.test(l)) sh.getRange(P.LINK).setValue(l);
  }

  configurarPrecios_();
  hojaPedidos_();
  sincronizarPlatos_();
  disenarHoja_();

  ss.setActiveSheet(sh);

  const link = linkWeb_();
  actualizarStock_();

  aviso_('✅ Listo. Programa el menú en la hoja "Menú" y revisa precios en "Precios".' +
    (viejo ? '\nLa hoja "Stock" ya no se usa: puedes borrarla.' : '') +
    (link ? '\n\nLink para WhatsApp:\n' + link
      : '\n\n⚠️ Falta el link: Implementar → Nueva implementación → Aplicación web, y pega la URL (/exec) en K2.'));
}

