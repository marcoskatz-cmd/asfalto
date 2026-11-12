/**
 * STOCK DE ASFALTO Y SALIDA DE MEZCLA — Planta asfáltica INGECO
 * Apps Script + Sheets. Frontend en GitHub Pages (docs/) → API JSON por doPost.
 *
 * Carga diaria del plantista: stock de asfalto en el tanque al cierre, asfalto recibido en el día
 * y toneladas de mezcla que salieron para cada destino: los fijos de CONFIG DESTINOS (Autovías, Bacheo)
 * más otras obras que se escriben libres y quedan guardadas en CONFIG OBRAS para elegirlas la próxima vez.
 *
 * Consumo de asfalto: entre dos mediciones consecutivas del tanque,
 *   consumo = stock anterior + asfalto recibido en el tramo − stock actual
 * y se reparte en los días del tramo en proporción a la mezcla que salió cada día.
 * % real de asfalto = consumo ÷ mezcla equivalente (del tramo). Se compara con el % teórico de CONFIG.
 *
 * Purga: cada vez que se prende la planta se tiran 2 t al arrancar y 2 t al apagar (no figuran en
 * la mezcla cargada). La 1ª y la última tonelada llevan 50 % del asfalto normal, la 2ª y la anteúltima
 * 90 % → cada arranque consume el asfalto de 0,5+0,9+0,9+0,5 = 2,8 t de mezcla normal.
 * Mezcla equivalente = mezcla despachada + 2,8 t × arranques. Por defecto 1 arranque por día con mezcla.
 */
const PURGA_T = 4;            // t de mezcla que se tiran por arranque
const PURGA_EQ_T = 2.8;       // t de mezcla normal equivalentes en asfalto por arranque

const TABS = { CONFIG: 'CONFIG', DIARIO: 'DIARIO', SALIDAS: 'SALIDAS' };
const TZ = 'America/Argentina/Buenos_Aires';

// ---------- Web app ----------
const URL_APP = 'https://asfalto.ingecosa.com.ar/';
const APP_INGECO_BACKEND = 'https://script.google.com/macros/s/AKfycbx9NN1FfOQjUjfymLXrMyd2QM6l96HP0lxY6RC7_xKsbDlO0PnwdHbaMtag54VYFor0/exec';
const MODULO_INGECO = 'ASFALTO';

function doGet(e) {
  const p = (e && e.parameter) || {};
  // Setup remoto de una sola vez (clasp run no anda sin proyecto GCP): ?accion=setup&k=<SETUP_KEY de Script Properties>
  if (p.accion === 'setup') {
    const k = PropertiesService.getScriptProperties().getProperty('SETUP_KEY');
    if (!k || p.k !== k) return ContentService.createTextOutput('no');
    setup();
    return ContentService.createTextOutput('setup-ok');
  }
  return HtmlService.createHtmlOutput(
    '<meta http-equiv="refresh" content="0;url=' + URL_APP + '">' +
    '<div style="font-family:sans-serif;padding:40px;text-align:center">Abriendo la app… <a href="' + URL_APP + '">' + URL_APP + '</a></div>'
  ).setTitle('Stock de asfalto — INGECO');
}

// Body (text/plain para evitar preflight): { accion, pin | sesion | clave, nombre, payload, con_all }
function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'Cuerpo inválido' }); }
  try {
    const p = body.payload || {};
    // Integración con otros sistemas (Autovías): clave de API de CONFIG, solo lectura.
    if (body.accion === 'mezcla') return json_({ ok: true, data: mezclaParaIntegracion_(body, p) });
    // Obra Autovías avisa la mezcla que cargó; si la planta no la cargó, se registra acá (espejo).
    if (body.accion === 'salidaExterna') return json_({ ok: true, data: salidaExterna_(body, p) });
    const user = autorizar_(body);
    let data;
    switch (body.accion) {
      case 'login': data = { ok: true }; break;
      case 'datos': data = getDatos(user); break;
      case 'guardarDia': data = guardarDia(p, user); break;
      case 'borrarDia': data = borrarDia(p.fecha, user); break;
      case 'config': data = leerConfigCompleta_(); break;
      case 'guardarConfig': data = guardarConfig(p, user); break;
      default: throw new Error('Acción desconocida: ' + body.accion);
    }
    if (body.con_all && body.accion !== 'datos' && !data.all) data.all = getDatos(user);
    return json_({ ok: true, data, usuario: user });
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Devuelve el nombre de quien opera. Entra con sesión de la app INGECO (sin PIN) o con el PIN de CONFIG.
function autorizar_(body) {
  const nombre = String(body.nombre || '').trim().slice(0, 60);
  if (body.sesion && sesionIngecoOk_(body.sesion)) return nombre || 'App INGECO';
  const cfg = leerConfig_();
  const pin = String(cfg.PIN === undefined || cfg.PIN === '' ? '1234' : cfg.PIN).trim();
  if (String(body.pin || '').trim() !== pin) throw new Error('PIN incorrecto');
  return nombre || 'PIN';
}

// Sesión de módulo emitida por app-ingeco al canjear el ticket (guard.js). Se cachea 10 min.
function sesionIngecoOk_(sesion) {
  if (!sesion || String(sesion).length < 20) return false;
  const cache = CacheService.getScriptCache();
  const k = 'ingeco_ses_' + String(sesion).slice(0, 40);
  const c = cache.get(k);
  if (c === '1') return true;
  if (c === '0') return false;
  let ok = false;
  try {
    const r = UrlFetchApp.fetch(APP_INGECO_BACKEND, {
      method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true,
      payload: JSON.stringify({ accion: 'validar_sesion_modulo', payload: { sesion: String(sesion), modulo: MODULO_INGECO } })
    });
    const j = JSON.parse(r.getContentText());
    ok = !!(j && j.ok);
  } catch (err) { ok = false; }
  cache.put(k, ok ? '1' : '0', ok ? 600 : 60);
  return ok;
}

// ---------- Setup inicial ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfg = ss.getSheetByName(TABS.CONFIG);
  if (!cfg) {
    cfg = ss.insertSheet(TABS.CONFIG);
    cfg.getRange(1, 1, 1, 3).setValues([['CLAVE', 'VALOR', 'DESCRIPCIÓN']]).setFontWeight('bold');
    const rows = [
      ['PIN', '1234', 'PIN de acceso a la app (desde la app INGECO no hace falta)'],
      ['DESTINOS', 'Autovías, Bacheo', 'Adónde puede salir la mezcla, separados por coma. El primero se usa para Autovías'],
      ['DESTINO_AUTOVIAS', 'Autovías', 'Nombre del destino cuya mezcla lee el sistema de Obra Autovías'],
      ['OBRAS', '', 'Otras obras a las que salió mezcla (separadas por coma). Se agregan solas desde la app al cargar una obra nueva; acá se pueden borrar o renombrar'],
      ['PORC_TEORICO', 4.8, '% de asfalto que debería llevar la mezcla según la fórmula'],
      ['TOLERANCIA', 0.4, 'Puntos de % de desvío admitidos antes de avisar (ej. 0,4 → entre 4,4 y 5,2)'],
      ['STOCK_MINIMO', 20, 't de asfalto en tanque por debajo de las cuales se avisa'],
      ['CAPACIDAD_TANQUE', 60, 't de asfalto que entran en el tanque (solo para mostrar el % lleno)'],
      ['CLAVE_API', Utilities.getUuid().replace(/-/g, ''), 'Clave con la que Obra Autovías lee la mezcla (ponerla en ASFALTO_CLAVE del script de Autovías)'],
    ];
    cfg.getRange(2, 1, rows.length, 3).setValues(rows);
    cfg.setColumnWidth(1, 180); cfg.setColumnWidth(2, 320); cfg.setColumnWidth(3, 520);
  }
  let d = ss.getSheetByName(TABS.DIARIO);
  if (!d) {
    d = ss.insertSheet(TABS.DIARIO);
    d.getRange(1, 1, 1, 7).setValues([['FECHA', 'STOCK_T', 'RECIBIDO_T', 'OBS', 'CARGADO POR', 'TIMESTAMP', 'ARRANQUES']]).setFontWeight('bold');
    d.getRange(1, 1, d.getMaxRows(), 1).setNumberFormat('yyyy-mm-dd');
    d.setFrozenRows(1);
  }
  let s = ss.getSheetByName(TABS.SALIDAS);
  if (!s) {
    s = ss.insertSheet(TABS.SALIDAS);
    s.getRange(1, 1, 1, 5).setValues([['FECHA', 'DESTINO', 'MEZCLA_T', 'CARGADO POR', 'TIMESTAMP']]).setFontWeight('bold');
    s.getRange(1, 1, s.getMaxRows(), 1).setNumberFormat('yyyy-mm-dd');
    s.setFrozenRows(1);
  }
  const h1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (h1 && ss.getSheets().length > 1) ss.deleteSheet(h1);
}

// Correr UNA vez desde el editor para autorizar UrlFetchApp (acceso desde la app INGECO y avisos).
function autorizarAppIngeco() {
  const r = UrlFetchApp.fetch(APP_INGECO_BACKEND + '?accion=ping', { muteHttpExceptions: true });
  Logger.log('App INGECO respondió: ' + r.getContentText().slice(0, 120));
}

// ---------- Config ----------
function leerConfig_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.CONFIG);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0] !== '') out[String(r[0]).trim()] = r[1]; });
  return out;
}
function leerConfigCompleta_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.CONFIG);
  const rows = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  return { config: rows.filter(r => r[0] !== '').map(r => ({ clave: String(r[0]).trim(), valor: r[1] instanceof Date ? fechaStr_(r[1]) : r[1], desc: r[2] })) };
}
function guardarConfig(p, user) {
  if (!user) throw new Error('Sin permiso');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.CONFIG);
  const cfg = p.config || {};
  const permitidas = ['PIN', 'DESTINOS', 'DESTINO_AUTOVIAS', 'OBRAS', 'PORC_TEORICO', 'TOLERANCIA', 'STOCK_MINIMO', 'CAPACIDAD_TANQUE'];
  const claves = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues().map(r => String(r[0]).trim());
  Object.keys(cfg).forEach(k => {
    if (permitidas.indexOf(k) < 0) return;
    let v = cfg[k];
    if (k === 'PIN') { v = String(v).trim(); if (!/^\d{4,8}$/.test(v)) throw new Error('El PIN debe tener entre 4 y 8 números'); }
    else if (k === 'DESTINOS' || k === 'DESTINO_AUTOVIAS') {
      v = String(v).trim(); if (!v) throw new Error('Falta ' + k);
      if (k === 'DESTINO_AUTOVIAS') {   // tiene que ser uno de los destinos, si no Autovías no recibe la mezcla
        const lista = String(cfg.DESTINOS !== undefined ? cfg.DESTINOS : leerConfig_().DESTINOS || '').split(',').map(x => x.trim()).filter(Boolean);
        if (lista.length && !lista.some(d => normDestino_(d) === normDestino_(v))) throw new Error('DESTINO_AUTOVIAS ("' + v + '") tiene que ser uno de los destinos: ' + lista.join(', ') + '. Si no, Obra Autovías no recibe la mezcla.');
      }
    }
    else if (k === 'OBRAS') { v = limpiarObras_(String(v).split(','), destinos_(cfg.DESTINOS !== undefined ? { DESTINOS: cfg.DESTINOS } : leerConfig_())).join(', '); }
    else if (String(v).trim() === '') { v = ''; }   // en blanco = sin definir (sin avisos ni comparaciones)
    else { v = Number(String(v).replace(',', '.')); if (isNaN(v) || v < 0) throw new Error('Valor inválido en ' + k); }
    const i = claves.indexOf(k);
    if (i >= 0) sh.getRange(i + 2, 2).setValue(v);
    else sh.appendRow([k, v, '']);
  });
  return { ok: true };
}
function destinos_(cfg) {
  return String(cfg.DESTINOS || 'Autovías, Bacheo').split(',').map(s => s.trim()).filter(Boolean);
}
/** Otras obras (destinos variables): CONFIG OBRAS más cualquier destino ya cargado en SALIDAS que no sea fijo. */
function obras_(cfg, salidas) {
  const fijos = destinos_(cfg);
  const lista = String(cfg.OBRAS || '').split(',');
  (salidas || []).forEach(s => { if (s.t > 0) lista.push(s.destino); });
  return limpiarObras_(lista, fijos);
}
/** Normaliza nombres de obra: recorta, saca duplicados (sin mayúsculas ni tildes) y los que ya son destinos fijos. */
function limpiarObras_(lista, fijos) {
  const out = [], vistos = (fijos || []).map(normDestino_);
  lista.forEach(x => {
    const n = String(x || '').replace(/,/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
    if (!n) return;
    const k = normDestino_(n);
    if (vistos.indexOf(k) >= 0) return;
    vistos.push(k); out.push(n);
  });
  return out.sort((a, b) => a.localeCompare(b, 'es'));
}
/** Agrega a CONFIG OBRAS las obras nuevas que vinieron en una carga (una sola vez cada una). */
function registrarObras_(nuevas) {
  if (!nuevas.length) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.CONFIG);
  const claves = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues().map(r => String(r[0]).trim());
  const i = claves.indexOf('OBRAS');
  const actuales = i >= 0 ? String(sh.getRange(i + 2, 2).getValue() || '').split(',') : [];
  const lista = limpiarObras_(actuales.concat(nuevas), destinos_(leerConfig_()));
  if (i >= 0) sh.getRange(i + 2, 2).setValue(lista.join(', '));
  else sh.appendRow(['OBRAS', lista.join(', '), 'Otras obras a las que salió mezcla (separadas por coma). Se agregan solas desde la app al cargar una obra nueva; acá se pueden borrar o renombrar']);
}

// ---------- Utilidades ----------
function fechaStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s); return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
function aNumero_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(String(v).replace(',', '.'));
  return isNaN(n) ? null : n;
}
function hoy_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function sumarDias_(f, n) { const d = new Date(f + 'T12:00:00'); d.setDate(d.getDate() + n); return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }

// ---------- Lectura de datos ----------
function leerDiario_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.DIARIO);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().forEach((r, i) => {
    const f = fechaStr_(r[0]); if (!f) return;
    out[f] = { fecha: f, stock: aNumero_(r[1]), recibido: aNumero_(r[2]) || 0, obs: String(r[3] || ''), por: String(r[4] || ''), arranques: aNumero_(r[6]), fila: i + 2 };
  });
  return out;
}
function leerSalidas_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.SALIDAS);
  const out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach((r, i) => {
    const f = fechaStr_(r[0]); if (!f) return;
    out.push({ fecha: f, destino: String(r[1] || '').trim(), t: aNumero_(r[2]) || 0, por: String(r[3] || ''), fila: i + 2 });
  });
  return out;
}

/**
 * Serie diaria completa desde el primer dato hasta hoy.
 * Por día: mezcla total y por destino, asfalto recibido, stock medido (o null), consumo de asfalto
 * (repartido dentro del tramo entre mediciones) y % real.
 */
function getDatos(user) {
  const cfg = leerConfig_();
  const destinos = destinos_(cfg);
  const diario = leerDiario_();
  const salidas = leerSalidas_();
  const obras = obras_(cfg, salidas);   // destinos variables (otras obras), para el desplegable de la app
  // Parámetros en blanco = sin definir: no se compara ni se estima nada con ellos.
  const porcTeo = aNumero_(cfg.PORC_TEORICO), tol = aNumero_(cfg.TOLERANCIA) || 0;
  const stockMin = aNumero_(cfg.STOCK_MINIMO), capacidad = aNumero_(cfg.CAPACIDAD_TANQUE);
  const hoy = hoy_();

  const fechas = Object.keys(diario).concat(salidas.map(s => s.fecha)).sort();
  const inicio = fechas.length ? fechas[0] : hoy;
  const fin = fechas.length && fechas[fechas.length - 1] > hoy ? fechas[fechas.length - 1] : hoy;

  // mezcla por día y destino
  const mezcla = {};
  salidas.forEach(s => {
    const m = mezcla[s.fecha] || (mezcla[s.fecha] = { total: 0, por: {}, espejo: [] });
    m.total += s.t; m.por[s.destino] = (m.por[s.destino] || 0) + s.t;
    if (esEspejo_(s) && s.t > 0 && m.espejo.indexOf(s.destino) < 0) m.espejo.push(s.destino);   // la cargó Obra Autovías
  });

  const calendario = [];
  for (let f = inicio; f <= fin; f = sumarDias_(f, 1)) calendario.push(f);

  // Arranques de planta por día (en blanco = 1 si salió mezcla) y mezcla equivalente con la purga
  const arranques = f => {
    const a = diario[f] ? diario[f].arranques : null;
    return a !== null && a !== undefined ? a : (mezcla[f] && mezcla[f].total > 0 ? 1 : 0);
  };
  const mezEq = f => (mezcla[f] ? mezcla[f].total : 0) + arranques(f) * PURGA_EQ_T;

  // Tramos entre mediciones de stock: consumo = stock_ant + recibido(tramo) − stock_act
  const medidos = calendario.filter(f => diario[f] && diario[f].stock !== null);
  const consumoDia = {}, tramoDe = {};
  const tramos = [];
  for (let i = 1; i < medidos.length; i++) {
    const a = medidos[i - 1], b = medidos[i];
    const dias = calendario.filter(f => f > a && f <= b);
    const recibido = dias.reduce((s, f) => s + (diario[f] ? diario[f].recibido : 0), 0);
    const mezReal = dias.reduce((s, f) => s + (mezcla[f] ? mezcla[f].total : 0), 0);
    const arr = dias.reduce((s, f) => s + arranques(f), 0);
    const mez = dias.reduce((s, f) => s + mezEq(f), 0);
    const consumo = diario[a].stock + recibido - diario[b].stock;
    const porc = mez > 0 ? consumo / mez * 100 : null;
    const tr = { desde: a, hasta: b, dias: dias.length, recibido, mezcla: mezReal, arranques: arr, purga: arr * PURGA_T, mezclaEq: mez, consumo, porc, raro: consumo < 0 || (mez === 0 && Math.abs(consumo) > 0.05) };
    tramos.push(tr);
    dias.forEach(f => {
      const m = mezEq(f);
      consumoDia[f] = mez > 0 ? consumo * m / mez : (f === b ? consumo : 0);
      tramoDe[f] = tr;
    });
  }

  // Stock estimado para los días sin medición: última medición + recibido − mezcla × % teórico
  let ultimoStock = null, ultimoStockFecha = null;
  const acum = { mezcla: 0, consumo: 0, mezclaConConsumo: 0, recibido: 0, porDestino: {}, diasConMezcla: 0, arranques: 0 };
  destinos.concat(obras).forEach(d => acum.porDestino[d] = 0);
  let stockEst = null;
  const serie = calendario.map(f => {
    const d = diario[f] || { stock: null, recibido: 0, obs: '', por: '' };
    const m = mezcla[f] || { total: 0, por: {} };
    const medido = d.stock !== null;
    const arr = arranques(f), eq = mezEq(f);
    if (stockEst !== null) stockEst = porcTeo === null ? null : stockEst + d.recibido - eq * porcTeo / 100;
    if (medido) { ultimoStock = d.stock; ultimoStockFecha = f; stockEst = d.stock; }
    const consumo = consumoDia[f] !== undefined ? consumoDia[f] : null;
    const tr = tramoDe[f] || null;
    acum.mezcla += m.total; acum.recibido += d.recibido; acum.arranques += arr; if (m.total > 0) acum.diasConMezcla++;
    Object.keys(m.por).forEach(k => acum.porDestino[k] = (acum.porDestino[k] || 0) + m.por[k]);
    if (consumo !== null) { acum.consumo += consumo; acum.mezclaConConsumo += eq; }
    const porc = tr && tr.porc !== null ? tr.porc : null;
    return {
      fecha: f, mezcla: m.total, porDestino: m.por, espejo: m.espejo || [], recibido: d.recibido, obs: d.obs, por: d.por,
      arranques: arr, arranquesCargados: d.arranques !== null && d.arranques !== undefined ? d.arranques : null, purga: arr * PURGA_T, mezclaEq: eq,
      stock: medido ? d.stock : null, stockEst: medido ? null : (stockEst !== null ? Math.max(0, stockEst) : null),
      consumo, porc, tramoDias: tr ? tr.dias : null, tramoRaro: tr ? tr.raro : false,
      cargado: !!diario[f] || m.total > 0,
      desvio: porc !== null && porcTeo !== null ? (Math.abs(porc - porcTeo) > tol) : false,
    };
  });

  // Consumo promedio diario (últimos 7 días con mezcla y consumo conocido) → autonomía
  const conConsumo = serie.filter(x => x.consumo !== null && x.mezcla > 0).slice(-7);
  const consumoProm = conConsumo.length ? conConsumo.reduce((s, x) => s + x.consumo, 0) / conConsumo.length : null;
  const mezclaProm = acum.diasConMezcla ? acum.mezcla / acum.diasConMezcla : null;
  const porcAcum = acum.mezclaConConsumo > 0 ? acum.consumo / acum.mezclaConConsumo * 100 : null;
  const ultimoTramo = tramos.length ? tramos[tramos.length - 1] : null;
  const diasSinMedir = ultimoStockFecha ? Math.round((new Date(hoy + 'T12:00:00') - new Date(ultimoStockFecha + 'T12:00:00')) / 86400000) : null;

  const avisos = [];
  if (ultimoStock !== null && stockMin !== null && stockMin > 0 && ultimoStock < stockMin) avisos.push({ tipo: 'mal', msg: 'Stock de asfalto bajo: ' + fmtT_(ultimoStock) + ' t en tanque (mínimo ' + fmtT_(stockMin) + ' t).' });
  if (ultimoTramo && ultimoTramo.porc !== null && porcTeo !== null && Math.abs(ultimoTramo.porc - porcTeo) > tol) avisos.push({ tipo: 'warn', msg: 'Último % de asfalto ' + fmtP_(ultimoTramo.porc) + ' (teórico ' + fmtP_(porcTeo) + '): revisar la dosificación o la medición del tanque.' });
  if (ultimoTramo && ultimoTramo.raro) avisos.push({ tipo: 'warn', msg: 'El último tramo da un consumo raro (' + fmtT_(ultimoTramo.consumo) + ' t con ' + fmtT_(ultimoTramo.mezcla) + ' t de mezcla): ¿faltó cargar un ingreso de asfalto o una salida?' });
  if (diasSinMedir !== null && diasSinMedir >= 2) avisos.push({ tipo: 'warn', msg: 'Hace ' + diasSinMedir + ' días que no se carga el stock del tanque (último: ' + fmtFechaAR_(ultimoStockFecha) + ').' });

  return {
    user, hoy, inicio, destinos, obras,
    cfg: { PURGA_T, PURGA_EQ_T, PORC_TEORICO: porcTeo, TOLERANCIA: tol, STOCK_MINIMO: stockMin, CAPACIDAD_TANQUE: capacidad, DESTINO_AUTOVIAS: String(cfg.DESTINO_AUTOVIAS || destinos[0] || '') },
    kpi: {
      stock: ultimoStock, stockFecha: ultimoStockFecha, diasSinMedir,
      consumoProm, mezclaProm, porcAcum,
      autonomia: ultimoStock !== null && consumoProm > 0 ? ultimoStock / consumoProm : null,
      ultimoTramo,
      mezclaTotal: acum.mezcla, consumoTotal: acum.consumo, mezclaConConsumo: acum.mezclaConConsumo, recibidoTotal: acum.recibido, arranquesTotal: acum.arranques, purgaTotal: acum.arranques * PURGA_T, diasConMezcla: acum.diasConMezcla, porDestino: acum.porDestino,
    },
    serie, tramos, avisos,
  };
}
const fmtT_ = v => (Math.round(v * 10) / 10).toLocaleString('es-AR');
const fmtP_ = v => (Math.round(v * 100) / 100).toLocaleString('es-AR') + ' %';
const fmtFechaAR_ = f => f.slice(8, 10) + '/' + f.slice(5, 7);

// ---------- Carga ----------
/** Cierre del día: { fecha, stock (t, vacío = no se midió), recibido (t), salidas: { destino: t }, obs } */
function guardarDia(p, user) {
  if (!user) throw new Error('Sin permiso');
  const fecha = String(p.fecha || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Fecha inválida');
  if (fecha > hoy_()) throw new Error('No se puede cargar una fecha futura');
  const cfg = leerConfig_();
  const capacidad = aNumero_(cfg.CAPACIDAD_TANQUE) || 0;
  const stock = aNumero_(p.stock);
  if (stock !== null && (stock < 0 || stock > 10000)) throw new Error('Stock fuera de rango');
  if (stock !== null && capacidad > 0 && stock > capacidad * 1.05) throw new Error('El stock (' + stock + ' t) supera la capacidad del tanque (' + capacidad + ' t)');
  const recibido = aNumero_(p.recibido) || 0;
  if (recibido < 0 || recibido > 10000) throw new Error('Asfalto recibido fuera de rango');
  // Destinos: los fijos de CONFIG (Autovías, Bacheo…) más otras obras, que se escriben libres y quedan
  // registradas en CONFIG OBRAS para el desplegable. Si una obra ya existe con otra mayúscula/tilde se usa el nombre guardado.
  const destinos = destinos_(cfg);
  const conocidas = obras_(cfg, leerSalidas_());
  const salidas = {}, obrasNuevas = [];
  Object.keys(p.salidas || {}).forEach(d0 => {
    const t = aNumero_(p.salidas[d0]) || 0;
    if (t < 0 || t > 5000) throw new Error('Mezcla fuera de rango para ' + d0);
    let d = String(d0 || '').replace(/,/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
    if (!d) { if (t > 0) throw new Error('Falta el nombre de la obra en una salida de ' + t + ' t'); return; }
    const n = normDestino_(d);
    const fijo = destinos.find(x => normDestino_(x) === n);
    if (fijo) d = fijo;
    else {
      const conocida = conocidas.find(x => normDestino_(x) === n);
      if (conocida) d = conocida; else if (t > 0) obrasNuevas.push(d);
    }
    salidas[d] = (salidas[d] || 0) + t;
  });
  const totalMezcla = Object.values(salidas).reduce((s, t) => s + t, 0);
  // Arranques: en blanco = automático (1 si salió mezcla)
  let arranques = '';
  if (p.arranques !== undefined && p.arranques !== null && String(p.arranques).trim() !== '') {
    arranques = Number(p.arranques);
    if (!Number.isInteger(arranques) || arranques < 0 || arranques > 10) throw new Error('Arranques de planta inválidos (0 a 10)');
  }
  if (stock === null && recibido === 0 && totalMezcla === 0) throw new Error('No hay nada para guardar: cargá el stock del tanque o alguna salida');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  let diario;
  try {
    // DIARIO: una fila por fecha
    const shD = ss.getSheetByName(TABS.DIARIO);
    diario = leerDiario_();
    if (shD.getRange(1, 7).getValue() === '') shD.getRange(1, 7).setValue('ARRANQUES').setFontWeight('bold');
    const rowD = [new Date(fecha + 'T12:00:00'), stock === null ? '' : stock, recibido, String(p.obs || '').slice(0, 300), user, new Date(), arranques];
    if (diario[fecha]) shD.getRange(diario[fecha].fila, 1, 1, 7).setValues([rowD]); else shD.appendRow(rowD);
    // SALIDAS: una fila por fecha y destino; las del día que no vienen o vienen en 0 se borran
    const shS = ss.getSheetByName(TABS.SALIDAS);
    const existentes = leerSalidas_().filter(s => s.fecha === fecha);
    const pendientes = Object.assign({}, salidas);
    existentes.sort((a, b) => b.fila - a.fila).forEach(s => {
      if (pendientes[s.destino] !== undefined && pendientes[s.destino] > 0) {
        shS.getRange(s.fila, 1, 1, 5).setValues([[new Date(fecha + 'T12:00:00'), s.destino, pendientes[s.destino], user, new Date()]]);
        delete pendientes[s.destino];
      } else shS.deleteRow(s.fila);
    });
    Object.keys(pendientes).forEach(d => { if (pendientes[d] > 0) shS.appendRow([new Date(fecha + 'T12:00:00'), d, pendientes[d], user, new Date()]); });
    registrarObras_(obrasNuevas);
  } finally { lock.releaseLock(); }
  // Un solo getDatos: lo usa el aviso y se devuelve como `all` (la app lo pide con con_all).
  // Antes se calculaba dos veces y, sumado al aviso, la respuesta podía pasar el tiempo de espera del celular.
  SpreadsheetApp.flush();
  let all = null, aviso = null;
  try { all = getDatos(user); } catch (err) { aviso = { ok: false, error: 'datos: ' + String(err.message || err) }; }
  if (all) { try { aviso = avisarCierre_(fecha, user, all); } catch (err) { aviso = { ok: false, error: String(err.message || err) }; } }
  return { ok: true, editado: !!diario[fecha], aviso, all };
}

function borrarDia(fecha, user) {
  if (!user) throw new Error('Sin permiso');
  fecha = String(fecha || '').trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = leerDiario_()[fecha];
    if (d) ss.getSheetByName(TABS.DIARIO).deleteRow(d.fila);
    const shS = ss.getSheetByName(TABS.SALIDAS);
    leerSalidas_().filter(s => s.fecha === fecha).sort((a, b) => b.fila - a.fila).forEach(s => shS.deleteRow(s.fila));
  } finally { lock.releaseLock(); }
  return { ok: true };
}

// ---------- Aviso por la app INGECO al cerrar el día ----------
function avisarCierre_(fecha, user, datosYaCalculados) {
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE_' + MODULO_INGECO);
  if (!clave) return { ok: false, error: 'Sin CLAVE_' + MODULO_INGECO };
  const datos = datosYaCalculados || getDatos(user);
  const d = datos.serie.find(x => x.fecha === fecha);
  if (!d) return { ok: false, error: 'día no encontrado' };
  const k = datos.kpi;
  const destinos = Object.keys(d.porDestino).filter(x => d.porDestino[x] > 0).map(x => x + ' ' + fmtT_(d.porDestino[x]) + ' t').join(' · ');
  const stockTxt = d.stock !== null ? 'stock ' + fmtT_(d.stock) + ' t' : 'sin medición de tanque';
  const titulo = 'Planta ' + fmtFechaAR_(fecha) + ': ' + fmtT_(d.mezcla) + ' t de mezcla · ' + stockTxt;
  const lineas = [
    'MEZCLA: ' + fmtT_(d.mezcla) + ' t' + (destinos ? ' (' + destinos + ')' : '') + (d.arranques ? ' + purga ' + fmtT_(d.purga) + ' t (' + d.arranques + (d.arranques === 1 ? ' arranque)' : ' arranques)') : ''),
    'ASFALTO: ' + (d.consumo !== null ? 'consumo ' + fmtT_(d.consumo) + ' t' + (d.tramoDias > 1 ? ' (tramo de ' + d.tramoDias + ' días)' : '') + ' → ' + fmtP_(d.porc) + (datos.cfg.PORC_TEORICO !== null ? ' real vs ' + fmtP_(datos.cfg.PORC_TEORICO) + ' teórico' : ' real') : 'consumo pendiente hasta la próxima medición del tanque') + (d.recibido ? ' · recibido ' + fmtT_(d.recibido) + ' t' : ''),
    'TANQUE: ' + (d.stock !== null ? fmtT_(d.stock) + ' t' + (k.autonomia !== null ? ' ≈ ' + fmtT_(k.autonomia) + ' días al ritmo actual' : '') : 'no se midió'),
  ];
  datos.avisos.forEach(a => lineas.push('⚠ ' + a.msg));
  lineas.push('Cargó: ' + user);
  const critica = datos.avisos.some(a => a.tipo === 'mal');
  const r = UrlFetchApp.fetch(APP_INGECO_BACKEND, {
    method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true,
    payload: JSON.stringify({
      accion: 'notificar', clave_sistema: clave, modulo: MODULO_INGECO,
      destinatarios: [{ con_modulo: MODULO_INGECO }, { admin: true }],
      tipo: 'aviso_general', titulo: titulo.slice(0, 120), cuerpo: lineas.join('\n').slice(0, 600),
      url_destino: URL_APP, prioridad: critica ? 'critica' : 'normal',
      clave_unica: 'planta-' + fecha + '-' + Math.round(d.mezcla * 10) + '-' + (d.stock === null ? 'x' : Math.round(d.stock * 10)) + '-' + Math.round(d.recibido * 10), dias_unica: 1
    })
  });
  let j = {}; try { j = JSON.parse(r.getContentText()); } catch (e) { }
  return j.ok ? { ok: true, destinatarios: (j.avisos || []).length } : { ok: false, error: r.getContentText().slice(0, 160) };
}

// ---------- Integración: mezcla por día para otro sistema (Obra Autovías) ----------
// POST { accion:'mezcla', clave, payload:{ destino?, desde? } } → { destino, dias:[{fecha, t}], destinos:[...], aviso }
// El destino se compara sin mayúsculas, tildes ni espacios de más ("Autovias" = "AUTOVÍAS " = "Autovías"),
// así un retoque del nombre en CONFIG o en la planilla no deja a Autovías sin mezcla.
function normDestino_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}
function mezclaParaIntegracion_(body, p) {
  const cfg = leerConfig_();
  const esperada = String(cfg.CLAVE_API || '').trim();
  if (!esperada || String(body.clave || '').trim() !== esperada) throw new Error('Clave de integración inválida');
  const pedido = String(p.destino || cfg.DESTINO_AUTOVIAS || '').trim();
  const desde = fechaStr_(p.desde) || '2000-01-01';
  const salidas = leerSalidas_();
  const enPlanilla = []; salidas.forEach(s => { if (s.destino && enPlanilla.indexOf(s.destino) < 0) enPlanilla.push(s.destino); });
  let destino = pedido, aviso = '';
  if (pedido && enPlanilla.indexOf(pedido) < 0) {
    const n = normDestino_(pedido);
    const raiz = n.split(' ')[0].slice(0, 5);   // "autov" alcanza para Autovía / Autovías / Autovía del NOA
    const parecido = enPlanilla.find(d => normDestino_(d) === n)
      || enPlanilla.find(d => normDestino_(d).indexOf(n) >= 0 || n.indexOf(normDestino_(d)) >= 0)
      || (raiz.length >= 4 ? enPlanilla.find(d => normDestino_(d).indexOf(raiz) === 0) : null);
    if (parecido) { destino = parecido; aviso = 'El destino "' + pedido + '" no existe tal cual en la planilla; se usa "' + parecido + '".'; }
    else aviso = 'Ningún despacho de mezcla tiene el destino "' + pedido + '". Destinos con mezcla cargada: ' + (enPlanilla.join(', ') || 'ninguno') + '. Revisar DESTINO_AUTOVIAS en la configuración de Stock de asfalto.';
  }
  const nd = normDestino_(destino), porDia = {}, deLaPlanta = {};
  salidas.forEach(s => {
    if (s.fecha < desde || (destino && normDestino_(s.destino) !== nd)) return;
    porDia[s.fecha] = (porDia[s.fecha] || 0) + s.t;
    if (!esEspejo_(s)) deLaPlanta[s.fecha] = true;
  });
  // espejo: true = esa mezcla la cargó Autovías, no el plantista (Autovías no la toma como "informada por la planta")
  return { destino, aviso, destinos: enPlanilla, dias: Object.keys(porDia).sort().map(f => ({ fecha: f, t: Math.round(porDia[f] * 100) / 100, espejo: !deLaPlanta[f] })) };
}

// ---------- Integración: salidas que cargó Obra Autovías y la planta no ----------
// POST { accion:'salidaExterna', clave, payload:{ dias:[{fecha, t, por?}], por } }
// Por cada día: si la planta ya cargó mezcla para Autovías ese día, manda lo de la planta y no se toca.
// Si no, se escribe (o actualiza) una fila en SALIDAS con CARGADO POR = "Autovías · <quién>"; con t = 0 se borra.
// Las filas espejo dejan de serlo apenas el plantista edita el día desde su app (guardarDia reescribe CARGADO POR).
// Fechas anteriores al arranque de Stock de asfalto se ignoran: antes de eso la planta no llevaba registro.
const ESPEJO_PREFIJO = 'Autovías · ';
function esEspejo_(s) { return String(s.por || '').indexOf(ESPEJO_PREFIJO) === 0; }

function salidaExterna_(body, p) {
  const cfg = leerConfig_();
  const esperada = String(cfg.CLAVE_API || '').trim();
  if (!esperada || String(body.clave || '').trim() !== esperada) throw new Error('Clave de integración inválida');
  const destino = String(cfg.DESTINO_AUTOVIAS || destinos_(cfg)[0] || 'Autovías').trim();
  const nd = normDestino_(destino);
  const por = ESPEJO_PREFIJO + (String(p.por || '').trim().slice(0, 50) || 'Obra Autovías');
  const dias = Array.isArray(p.dias) ? p.dias.slice(0, 400) : [];
  const hoy = hoy_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shS = ss.getSheetByName(TABS.SALIDAS);
  const out = [];
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    // Arranque de Stock: primer día con carga propia de la planta (tanque o salidas que no son espejo)
    const propias = Object.keys(leerDiario_()).concat(leerSalidas_().filter(s => !esEspejo_(s)).map(s => s.fecha)).sort();
    const inicio = propias[0] || hoy;
    dias.forEach(d => {
      const fecha = fechaStr_(d && d.fecha);
      const t = Math.round((aNumero_(d && d.t) || 0) * 100) / 100;
      if (!fecha) return out.push({ fecha: String(d && d.fecha), r: 'fecha inválida' });
      if (fecha < inicio) return out.push({ fecha, r: 'antes de Stock (' + inicio + ')' });
      if (fecha > hoy) return out.push({ fecha, r: 'fecha futura' });
      if (t < 0 || t > 5000) return out.push({ fecha, r: 't fuera de rango' });
      const del = leerSalidas_().filter(s => s.fecha === fecha && normDestino_(s.destino) === nd);
      if (del.some(s => !esEspejo_(s))) return out.push({ fecha, r: 'ya cargada por la planta' });
      const filas = del.map(s => s.fila).sort((a, b) => b - a);
      if (t > 0) {
        const quien = d.por ? ESPEJO_PREFIJO + String(d.por).trim().slice(0, 50) : por;   // quién lo cargó en Autovías
        const row = [new Date(fecha + 'T12:00:00'), destino, t, quien, new Date()];
        if (filas.length) {
          if (del.length === 1 && Math.abs(del[0].t - t) < 0.005) return out.push({ fecha, r: 'sin cambios' });
          shS.getRange(filas.pop(), 1, 1, 5).setValues([row]);
          filas.forEach(f => shS.deleteRow(f));   // de mayor a menor: no corre las filas pendientes
          out.push({ fecha, r: 'actualizada', t });
        } else { shS.appendRow(row); out.push({ fecha, r: 'agregada', t }); }
      } else if (filas.length) {
        filas.forEach(f => shS.deleteRow(f));
        out.push({ fecha, r: 'borrada' });
      } else out.push({ fecha, r: 'nada que hacer' });
    });
  } finally { lock.releaseLock(); }
  return { destino, dias: out };
}
