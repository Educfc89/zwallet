/**
 * zwallet · puente con el Google Sheet "Balance Mensual"
 * Pega este archivo en Extensiones → Apps Script del Sheet, ejecuta "preparar" una vez
 * y publica como Aplicación web (Ejecutar como: yo · Acceso: cualquier persona).
 * La seguridad la da el TOKEN de la hoja zwallet_Config (la app lo manda en cada llamada).
 */
var HOJA_CFG = 'zwallet_Config', HOJA_DIAG = 'zwallet_Diagnóstico', HOJA_AYUDA = 'zwallet_Instrucciones';
var MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

/* ---------------- menú y preparación ---------------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('zwallet')
    .addItem('Preparar hojas (Config, Diagnóstico, Instrucciones)', 'preparar')
    .addItem('Probar conexión', 'probar')
    .addItem('Generar token nuevo', 'tokenNuevo')
    .addItem('Ver diagnóstico', 'verDiag')
    .addToUi();
}
function preparar() {
  var ss = SpreadsheetApp.getActive();
  var cfg = ss.getSheetByName(HOJA_CFG);
  if (!cfg) { cfg = ss.insertSheet(HOJA_CFG); escribirConfig_(cfg, null, { token: nuevoToken_(), hoja: 'auto', prefijo: 'zwallet' }); }
  if (!ss.getSheetByName(HOJA_DIAG)) {
    var d = ss.insertSheet(HOJA_DIAG);
    d.getRange(1, 1, 1, 4).setValues([['Fecha', 'Acción', 'Resultado', 'Detalle']]).setFontWeight('bold').setBackground('#0B0F0E').setFontColor('#ffffff');
    d.setFrozenRows(1); d.setColumnWidths(1, 1, 150); d.setColumnWidths(2, 1, 120); d.setColumnWidths(3, 1, 90); d.setColumnWidths(4, 1, 520);
  }
  if (!ss.getSheetByName(HOJA_AYUDA)) ayuda_(ss.insertSheet(HOJA_AYUDA));
  diag_('preparar', 'ok', 'Hojas listas');
  SpreadsheetApp.getUi().alert('Listo. Ahora: Implementar → Nueva implementación → Aplicación web (Ejecutar como: yo, Acceso: cualquier persona). Copia la dirección /exec y pégala en la app, Ajustes, junto con el token de la hoja zwallet_Config.');
}
function tokenNuevo() {
  var sh = SpreadsheetApp.getActive().getSheetByName(HOJA_CFG); if (!sh) return preparar();
  var c = leerConfig_(sh); c.conn.token = nuevoToken_(); escribirConfig_(sh, c, c.conn);
  diag_('token', 'ok', 'Token regenerado: actualízalo en la app'); SpreadsheetApp.getUi().alert('Token nuevo generado en zwallet_Config. Pégalo otra vez en la app (Ajustes).');
}
function verDiag() { var s = SpreadsheetApp.getActive().getSheetByName(HOJA_DIAG); if (s) SpreadsheetApp.setActiveSheet(s); }
function probar() {
  var out = [], ss = SpreadsheetApp.getActive();
  out.push((ss.getSheetByName(HOJA_CFG) ? 'OK' : 'FALTA') + ' hoja ' + HOJA_CFG);
  out.push((ss.getSheetByName(HOJA_DIAG) ? 'OK' : 'FALTA') + ' hoja ' + HOJA_DIAG);
  var t = tabMes_(ss, 'auto');
  out.push(t ? 'OK hoja de gastos detectada: ' + t.getName() : 'FALTA una hoja con nombre tipo "noviembre 2026"');
  out.push('Web app: ' + (ScriptApp.getService().getUrl() || 'sin publicar todavía'));
  diag_('probar', 'ok', out.join(' | ')); SpreadsheetApp.getUi().alert(out.join('\n'));
}
function nuevoToken_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 20); }
function diag_(acc, res, det) {
  try {
    var s = SpreadsheetApp.getActive().getSheetByName(HOJA_DIAG); if (!s) return;
    s.insertRowAfter(1); s.getRange(2, 1, 1, 4).setValues([[new Date(), acc, res, String(det).slice(0, 500)]]);
    s.getRange(2, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    if (s.getLastRow() > 500) s.deleteRows(502, s.getLastRow() - 501);
  } catch (e) {}
}

/* ---------------- web app ---------------- */
function doGet(e) { return resp_(manejar_((e && e.parameter) || {})); }
function doPost(e) {
  var p = {}; try { p = JSON.parse(e.postData.contents); } catch (x) { return resp_({ ok: false, err: 'json' }); }
  return resp_(manejar_(p));
}
function resp_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function manejar_(p) {
  var ss = SpreadsheetApp.getActive(), a = p.a;
  var cfgSh = ss.getSheetByName(HOJA_CFG);
  if (!cfgSh) return { ok: false, err: 'sin_config' };
  var cfg = leerConfig_(cfgSh);
  if (!cfg.conn.token || String(p.t) !== String(cfg.conn.token)) { diag_(a || '?', 'auth', 'Token incorrecto'); return { ok: false, err: 'auth' }; }
  try {
    if (a === 'ping') { diag_('ping', 'ok', ''); return { ok: true, sheet: ss.getName(), url: ss.getUrl() }; }
    if (a === 'read') {
      var tab = tabMes_(ss, cfg.conn.hoja); if (!tab) { diag_('read', 'error', 'No hay hoja de mes'); return { ok: false, err: 'sin_hoja' }; }
      var vals = tab.getDataRange().getDisplayValues();
      var txt = '### Sheet Name: ' + tab.getName() + '\n' + vals.map(function (r) { return '| ' + r.join(' | ') + ' |'; }).join('\n') + '\n';
      diag_('read', 'ok', tab.getName() + ' · ' + vals.length + ' filas'); return { ok: true, name: tab.getName(), text: txt };
    }
    if (a === 'notify') {
      var to = (p.to || []).filter(function (m) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m); }).slice(0, 5);
      if (!to.length) { diag_('notify', 'error', 'Sin destinatarios'); return { ok: false, err: 'sin_destino' }; }
      MailApp.sendEmail({ to: to.join(','), subject: String(p.subject || 'zwallet').slice(0, 150), body: String(p.body || '').slice(0, 3000), name: 'zwallet' });
      diag_('notify', 'ok', String(p.subject || '').slice(0, 120) + ' → ' + to.join(', ')); return { ok: true };
    }
    if (a === 'write') return escribirMes_(ss, cfg, p);
    if (a === 'config') { diag_('config', 'ok', 'Configuración leída'); return { ok: true, config: cfg.app, url: ss.getUrl() }; }
    if (a === 'setconfig') { escribirConfig_(cfgSh, { app: p.config }, cfg.conn); diag_('setconfig', 'ok', 'Configuración guardada desde la app'); return { ok: true }; }
    return { ok: false, err: 'accion' };
  } catch (err) { diag_(a, 'error', err); return { ok: false, err: 'script', msg: String(err) }; }
}

/* ---------------- hoja de gastos ---------------- */
function claveMes_(nombre) {
  var m = /^([A-Za-zñÑáéíóú]+)\s+(\d{4})/.exec(String(nombre).trim()); if (!m) return null;
  var i = MESES.indexOf(m[1].toLowerCase()); if (i < 0) return null; return m[2] + '-' + ('0' + (i + 1)).slice(-2);
}
function tabMes_(ss, hoja) {
  var shs = ss.getSheets();
  if (hoja && hoja !== 'auto') { var h = ss.getSheetByName(hoja); if (h) return h; }
  var best = null, bk = '';
  shs.forEach(function (s) { var k = claveMes_(s.getName()); if (k && k > bk) { bk = k; best = s; } });
  return best;
}
function escribirMes_(ss, cfg, p) {
  var nombre = (cfg.conn.prefijo ? cfg.conn.prefijo + ' ' : '') + String(p.name || 'mes').replace(/[\[\]\*\?\/\\:]/g, '').slice(0, 60);
  var rows = Utilities.parseCsv(String(p.csv || '').replace(/^﻿/, ''));
  if (rows.length < 2) return { ok: false, err: 'vacio' };
  var sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre); sh.clear();
  var NUM = { 3: 1, 6: 1, 7: 1, 8: 1, 10: 1 }, out = rows.map(function (r, i) {
    return r.map(function (v, c) {
      if (i === 0 || v === '') return v;
      if (NUM[c]) { var n = parseFloat(String(v).replace(/\./g, '').replace(',', '.')); return isNaN(n) ? v : n; }
      var m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(v); if ((c === 4 || c === 9) && m) return new Date(+m[3], +m[2] - 1, +m[1]);
      return v;
    });
  });
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(1, 1, 1, out[0].length).setFontWeight('bold').setBackground('#0B0F0E').setFontColor('#ffffff'); sh.setFrozenRows(1);
  sh.getRange(2, 4, out.length - 1, 1).setNumberFormat('#,##0.00'); sh.getRange(2, 7, out.length - 1, 3).setNumberFormat('#,##0.00');
  sh.getRange(2, 5, out.length - 1, 1).setNumberFormat('dd/mm/yyyy'); sh.getRange(2, 10, out.length - 1, 1).setNumberFormat('dd/mm/yyyy');
  sh.autoResizeColumns(1, out[0].length);
  diag_('write', 'ok', nombre + ' · ' + (out.length - 1) + ' gastos');
  return { ok: true, name: nombre, url: ss.getUrl() + '#gid=' + sh.getSheetId() };
}

/* ---------------- configuración ---------------- */
var BLOQUES = ['CONEXIÓN', 'AJUSTES', 'TARJETAS', 'CATEGORÍAS'];
function leerConfig_(sh) {
  var v = sh.getDataRange().getDisplayValues(), conn = { token: '', hoja: 'auto', prefijo: 'zwallet' }, app = { settings: {}, cards: [], cats: [] }, blk = null, skip = 0;
  for (var i = 0; i < v.length; i++) {
    var a = String(v[i][0]).trim();
    if (BLOQUES.indexOf(a) >= 0) { blk = a; skip = 1; continue; }
    if (skip) { skip--; continue; }
    if (!blk || !a) { if (!a) blk = blk; continue; }
    if (blk === 'CONEXIÓN') { if (a === 'Token de acceso') conn.token = String(v[i][1]).trim(); else if (a === 'Hoja a leer') conn.hoja = String(v[i][1]).trim() || 'auto'; else if (a === 'Prefijo hoja de salida') conn.prefijo = String(v[i][1]).trim(); }
    else if (blk === 'AJUSTES') app.settings[a] = String(v[i][1]).trim();
    else if (blk === 'TARJETAS') app.cards.push({ id: a === '(nueva)' ? '' : a, name: v[i][1], kind: v[i][2], red: v[i][3], bank: v[i][4], closing: v[i][5], due: v[i][6], owner: v[i][7], on: !/^(no|n|0|false)$/i.test(String(v[i][8]).trim()) });
    else if (blk === 'CATEGORÍAS') app.cats.push({ id: a === '(nueva)' ? '' : a, name: v[i][1] });
  }
  return { conn: conn, app: app };
}
function escribirConfig_(sh, datos, conn) {
  sh.clear(); var r = [], fmt = [];
  function add(row, tipo) { r.push(row); fmt.push(tipo || ''); }
  add(['zwallet · Configuración', '', '', '', '', '', '', '', ''], 'T');
  add(['Edita aquí y en la app usa Ajustes → "Traer configuración del Sheet". No cambies los nombres de las columnas ni de los bloques.', '', '', '', '', '', '', '', '']);
  add(['', '', '', '', '', '', '', '', '']);
  add(['CONEXIÓN', '', '', '', '', '', '', '', ''], 'B'); add(['Parámetro', 'Valor', 'Nota', '', '', '', '', '', ''], 'H');
  add(['Token de acceso', conn.token, 'Clave que la app manda en cada llamada. Menú zwallet → Generar token nuevo.', '', '', '', '', '', '']);
  add(['Hoja a leer', conn.hoja || 'auto', '"auto" = la pestaña de mes más reciente (ej: noviembre 2026).', '', '', '', '', '', '']);
  add(['Prefijo hoja de salida', conn.prefijo || 'zwallet', 'Las hojas que crea la app se llaman "<prefijo> Noviembre 2026".', '', '', '', '', '', '']);
  add(['', '', '', '', '', '', '', '', '']);
  var app = datos && datos.app;
  if (app) {
    var S = app.settings || {};
    add(['AJUSTES', '', '', '', '', '', '', '', ''], 'B'); add(['Parámetro', 'Valor', 'Nota', '', '', '', '', '', ''], 'H');
    Object.keys(S).forEach(function (k) { add([k, S[k], '', '', '', '', '', '', '']); });
    add(['', '', '', '', '', '', '', '', '']);
    add(['TARJETAS', '', '', '', '', '', '', '', ''], 'B'); add(['ID', 'Nombre', 'Tipo', 'Red', 'Banco', 'Cierre', 'Vence', 'Titular', 'Activa'], 'H');
    (app.cards || []).forEach(function (c) { add([c.id, c.name, c.kind, c.red, c.bank, c.closing, c.due, c.owner, c.on ? 'Sí' : 'No']); });
    add(['', '', '', '', '', '', '', '', '']);
    add(['CATEGORÍAS', '', '', '', '', '', '', '', ''], 'B'); add(['ID', 'Nombre', '', '', '', '', '', '', ''], 'H');
    (app.cats || []).forEach(function (c) { add([c.id, c.name, '', '', '', '', '', '', '']); });
  } else {
    add(['AJUSTES', '', '', '', '', '', '', '', ''], 'B'); add(['Parámetro', 'Valor', 'Nota', '', '', '', '', '', ''], 'H');
    add(['(vacío)', 'En la app: Ajustes → "Enviar configuración al Sheet" llena estos bloques.', '', '', '', '', '', '', '']);
  }
  sh.getRange(1, 1, r.length, 9).setNumberFormat('@').setValues(r);
  fmt.forEach(function (f, i) {
    var g = sh.getRange(i + 1, 1, 1, 9);
    if (f === 'T') g.setFontSize(14).setFontWeight('bold');
    if (f === 'B') g.setFontWeight('bold').setBackground('#1F7A4D').setFontColor('#ffffff');
    if (f === 'H') g.setFontWeight('bold').setBackground('#E3E9E6');
  });
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 220); sh.setColumnWidth(3, 130);
}
function ayuda_(sh) {
  var t = [
    ['zwallet · Instrucciones de uso'], [''],
    ['1. Menú zwallet → "Preparar hojas": crea Config, Diagnóstico y esta hoja.'],
    ['2. Implementar → Nueva implementación → tipo Aplicación web → Ejecutar como: yo → Acceso: cualquier persona → Implementar. Copia la dirección que termina en /exec.'],
    ['3. En zwallet (web) → Ajustes → "Conexión con el Sheet": pega la dirección y el Token de acceso que está en zwallet_Config. Guardar.'],
    ['4. "Actualizar desde el Sheet": trae a la app los gastos nuevos, borrados y cambios de monto de la pestaña del mes más reciente.'],
    ['5. "Crear hoja de <mes>" en la app: crea o reemplaza una pestaña "zwallet <Mes Año>" aquí con los gastos de la app en formato Balance Mensual.'],
    ['6. Configuración: en la app "Enviar configuración al Sheet" llena zwallet_Config (ajustes, tarjetas, categorías). Edita las celdas y usa "Traer configuración del Sheet". Fila de tarjeta/categoría nueva: pon (nueva) en ID. Para borrar tarjetas o categorías usa la app.'],
    ['7. Si algo falla: mira la hoja zwallet_Diagnóstico (cada llamada deja una fila). Menú zwallet → "Probar conexión".'],
    ['8. Si cambias el código del script: Implementar → Administrar implementaciones → editar → Nueva versión (la dirección /exec no cambia).'],
    ['Seguridad: quien tenga la dirección y el token puede leer y escribir este Sheet. Si se filtra: menú zwallet → Generar token nuevo.']
  ];
  sh.getRange(1, 1, t.length, 1).setValues(t).setWrap(true); sh.setColumnWidth(1, 760); sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
}
