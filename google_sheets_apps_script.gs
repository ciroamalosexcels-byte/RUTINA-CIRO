/**
 * Rutina Ciro — Google Sheets API
 * Pegá este código en Extensiones → Apps Script dentro de un Google Sheet.
 * Después: Implementar → Nueva implementación → Aplicación web.
 * Ejecutar como: Yo. Acceso: Cualquier persona.
 */

const APP_NAME = 'Rutina Ciro';
const SHEETS = {
  dashboard: 'Dashboard',
  config: 'Config',
  rutina: 'Rutina',
  tiempoAreas: 'Tiempo_Areas',
  tiempoSesiones: 'Tiempo_Sesiones',
  pomodoro: 'Pomodoro',
  cumplimiento: 'Cumplimiento',
  clientes: 'Clientes',
  backup: 'BackupJSON',
  historial: 'HistorialSync'
};

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'ping';
    const callback = e && e.parameter && e.parameter.callback;
    let out;
    if (action === 'ping') out = handlePing_();
    else if (action === 'pullAll') out = handlePullAll_();
    else out = { ok: true, app: APP_NAME, message: 'API activa', time: new Date().toISOString() };
    return callback ? jsonpResponse(out, callback) : jsonResponse(out);
  } catch (err) {
    const out = { ok: false, error: String(err && err.stack ? err.stack : err) };
    const callback = e && e.parameter && e.parameter.callback;
    return callback ? jsonpResponse(out, callback) : jsonResponse(out);
  }
}

function doPost(e) {
  try {
    const body = parseBody_(e);
    const action = body.action || 'ping';
    if (action === 'ping') return jsonResponse(handlePing_());
    if (action === 'syncAll') return jsonResponse(handleSyncAll_(body.snapshot));
    if (action === 'pullAll') return jsonResponse(handlePullAll_());
    return jsonResponse({ ok: false, error: 'Acción desconocida: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.stack ? err.stack : err) });
  }
}

function parseBody_(e) {
  if (e && e.parameter && e.parameter.payload) return JSON.parse(e.parameter.payload);
  if (!e || !e.postData || !e.postData.contents) return {};
  return JSON.parse(e.postData.contents);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpResponse(obj, callback) {
  const cb = String(callback || '').replace(/[^A-Za-z0-9_.$]/g, '');
  if (!cb) return jsonResponse({ ok:false, error:'Callback inválido' });
  const json = JSON.stringify(obj).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return ContentService
    .createTextOutput(cb + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function handlePing_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllSheets_(ss);
  return { ok: true, spreadsheetName: ss.getName(), spreadsheetUrl: ss.getUrl(), time: new Date().toISOString() };
}

function handlePullAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.backup);
  if (!sh) return { ok: false, error: 'No existe la pestaña BackupJSON. Primero sincronizá desde la app.' };
  const lastRow = Math.max(2, sh.getLastRow());
  const rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const json = rows
    .filter(r => String(r[0]).indexOf('Backup') !== -1 || String(r[0]).indexOf('Último backup completo') !== -1)
    .map(r => r[1])
    .filter(String)
    .join('');
  if (!json) return { ok: false, error: 'BackupJSON está vacío. Primero sincronizá desde la app.' };
  return { ok: true, snapshot: JSON.parse(json), spreadsheetUrl: ss.getUrl() };
}

function handleSyncAll_(snapshot) {
  if (!snapshot) throw new Error('No llegó snapshot desde la app.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllSheets_(ss);
  writeConfig_(ss, snapshot);
  writeDashboard_(ss, snapshot);
  writeRutina_(ss, snapshot);
  writeTiempoAreas_(ss, snapshot);
  writeTiempoSesiones_(ss, snapshot);
  writePomodoro_(ss, snapshot);
  writeCumplimiento_(ss, snapshot);
  writeClientes_(ss, snapshot);
  writeBackup_(ss, snapshot);
  writeHistorial_(ss, snapshot);
  return { ok: true, spreadsheetName: ss.getName(), spreadsheetUrl: ss.getUrl(), updatedAt: new Date().toISOString() };
}

function ensureAllSheets_(ss) {
  Object.values(SHEETS).forEach(name => getOrCreate_(ss, name));
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function clearAndWrite_(sh, values, frozenRows) {
  sh.clear();
  if (!values || !values.length) values = [['Sin datos']];
  sh.getRange(1, 1, values.length, values[0].length).setValues(values);
  if (frozenRows) sh.setFrozenRows(frozenRows);
  styleSheet_(sh, values[0].length);
}

function styleSheet_(sh, cols) {
  const lastRow = Math.max(1, sh.getLastRow());
  const lastCol = Math.max(1, cols || sh.getLastColumn());
  sh.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#1A1814').setFontColor('#F5F2EC');
  sh.getRange(1, 1, lastRow, lastCol).setVerticalAlignment('middle').setWrap(true);
  try { sh.autoResizeColumns(1, lastCol); } catch(e) {}
}

function writeConfig_(ss, s) {
  const sh = ss.getSheetByName(SHEETS.config);
  clearAndWrite_(sh, [
    ['Campo', 'Valor'],
    ['App', APP_NAME],
    ['Versión payload', s.version || ''],
    ['Última sincronización', new Date()],
    ['Fecha actual en app', s.fechaActual || ''],
    ['Uso', 'Esta planilla se actualiza desde la app. La pestaña BackupJSON guarda el estado completo para restaurar.'],
    ['Pestañas principales', 'Dashboard, Rutina, Tiempo_Areas, Tiempo_Sesiones, Pomodoro, Cumplimiento, Clientes']
  ], 1);
}

function writeDashboard_(ss, s) {
  const tiempo = s.tiempo || {};
  const data = tiempo.data || {};
  const counts = data.counts || {};
  const clientes = s.clientes || {};
  const cumplimientoHoy = getCumplimientoHoy_(s);
  const values = [
    ['Métrica', 'Valor'],
    ['Última sincronización', new Date()],
    ['Fecha app', s.fechaActual || ''],
    ['Tiempo trabajado hoy', secondsToText_(tiempo.totalTrabajadoSegundos || 0)],
    ['Tiempo trabajado minutos', Math.round((tiempo.totalTrabajadoSegundos || 0) / 60)],
    ['Objetivo diario', secondsToText_(tiempo.objetivoTotalSegundos || 0)],
    ['Cumplimiento rutina hoy', cumplimientoHoy],
    ['Sesiones de tiempo hoy', (data.sessions || []).length],
    ['Contactos prospección', counts.contactos || 0],
    ['Reuniones prospección', counts.reuniones || 0],
    ['Clientes desde tiempo', counts.clientes || 0],
    ['Clientes cerrados app', clientes.total || 0],
    ['Área pomodoro', (s.pomodoro || {}).areaId || ''],
    ['Pomodoros completados hoy', (s.pomodoro || {}).completadosHoy || 0]
  ];
  clearAndWrite_(ss.getSheetByName(SHEETS.dashboard), values, 1);
}

function writeRutina_(ss, s) {
  const items = Array.isArray(s.rutina) ? s.rutina : [];
  const checks = s.checksHoy || {};
  let section = '';
  const rows = [['Orden', 'Tipo', 'Sección', 'Inicio', 'Fin', 'Ícono', 'Título', 'Detalle', 'Etiqueta', 'Color', 'Hecho hoy', 'UID']];
  items.forEach((it, i) => {
    if (it.type === 'section') {
      section = it.titulo || '';
      rows.push([i + 1, 'SECCIÓN', section, '', '', '', section, '', '', '', '', it.uid || '']);
    } else {
      rows.push([i + 1, 'TAREA', section, it.desde || '', it.hasta || '', it.icon || '', it.titulo || '', it.detalle || '', it.tag || '', it.tagClass || '', checks[it.uid] ? 'SI' : 'NO', it.uid || '']);
    }
  });
  clearAndWrite_(ss.getSheetByName(SHEETS.rutina), rows, 1);
}

function writeTiempoAreas_(ss, s) {
  const tiempo = s.tiempo || {};
  const data = tiempo.data || {};
  const counts = data.counts || {};
  const rows = [['N°', 'ID', 'Área', 'Objetivo min', 'Trabajado min', 'Trabajado', '%', 'Faltan min', 'Nota', 'Detalle']];
  const areas = tiempo.areas || [];
  areas.forEach(a => {
    const objetivoMin = a.objetivoMinutos || Math.round((a.objetivo || 0) / 60);
    const trabajadoMin = a.trabajadoMinutos || Math.round((a.trabajadoSegundos || 0) / 60);
    const pct = objetivoMin ? Math.round((trabajadoMin / objetivoMin) * 100) : 0;
    rows.push([a.num || '', a.id || '', a.nombre || '', objetivoMin, trabajadoMin, secondsToText_(a.trabajadoSegundos || 0), pct + '%', Math.max(0, objetivoMin - trabajadoMin), a.nota || '', a.detalle || '']);
  });
  rows.push(['', '', 'Contactos', '', counts.contactos || 0, '', '', '', '', '']);
  rows.push(['', '', 'Reuniones', '', counts.reuniones || 0, '', '', '', '', '']);
  rows.push(['', '', 'Clientes', '', counts.clientes || 0, '', '', '', '', '']);
  rows.push(['', '', 'Limpieza casa', '', counts.limpiezaCasa ? 'SI' : 'NO', '', '', '', '', '']);
  rows.push(['', '', 'Limpieza oficina', '', counts.limpiezaOficina ? 'SI' : 'NO', '', '', '', '', '']);
  clearAndWrite_(ss.getSheetByName(SHEETS.tiempoAreas), rows, 1);
}

function writeTiempoSesiones_(ss, s) {
  const data = (s.tiempo || {}).data || {};
  const sessions = data.sessions || [];
  const areaMap = {};
  ((s.tiempo || {}).areas || []).forEach(a => areaMap[a.id] = a.nombre);
  const rows = [['Fecha', 'Área ID', 'Área', 'Inicio', 'Fin', 'Segundos', 'Minutos', 'Origen']];
  sessions.forEach(sess => {
    rows.push([
      s.fechaActual || '',
      sess.id || '',
      areaMap[sess.id] || sess.id || '',
      sess.desde || '',
      sess.hasta || '',
      sess.segundos || 0,
      Math.round((sess.segundos || 0) / 60),
      sess.pomodoro ? 'Pomodoro' : sess.manual ? 'Manual' : 'Cronómetro'
    ]);
  });
  clearAndWrite_(ss.getSheetByName(SHEETS.tiempoSesiones), rows, 1);
}

function writePomodoro_(ss, s) {
  const p = s.pomodoro || {};
  const rows = [
    ['Campo', 'Valor'],
    ['Fecha', s.fechaActual || ''],
    ['Área seleccionada', p.areaId || ''],
    ['Pomodoros completados hoy', p.completadosHoy || 0]
  ];
  clearAndWrite_(ss.getSheetByName(SHEETS.pomodoro), rows, 1);
}

function writeCumplimiento_(ss, s) {
  const progreso = s.cumplimiento || {};
  const rows = [['Fecha', 'Total tareas', 'Hechas', 'Porcentaje', 'Completo', 'Actualizado']];
  Object.keys(progreso).sort().forEach(fecha => {
    const r = progreso[fecha] || {};
    rows.push([fecha, r.total || 0, r.done || 0, (r.pct || 0) + '%', r.completo ? 'SI' : 'NO', r.updatedAt || '']);
  });
  clearAndWrite_(ss.getSheetByName(SHEETS.cumplimiento), rows, 1);
}

function writeClientes_(ss, s) {
  const c = s.clientes || {};
  const rows = [['Tipo', 'Fecha', 'Estado', 'Valor']];
  rows.push(['Resumen', s.fechaActual || '', 'Clientes cerrados', c.total || 0]);
  const estados = c.calEstado || {};
  Object.keys(estados).sort().forEach(fecha => {
    const val = Number(estados[fecha] || 0);
    if (!val) return;
    rows.push(['Calendario', fecha, val === 2 ? 'Cliente cerrado' : 'Visita', val]);
  });
  clearAndWrite_(ss.getSheetByName(SHEETS.clientes), rows, 1);
}

function writeBackup_(ss, s) {
  const json = JSON.stringify(s);
  const chunkSize = 45000;
  const rows = [['Campo', 'Valor']];
  for (let i = 0; i < json.length; i += chunkSize) {
    rows.push([i === 0 ? 'Último backup completo' : 'Backup continuación', json.slice(i, i + chunkSize)]);
  }
  rows.push(['Actualizado', new Date()]);
  rows.push(['Nota', 'Esta pestaña permite restaurar la app con el botón Cargar desde Sheets. No borres las filas de backup.']);
  clearAndWrite_(ss.getSheetByName(SHEETS.backup), rows, 1);
}

function writeHistorial_(ss, s) {
  const sh = ss.getSheetByName(SHEETS.historial);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Sincronizado', 'Fecha app', 'Tiempo min', 'Sesiones', 'Clientes cerrados', 'Versión']);
    styleSheet_(sh, 6);
  }
  const tiempoMin = Math.round(((s.tiempo || {}).totalTrabajadoSegundos || 0) / 60);
  const sesiones = (((s.tiempo || {}).data || {}).sessions || []).length;
  sh.appendRow([new Date(), s.fechaActual || '', tiempoMin, sesiones, ((s.clientes || {}).total || 0), s.version || '']);
}

function getCumplimientoHoy_(s) {
  const fecha = s.fechaActual || '';
  const progreso = s.cumplimiento || {};
  if (progreso[fecha]) return (progreso[fecha].done || 0) + '/' + (progreso[fecha].total || 0) + ' · ' + (progreso[fecha].pct || 0) + '%';
  const checks = s.checksHoy || {};
  const totalChecks = Object.keys(checks).length;
  const done = Object.values(checks).filter(Boolean).length;
  return done + '/' + totalChecks;
}

function secondsToText_(seconds) {
  seconds = Math.max(0, Math.round(seconds || 0));
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  if (h && m) return h + 'h ' + String(m).padStart(2, '0') + 'm';
  if (h) return h + 'h';
  return m + 'm';
}
