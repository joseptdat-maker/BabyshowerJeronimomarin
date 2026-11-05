/**
 * Lista de regalos · Baby shower de Jerónimo
 * Guarda en esta hoja de cálculo quién va a dar cada regalo.
 *
 * Pestaña "Regalos": una fila por cada regalo apartado.
 *   Fecha | Regalo | Categoría | Nombre | Mensaje | Estado | Código del regalo | Dispositivo
 * Estado = "Activo" o "Cancelado" (si la persona se arrepiente, la fila queda como Cancelado).
 *
 * Los invitados NO ven los nombres; solo ven cuántas personas más darán cada regalo.
 */

const HOJA = 'Regalos';
const ENCABEZADOS = ['Fecha', 'Regalo', 'Categoría', 'Nombre', 'Mensaje', 'Estado', 'Código del regalo', 'Dispositivo'];

// Varias personas pueden escoger el mismo regalo; la página les avisa que alguien más también lo dará.

// Lo que hace la página: reservar, cancelar, otro regalo o consultar el estado.
function procesar(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // evita choques si dos personas guardan al mismo tiempo
    d = d || {};
    const device = limpiar(d.device, 60);
    const hoja = obtenerHoja();

    if (d.accion === 'reservar') {
      const id = limpiar(d.id, 40);
      const nombre = limpiar(d.nombre, 80);
      if (!id || nombre.length < 2) return ({ ok: false, error: 'Escribe tu nombre, por favor.', estado: estado(hoja, device) });
      const activos = filas(hoja).filter(r => r.estado === 'Activo' && r.id === id);
      if (activos.some(r => r.device === device)) return ({ ok: false, error: 'Ya apartaste este regalo.', estado: estado(hoja, device) });
      hoja.appendRow([new Date(), limpiar(d.regalo, 120), limpiar(d.categoria, 60), nombre, limpiar(d.mensaje, 300), 'Activo', id, device]);
    }

    else if (d.accion === 'cancelar') {
      const id = limpiar(d.id, 40);
      filas(hoja).forEach(r => {
        if (r.id === id && r.device === device && r.estado === 'Activo') hoja.getRange(r.fila, 6).setValue('Cancelado');
      });
    }

    else if (d.accion === 'otro') {
      const nombre = limpiar(d.nombre, 80), regalo = limpiar(d.regalo, 120);
      if (!regalo || nombre.length < 2) return ({ ok: false, error: 'Faltan datos.' });
      hoja.appendRow([new Date(), regalo, 'Otro regalo (fuera de la lista)', nombre, '', 'Activo', 'otro', device]);
    }

    return ({ ok: true, estado: estado(hoja, device) });
  } catch (err) {
    return ({ ok: false, error: 'No se pudo guardar. Intenta de nuevo.' });
  } finally {
    lock.releaseLock();
  }
}

// La página llama a esta función directamente (google.script.run).
function api(d) {
  return procesar(d);
}

// Por si la página se publica en otro sitio (usa fetch POST).
function doPost(e) {
  return json(procesar(JSON.parse(e.postData.contents || '{}')));
}

// Abrir el enlace /exec muestra la página de regalos.
// Para ver solo los datos: agrega ?datos=1 al final del enlace.
function doGet(e) {
  if (e && e.parameter && e.parameter.datos) return json({ ok: true, estado: estado(obtenerHoja(), '') });
  return HtmlService.createHtmlOutputFromFile('Pagina')
    .setTitle('Lista de regalos · Baby shower de Jerónimo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- utilidades ---------- */

// Ejecuta esta función una vez desde el editor (botón ▶ Ejecutar) para autorizar
// y ver en el registro el enlace de la hoja donde quedan los regalos.
function crearHoja() {
  obtenerHoja();
  Logger.log('Tus regalos se guardan aquí: ' + libro().getUrl());
}

// Funciona tanto si el script se creó desde la hoja (Extensiones → Apps Script)
// como si se creó aparte en script.google.com: en ese caso crea la hoja
// "Regalos Jerónimo" en tu Drive y la recuerda.
function libro() {
  const activa = SpreadsheetApp.getActiveSpreadsheet();
  if (activa) return activa;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const nueva = SpreadsheetApp.create('Regalos Jerónimo');
  props.setProperty('SHEET_ID', nueva.getId());
  return nueva;
}

function obtenerHoja() {
  const ss = libro();
  let hoja = ss.getSheetByName(HOJA);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA);
    hoja.appendRow(ENCABEZADOS);
    hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold').setBackground('#BBD5F2').setFontColor('#2F5C9E');
    hoja.setFrozenRows(1);
    hoja.setColumnWidths(1, ENCABEZADOS.length, 170);
    hoja.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  }
  return hoja;
}

function filas(hoja) {
  const n = hoja.getLastRow() - 1;
  if (n < 1) return [];
  return hoja.getRange(2, 1, n, ENCABEZADOS.length).getValues().map((v, i) => ({
    fila: i + 2, estado: String(v[5]), id: String(v[6]), device: String(v[7]),
  }));
}

// Lo que ven los invitados: cuántas personas apartaron cada regalo (sin nombres)
function estado(hoja, device) {
  const out = {};
  filas(hoja).forEach(r => {
    if (r.estado !== 'Activo' || r.id === 'otro') return;
    out[r.id] = out[r.id] || { total: 0, mio: false };
    out[r.id].total++;
    if (device && r.device === device) out[r.id].mio = true;
  });
  return out;
}

function limpiar(v, max) {
  // quita espacios y evita que un texto se tome como fórmula en la hoja
  return String(v == null ? '' : v).trim().replace(/^[=+\-@]+/, '').slice(0, max);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
