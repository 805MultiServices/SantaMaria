/**
 * 805 MultiServices — Recibir citas de la página web
 * --------------------------------------------------
 * Guarda cada solicitud en la hoja "Citas" y manda un correo a NOTIFY_TO.
 * Instalar en una Google Sheet de la cuenta 805MultiDMV@gmail.com:
 *   Extensions → Apps Script → pega este código → Deploy → New deployment
 *   → Web app → Execute as: Me · Who has access: Anyone → Deploy.
 * Copia la URL que termina en /exec y pégala en index.html (CITAS_URL).
 *
 * Este archivo no contiene datos de clientes ni contraseñas.
 */

const NOTIFY_TO = '805MultiDMV@gmail.com';
const SHEET_NAME = 'Citas';
const FIELDS = ['servicio', 'nombre', 'telefono', 'correo', 'fecha', 'hora', 'notas', 'idioma'];
const REQUIRED = ['servicio', 'nombre', 'telefono', 'fecha', 'hora'];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Trampa para robots: el campo "website" está oculto para las personas.
  if (p.website) return json_({ ok: true });

  for (const k of REQUIRED) {
    if (!String(p[k] || '').trim()) return json_({ ok: false, error: 'Falta: ' + k });
  }

  const v = {};
  FIELDS.forEach(k => { v[k] = clean_(p[k]); });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Recibido', 'Servicio', 'Nombre', 'Teléfono', 'Correo', 'Día', 'Hora', 'Comentarios', 'Idioma', 'Estado']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date()].concat(FIELDS.map(k => safeCell_(v[k]))).concat(['Pendiente']));
  } finally {
    lock.releaseLock();
  }

  const body =
    'Nueva solicitud de cita desde la página web\n\n' +
    'Servicio: ' + v.servicio + '\n' +
    'Nombre: ' + v.nombre + '\n' +
    'Teléfono: ' + v.telefono + '\n' +
    'Correo: ' + (v.correo || '—') + '\n' +
    'Día: ' + v.fecha + '\n' +
    'Hora: ' + v.hora + '\n' +
    'Comentarios: ' + (v.notas || '—') + '\n' +
    'Idioma: ' + v.idioma + '\n\n' +
    'Recuerda llamar o escribir al cliente para confirmar.';

  const opts = { to: NOTIFY_TO, subject: 'Nueva cita: ' + v.servicio + ' – ' + v.fecha + ' ' + v.hora, body: body, name: '805 MultiServices Web' };
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.correo)) opts.replyTo = v.correo;
  MailApp.sendEmail(opts);

  return json_({ ok: true });
}

// Para probar desde el editor: elige "testCita" en el menú y presiona Run.
function testCita() {
  doPost({ parameter: {
    servicio: 'Live Scan (huellas / fingerprints)', nombre: 'Cliente de Prueba', telefono: '(805) 555-0100',
    correo: '', fecha: '2026-11-02', hora: '10:00 AM', notas: 'Prueba', idioma: 'es'
  } });
}

function clean_(s) {
  return String(s || '').replace(/[\r\n]+/g, ' ').trim().slice(0, 500);
}

// Evita que un texto que empieza con = + - @ se ejecute como fórmula en la hoja.
function safeCell_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
