/**
 * 805 MultiServices — API para el Dashboard de Horas Pico
 * --------------------------------------------------------
 * Lee la hoja "DetalleOrden" de CCIAPP en vivo y la entrega en JSON al
 * dashboard (dashboard/index.html). NO envía nombres, teléfonos ni
 * direcciones de clientes: solo fecha, hora, movimiento, empleado,
 * servicio, forma de pago, montos e IDs opacos (convertidos a números).
 *
 * INSTALACIÓN (una sola vez, con la cuenta dueña de CCIAPP):
 *  1. Abre https://script.google.com → "New project". Ponle nombre
 *     "Dashboard Horas Pico".
 *  2. Borra lo que trae Code.gs y pega TODO este archivo. Ctrl+S (Save).
 *  3. En el desplegable de funciones (arriba, junto a "Debug") elige
 *     crearToken → "Run" → "Review permissions" → tu cuenta → "Advanced"
 *     → "Go to Dashboard Horas Pico (unsafe)" → "Allow".
 *  4. Abre "Execution log": copia el TOKEN que aparece.
 *  5. "Deploy" → "New deployment" → engrane ⚙ → "Web app".
 *       Execute as: Me   ·   Who has access: Anyone   → "Deploy".
 *     Copia la "Web app URL" (termina en /exec).
 *  6. Abre el dashboard → botón "Conexión" → pega URL y TOKEN.
 *
 * Si cambias este código: Deploy → Manage deployments → ✏️ → Version:
 * "New version" → Deploy (la URL no cambia).
 * Para invalidar el token: vuelve a correr crearToken y actualízalo en
 * el dashboard.
 *
 * Este archivo no contiene datos de clientes ni contraseñas.
 */

const SPREADSHEET_ID = '13n14htWYPBgfzNOSfCzOKrFITxStsiURs3l-EPbVDpY'; // CCIAPP
const SHEET_NAME = 'DetalleOrden';

// Encabezados que se leen (deben existir en la fila 1 de DetalleOrden)
const COLS = {
  id: 'IDDetalle',
  ticket: 'IDTicket',
  vig: 'Vigencia',
  fecha: 'Date',
  hora: 'Hora',
  mov: 'Movimiento',
  emp: 'NombreEmpleado',
  prod: 'DescripcionProducto',
  cat: 'Categoria',
  fp: 'FormaPago',
  total: 'Total venta',
  entrada: 'Entrada',
  cliente: 'IDCLIENTE'
};

/** Corre esta función UNA vez para crear (o cambiar) el token. */
function crearToken() {
  const token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('DASH_TOKEN', token);
  Logger.log('TOKEN: ' + token);
}

/** Prueba local: corre esta función y revisa el log (no necesita token). */
function probarLectura() {
  const d = leerDatos_();
  Logger.log('Filas: ' + d.filas.length + ' · Empleados: ' + d.dic.emp.length +
             ' · Servicios: ' + d.dic.prod.length);
  Logger.log(JSON.stringify(d.filas.slice(-3)));
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const saved = PropertiesService.getScriptProperties().getProperty('DASH_TOKEN');
  if (!saved) return json_({ ok: false, error: 'Falta crear el token (corre crearToken).' });
  if (String(p.token || '') !== saved) return json_({ ok: false, error: 'Token inválido.' });
  try {
    return json_(Object.assign({ ok: true }, leerDatos_()));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function leerDatos_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('No existe la hoja ' + SHEET_NAME);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const n = lastRow - 1;

  const col = {};
  Object.keys(COLS).forEach(k => {
    const i = headers.indexOf(COLS[k]);
    if (i < 0) throw new Error('No encuentro la columna "' + COLS[k] + '"');
    col[k] = i + 1;
  });

  const empty = { ok: true, generado: new Date().toISOString(), filas: [],
                  dic: { emp: [], prod: [], cat: [], fp: [], vig: [] } };
  if (n < 1) return empty;

  // Fecha y hora se leen como TEXTO mostrado para evitar el desfase de zona
  // horaria que tienen las horas (serial 1899) al leerlas como Date.
  const disp = k => sh.getRange(2, col[k], n, 1).getDisplayValues().map(r => r[0]);
  const vals = k => sh.getRange(2, col[k], n, 1).getValues().map(r => r[0]);

  const fecha = disp('fecha'), hora = disp('hora');
  const id = vals('id'), ticket = vals('ticket'), vig = vals('vig'), mov = vals('mov');
  const emp = vals('emp'), prod = vals('prod'), cat = vals('cat'), fp = vals('fp');
  const total = vals('total'), entrada = vals('entrada'), cliente = vals('cliente');

  const dic = { emp: [], prod: [], cat: [], fp: [], vig: [] };
  const idx = { emp: {}, prod: {}, cat: {}, fp: {}, vig: {}, tk: {}, cl: {} };
  let tkN = 0, clN = 0;
  const code = (k, v) => {
    v = String(v == null ? '' : v).trim();
    if (!(v in idx[k])) { idx[k][v] = dic[k].length; dic[k].push(v); }
    return idx[k][v];
  };

  const filas = [];
  for (let i = 0; i < n; i++) {
    const ymd = parseFecha_(fecha[i]);
    if (!ymd) continue;                       // fila vacía o sin fecha
    const min = parseHora_(hora[i]);
    if (min < 0) continue;                    // sin hora válida
    const tkRaw = String(ticket[i] || id[i] || ('r' + i)).trim();
    const tkKey = tkRaw in idx.tk ? idx.tk[tkRaw] : (idx.tk[tkRaw] = tkN++);
    const clRaw = String(cliente[i] || '').trim();
    const clKey = !clRaw ? -1 : (clRaw in idx.cl ? idx.cl[clRaw] : (idx.cl[clRaw] = clN++));
    filas.push([
      ymd, min, Number(mov[i]) || 0,
      code('emp', emp[i]), code('prod', prod[i]), code('cat', cat[i]),
      code('fp', fp[i]), code('vig', vig[i]),
      num_(total[i]), num_(entrada[i]), tkKey, clKey
    ]);
  }

  return {
    generado: new Date().toISOString(),
    zona: ss.getSpreadsheetTimeZone(),
    columnas: ['ymd', 'min', 'mov', 'emp', 'prod', 'cat', 'fp', 'vig', 'total', 'entrada', 'ticket', 'cliente'],
    dic: dic,
    filas: filas
  };
}

/** "11/6/2026" (M/D/YYYY) o "2026-11-06" → 20261106 */
function parseFecha_(s) {
  s = String(s || '').trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return (+m[3]) * 10000 + (+m[1]) * 100 + (+m[2]);
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return (+m[1]) * 10000 + (+m[2]) * 100 + (+m[3]);
  return 0;
}

/** "9:13:00", "17:45", "5:45 PM" → minutos desde medianoche (-1 si no es válida) */
function parseHora_(s) {
  s = String(s || '').trim();
  const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp]\.?\s*[Mm]\.?)?/);
  if (!m) return -1;
  let h = +m[1];
  const mi = +m[2];
  if (m[3]) {
    const pm = /p/i.test(m[3]);
    if (pm && h < 12) h += 12;
    if (!pm && h === 12) h = 0;
  }
  if (h > 23 || mi > 59) return -1;
  return h * 60 + mi;
}

function num_(v) {
  if (typeof v === 'number') return Math.round(v * 100) / 100;
  const x = parseFloat(String(v || '').replace(/[$,\s]/g, ''));
  return isFinite(x) ? Math.round(x * 100) / 100 : 0;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
