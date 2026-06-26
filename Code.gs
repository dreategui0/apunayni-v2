/**
 * Apunayni / Salvawasi / Intiwasi — Apps Script backend
 *
 * Hoja de cálculo: APUNAYNI_v2
 * Tabs requeridas: FLUJO, RESERVAS, LOGS_IA, USUARIOS
 *
 * RESERVAS — columnas (fila 1):
 *   A:ID  B:CABANA  C:INICIO  D:FIN  E:NOMBRE  F:CEDULA  G:TELEFONO
 *   H:EMAIL  I:N_PERSONAS  J:ABONO  K:TOTAL  L:SALDO  M:ESTADO
 *   N:FRECUENTE  O:NOTAS  P:FECHA_CREACION
 */

// ── Configuración ──────────────────────────────────────────────────────────
const SHEET_ID        = SpreadsheetApp.getActiveSpreadsheet().getId();
const TAB_FLUJO       = 'FLUJO';
const TAB_RESERVAS    = 'RESERVAS';
const TAB_LOGS        = 'LOGS_IA';
const TAB_USUARIOS    = 'USUARIOS';
const CABANAS_VALIDAS = ['APUNAYNI', 'SALVAWASI', 'INTIWASI'];

// ── Punto de entrada HTTP ──────────────────────────────────────────────────
function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    const email  = (body.userEmail || '').trim().toLowerCase();

    // Verificar usuario autorizado
    if (!isAuthorized(email)) {
      return response({ error: 'No autorizado', authRequired: true }, headers);
    }

    switch (action) {
      case 'readRows':       return response(readRows(body),       headers);
      case 'appendRow':      return response(appendRow(body),      headers);
      case 'updateRow':      return response(updateRow(body),      headers);
      case 'readReservas':   return response(readReservas(body),   headers);
      case 'createReserva':  return response(createReserva(body),  headers);
      case 'updateReserva':  return response(updateReserva(body),  headers);
      default:
        return response({ error: 'Acción desconocida: ' + action }, headers);
    }
  } catch (err) {
    return response({ error: err.message }, headers);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: 'Apunayni API activa' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Autorización ───────────────────────────────────────────────────────────
function isAuthorized(email) {
  if (!email) return false;
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(TAB_USUARIOS);
    if (!sheet) return false;
    const data  = sheet.getDataRange().getValues();
    // Columna A = email, columna B = activo (TRUE/Si/1)
    for (let i = 1; i < data.length; i++) {
      const rowEmail  = (data[i][0] || '').toString().trim().toLowerCase();
      const rowActivo = data[i][1];
      if (rowEmail === email && rowActivo) return true;
    }
    return false;
  } catch (err) {
    return false;
  }
}

// ── FLUJO: leer / agregar / actualizar filas ───────────────────────────────
function readRows(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_FLUJO);
  if (!sheet) return { error: 'Hoja FLUJO no encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const rows    = [];

  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });
    rows.push(row);
  }
  return { ok: true, rows };
}

function appendRow(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_FLUJO);
  if (!sheet) return { error: 'Hoja FLUJO no encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const newRow  = headers.map(h => body[h] !== undefined ? body[h] : '');

  // Auto-ID si la columna ID está vacía
  const idIdx = headers.indexOf('ID');
  if (idIdx >= 0 && !newRow[idIdx]) {
    newRow[idIdx] = 'F-' + Date.now();
  }

  sheet.appendRow(newRow);
  return { ok: true, id: newRow[idIdx] || '' };
}

function updateRow(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_FLUJO);
  if (!sheet) return { error: 'Hoja FLUJO no encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const idIdx   = headers.indexOf('ID');
  if (idIdx < 0) return { error: 'Columna ID no encontrada' };

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx].toString() === body.id.toString()) {
      headers.forEach((h, j) => {
        if (h !== 'ID' && body[h] !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(body[h]);
        }
      });
      return { ok: true };
    }
  }
  return { error: 'Fila no encontrada: ' + body.id };
}

// ── RESERVAS: leer ─────────────────────────────────────────────────────────
function readReservas(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_RESERVAS);
  if (!sheet) return { error: 'Hoja RESERVAS no encontrada' };

  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, rows: [] };

  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  const rows    = [];

  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });

    rows.push({
      id:        toString(row['ID']),
      cabana:    toString(row['CABANA']),
      inicio:    formatDate(row['INICIO']),
      fin:       formatDate(row['FIN']),
      nombre:    toString(row['NOMBRE']),
      cedula:    toString(row['CEDULA']),
      telefono:  toString(row['TELEFONO']),
      email:     toString(row['EMAIL']),
      nPersonas: toNum(row['N_PERSONAS']),
      abono:     toNum(row['ABONO']),
      total:     toNum(row['TOTAL']),
      saldo:     toNum(row['SALDO']),
      estado:    toString(row['ESTADO']) || 'PENDIENTE',
      frecuente: row['FRECUENTE'] ? true : false,
      notas:     toString(row['NOTAS'])
    });
  }

  return { ok: true, rows };
}

// ── RESERVAS: crear ────────────────────────────────────────────────────────
function createReserva(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_RESERVAS);
  if (!sheet) return { error: 'Hoja RESERVAS no encontrada' };

  // Validar cabaña
  const cabana = (body.cabana || '').toString().toUpperCase();
  if (!CABANAS_VALIDAS.includes(cabana)) {
    return { error: 'Cabaña inválida: ' + cabana };
  }

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());

  // Generar ID único
  const existing = data.slice(1).map(r => r[headers.indexOf('ID')].toString());
  const id = generarIdReserva(existing);

  // Construir fila según las columnas presentes
  const newRow = headers.map(h => {
    switch (h) {
      case 'ID':              return id;
      case 'CABANA':          return cabana;
      case 'INICIO':          return body.fecha_inicio || '';
      case 'FIN':             return body.fecha_fin    || '';
      case 'NOMBRE':          return body.nombre       || '';
      case 'CEDULA':          return body.cedula       || '';
      case 'TELEFONO':        return body.telefono     || '';
      case 'EMAIL':           return body.email        || '';
      case 'N_PERSONAS':      return body.n_personas   || 1;
      case 'ABONO':           return body.abono        || 0;
      case 'TOTAL':           return body.total        || 0;
      case 'SALDO':           return body.saldo        || 0;
      case 'ESTADO':          return 'PENDIENTE';
      case 'FRECUENTE':       return '';
      case 'NOTAS':           return body.notas        || '';
      case 'FECHA_CREACION':  return new Date().toISOString().slice(0, 10);
      default:                return '';
    }
  });

  sheet.appendRow(newRow);
  logAction('createReserva', body.userEmail, { id, cabana, nombre: body.nombre });
  return { ok: true, id };
}

// ── RESERVAS: actualizar ───────────────────────────────────────────────────
function updateReserva(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_RESERVAS);
  if (!sheet) return { error: 'Hoja RESERVAS no encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  const idIdx   = headers.indexOf('ID');
  if (idIdx < 0) return { error: 'Columna ID no encontrada en RESERVAS' };

  const targetId = (body.id || '').toString();
  let rowIndex   = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx].toString() === targetId) { rowIndex = i; break; }
  }
  if (rowIndex < 0) return { error: 'Reserva no encontrada: ' + targetId };

  // Mapa de campos que se pueden actualizar
  const fieldMap = {
    'ESTADO':     body.estado_pago,
    'TOTAL':      body.total_alquiler !== undefined ? body.total_alquiler : body.total,
    'ABONO':      body.abono,
    'SALDO':      body.saldo,
    'NOTAS':      body.notas,
    'CABANA':     body.cabana ? body.cabana.toUpperCase() : undefined,
    'INICIO':     body.fecha_inicio,
    'FIN':        body.fecha_fin,
    'NOMBRE':     body.nombre,
    'CEDULA':     body.cedula,
    'TELEFONO':   body.telefono,
    'N_PERSONAS': body.n_personas,
    'FRECUENTE':  body.frecuente
  };

  headers.forEach((h, j) => {
    const val = fieldMap[h];
    if (val !== undefined && val !== null) {
      sheet.getRange(rowIndex + 1, j + 1).setValue(val);
    }
  });

  // Recalcular saldo si cambiaron total o abono pero no se envió saldo explícito
  if ((body.total_alquiler !== undefined || body.abono !== undefined) && body.saldo === undefined) {
    const totalIdx = headers.indexOf('TOTAL');
    const abonoIdx = headers.indexOf('ABONO');
    const saldoIdx = headers.indexOf('SALDO');
    if (totalIdx >= 0 && abonoIdx >= 0 && saldoIdx >= 0) {
      const updatedData = sheet.getRange(rowIndex + 1, 1, 1, headers.length).getValues()[0];
      const nuevoSaldo  = Math.max(0, (updatedData[totalIdx] || 0) - (updatedData[abonoIdx] || 0));
      sheet.getRange(rowIndex + 1, saldoIdx + 1).setValue(nuevoSaldo);
    }
  }

  logAction('updateReserva', body.userEmail, { id: targetId, estado: body.estado_pago });
  return { ok: true };
}

// ── Helpers ────────────────────────────────────────────────────────────────
function toString(val) {
  if (val === null || val === undefined) return '';
  return val.toString().trim();
}

function toNum(val) {
  const n = parseFloat(val);
  return isNaN(n) ? 0 : n;
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return val.toString().slice(0, 10);
}

function generarIdReserva(existingIds) {
  const prefix = 'RES-';
  let max = 0;
  existingIds.forEach(id => {
    if (id.startsWith(prefix)) {
      const n = parseInt(id.replace(prefix, ''), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return prefix + String(max + 1).padStart(3, '0');
}

function logAction(action, email, extra) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(TAB_LOGS);
    if (!sheet) return;
    sheet.appendRow([new Date().toISOString(), action, email, JSON.stringify(extra || {})]);
  } catch (e) { /* silencioso */ }
}

function response(data, headers) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── Setup inicial: crea encabezados si las hojas están vacías ──────────────
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // RESERVAS
  let res = ss.getSheetByName(TAB_RESERVAS);
  if (!res) res = ss.insertSheet(TAB_RESERVAS);
  if (res.getLastRow() === 0) {
    res.appendRow(['ID','CABANA','INICIO','FIN','NOMBRE','CEDULA','TELEFONO',
                   'EMAIL','N_PERSONAS','ABONO','TOTAL','SALDO','ESTADO',
                   'FRECUENTE','NOTAS','FECHA_CREACION']);
    res.setFrozenRows(1);
  }

  // LOGS_IA
  let logs = ss.getSheetByName(TAB_LOGS);
  if (!logs) logs = ss.insertSheet(TAB_LOGS);
  if (logs.getLastRow() === 0) {
    logs.appendRow(['TIMESTAMP','ACCION','USUARIO','DETALLE']);
    logs.setFrozenRows(1);
  }

  // USUARIOS
  let usuarios = ss.getSheetByName(TAB_USUARIOS);
  if (!usuarios) usuarios = ss.insertSheet(TAB_USUARIOS);
  if (usuarios.getLastRow() === 0) {
    usuarios.appendRow(['EMAIL','ACTIVO','NOMBRE']);
    usuarios.setFrozenRows(1);
    // Agregar usuario inicial
    usuarios.appendRow(['dreategui@unal.edu.co', true, 'Admin']);
  }

  SpreadsheetApp.flush();
  Browser.msgBox('✅ Hojas configuradas correctamente');
}
