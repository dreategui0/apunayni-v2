// ============================================================
// APUNAYNI v2 — Google Apps Script
// Proxy seguro: contabilidad + agente IA + reservas
// ============================================================

var CLAUDE_API_KEY = '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';
var SHEET_ID       = '1ueAv6JhzTYhJ7y3ug-N6-IbPBOfuuGL4rUKUE1WuHaQ';
var CLAUDE_MODEL   = 'claude-sonnet-4-5';

// ── ROUTER PRINCIPAL ─────────────────────────────────────────
function doGet(e) {
  var action = (e.parameter && e.parameter.action) ? e.parameter.action : '';
  if (action === 'ping') return jsonResponse({status: 'ok', version: '2.1'});
  return jsonResponse({error: 'GET no soportado'});
}

function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action;
    if (action === 'verifyUser')         return handleVerifyUser(data);
    if (!isAuthorizedEmail(data.userEmail)) {
      return jsonResponse({error: 'No autorizado', authRequired: true});
    }
    if (action === 'appendRow')          return handleAppendRow(data);
    if (action === 'chat')               return handleChat(data);
    if (action === 'createReserva')      return handleCreateReserva(data);
    if (action === 'updateReserva')      return handleUpdateReserva(data);
    if (action === 'getDisponibilidad')  return handleGetDisponibilidad(data);
    if (action === 'readFlujo')          return handleReadFlujo(data);
    if (action === 'readReservas')       return handleReadReservas(data);
    if (action === 'readDashboard')      return handleReadDashboard(data);
    return jsonResponse({error: 'Accion desconocida: ' + action});
  } catch(err) {
    return jsonResponse({error: err.message});
  }
}

// ── CONTABILIDAD ──────────────────────────────────────────────
function handleAppendRow(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('FLUJO');
  if (!ws) return jsonResponse({error: 'Hoja FLUJO no encontrada'});
  ws.appendRow([data.fecha, data.valor, data.desc, data.medio, data.tipo, data.ie, data.notas || '']);
  return jsonResponse({ok: true});
}

// ── AGENTE IA ─────────────────────────────────────────────────
function handleChat(data) {
  var startTime  = Date.now();
  var messages   = data.messages  || [];
  var sessionId  = data.sessionId || 'default';

  var context = buildRAGContext();

  var systemPrompt =
    'Eres el asistente operativo de Apunayni, un negocio de cabanas en Colombia ' +
    '(cabanas Apunayni, Salvawasi e Intiwasi).\n' +
    'Tienes acceso COMPLETO en tiempo real a TODOS los datos del negocio: contabilidad, reservas, ingresos detallados, egresos detallados, montos esperados, abonos, saldos.\n\n' +
    'DATOS ACTUALES:\n' + context + '\n\n' +
    'REGLAS:\n' +
    '- Responde siempre en espanol\n' +
    '- Se conciso pero completo\n' +
    '- SOLO usa los numeros exactos del contexto. Nunca inventes ni estimes cantidades.\n' +
    '- El contexto SI incluye ingresos por mes, detalle de reservas con sus valores monetarios, abonos recibidos y saldos por cobrar. Usalo activamente.\n' +
    '- Cuando te pregunten por ingresos de un mes, SUMA los ingresos registrados en FLUJO de ese mes mas los abonos VALIDADOS de reservas si aplica.\n' +
    '- Para reservas PENDIENTES, ya tienes el total esperado y el saldo por cobrar. Usalo para calcular ingresos potenciales.\n' +
    '- Si el contexto no tiene una respuesta exacta, di claramente que no tienes ese dato\n' +
    '- Cabanas: Apunayni (🏕), Salvawasi (⛺), Intiwasi (🌄)\n' +
    '- Precios: 1 persona=150000, 2=200000, 3+=200000+20000*(N-2)\n' +
    '- Clientes frecuentes (mas de 1 visita): descuento 10%\n' +
    '- Cancelacion: reembolso total si cancela 48h+ antes\n\n' +
    'Si necesitas ejecutar una accion responde SOLO con JSON valido:\n' +
    '{"tool":"crear_reserva","cabana":"APUNAYNI","fecha_inicio":"2026-06-01","fecha_fin":"2026-06-02","nombre":"Juan","cedula":"123","telefono":"310","n_personas":2,"abono":100000}\n' +
    '{"tool":"calcular_precio","n_personas":3,"cliente_frecuente":false}\n' +
    '{"tool":"buscar_reservas","filtro":"Juan"}';

  var payload = {
    model:      CLAUDE_MODEL,
    max_tokens: 1024,
    system:     systemPrompt,
    messages:   messages
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:            'POST',
    headers: {
      'x-api-key':          CLAUDE_API_KEY,
      'anthropic-version':  '2023-06-01',
      'content-type':       'application/json'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var latency = Date.now() - startTime;
  var result  = JSON.parse(response.getContentText());

  if (result.error) {
    return jsonResponse({error: result.error.message, latency: latency});
  }

  var assistantText = result.content[0].text;
  var usage         = result.usage || {};

  logObservabilidad({
    sessionId:    sessionId,
    latency:      latency,
    inputTokens:  usage.input_tokens  || 0,
    outputTokens: usage.output_tokens || 0,
    model:        CLAUDE_MODEL,
    error:        null
  });

  var toolResult = null;
  try {
    var trimmed = assistantText.trim();
    if (trimmed.charAt(0) === '{') {
      var parsed = JSON.parse(trimmed);
      if (parsed.tool) toolResult = executeTool(parsed);
    }
  } catch(e) { /* respuesta normal de texto */ }

  return jsonResponse({
    reply:        assistantText,
    toolResult:   toolResult,
    usage:        usage,
    latency:      latency,
    costEstimate: estimateCost(usage)
  });
}

// ── RAG: construye contexto agregado desde el Sheet ──────────
function buildRAGContext() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var parts = [];
  var hoy   = new Date(); hoy.setHours(0,0,0,0);
  var hoyStr = hoy.getFullYear() + '-' + ('0'+(hoy.getMonth()+1)).slice(-2) + '-' + ('0'+hoy.getDate()).slice(-2);
  var mesActualKey = hoy.getFullYear() + '-' + ('0'+(hoy.getMonth()+1)).slice(-2);
  var mesesNom = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  parts.push('FECHA ACTUAL: ' + hoyStr + ' (' + mesesNom[hoy.getMonth()] + ' ' + hoy.getFullYear() + ')');

  // ═════ CONTABILIDAD COMPLETA ═════
  try {
    var flujo = ss.getSheetByName('FLUJO');
    if (flujo) {
      var rows = flujo.getDataRange().getValues();
      var saldos = {};
      var ingresosPorMes = {};
      var egresosPorMes  = {};
      var ingresosMesActual = [];
      var egresosMesActual  = [];

      for (var i = 3; i < rows.length; i++) {
        if (!rows[i][0]) continue;
        var medio = String(rows[i][3] || '').toUpperCase().trim();
        var valor = parseFloat(rows[i][1]) || 0;
        var desc  = String(rows[i][2] || '').trim();
        var tipo  = String(rows[i][4] || '').toUpperCase().trim();
        var ie    = String(rows[i][5] || '').toUpperCase().trim();
        var fecha = rows[i][0];
        var fechaObj = (fecha instanceof Date) ? fecha : new Date(fecha);
        if (isNaN(fechaObj.getTime())) continue;
        var mesKey  = fechaObj.getFullYear() + '-' + ('0'+(fechaObj.getMonth()+1)).slice(-2);
        var fechaStr = fechaObj.getFullYear() + '-' + ('0'+(fechaObj.getMonth()+1)).slice(-2) + '-' + ('0'+fechaObj.getDate()).slice(-2);

        if (medio) saldos[medio] = (saldos[medio] || 0) + valor;
        if (tipo === 'SALDOS') continue;

        if (ie === 'INGRESO') {
          if (!ingresosPorMes[mesKey]) ingresosPorMes[mesKey] = {};
          ingresosPorMes[mesKey][tipo || 'OTROS'] = (ingresosPorMes[mesKey][tipo || 'OTROS'] || 0) + Math.abs(valor);
          if (mesKey === mesActualKey) ingresosMesActual.push({fecha: fechaStr, valor: Math.abs(valor), desc: desc, tipo: tipo, medio: medio});
        } else if (ie === 'EGRESO') {
          if (!egresosPorMes[mesKey]) egresosPorMes[mesKey] = {};
          egresosPorMes[mesKey][tipo || 'OTROS'] = (egresosPorMes[mesKey][tipo || 'OTROS'] || 0) + Math.abs(valor);
          if (mesKey === mesActualKey) egresosMesActual.push({fecha: fechaStr, valor: Math.abs(valor), desc: desc, tipo: tipo, medio: medio});
        }
      }

      var saldoLines = 'SALDOS ACTUALES POR CUENTA:';
      var totalSaldo = 0;
      Object.keys(saldos).forEach(function(m) {
        saldoLines += '\n  ' + m + ': $' + Math.round(saldos[m]).toLocaleString('es-CO');
        totalSaldo += saldos[m];
      });
      saldoLines += '\n  TOTAL GENERAL: $' + Math.round(totalSaldo).toLocaleString('es-CO');
      parts.push(saldoLines);

      var ingLines = 'INGRESOS POR MES Y CATEGORIA:';
      Object.keys(ingresosPorMes).sort().forEach(function(mk) {
        var totMes = 0;
        Object.keys(ingresosPorMes[mk]).forEach(function(t) { totMes += ingresosPorMes[mk][t]; });
        ingLines += '\n  ' + mk + ' (total $' + Math.round(totMes).toLocaleString('es-CO') + '):';
        Object.keys(ingresosPorMes[mk]).forEach(function(t) {
          ingLines += '\n    ' + t + ': $' + Math.round(ingresosPorMes[mk][t]).toLocaleString('es-CO');
        });
      });
      parts.push(ingLines);

      var egrLines = 'EGRESOS POR MES Y CATEGORIA:';
      Object.keys(egresosPorMes).sort().forEach(function(mk) {
        var totMes = 0;
        Object.keys(egresosPorMes[mk]).forEach(function(t) { totMes += egresosPorMes[mk][t]; });
        egrLines += '\n  ' + mk + ' (total $' + Math.round(totMes).toLocaleString('es-CO') + '):';
        Object.keys(egresosPorMes[mk]).forEach(function(t) {
          egrLines += '\n    ' + t + ': $' + Math.round(egresosPorMes[mk][t]).toLocaleString('es-CO');
        });
      });
      parts.push(egrLines);

      if (ingresosMesActual.length > 0) {
        var detIng = 'INGRESOS DETALLADOS DEL MES ACTUAL (' + mesActualKey + '):';
        ingresosMesActual.sort(function(a,b) { return a.fecha < b.fecha ? -1 : 1; });
        for (var k = 0; k < Math.min(ingresosMesActual.length, 50); k++) {
          var x = ingresosMesActual[k];
          detIng += '\n  ' + x.fecha + ' | $' + Math.round(x.valor).toLocaleString('es-CO') + ' | ' + x.tipo + ' | ' + x.medio + ' | ' + x.desc;
        }
        parts.push(detIng);
      }

      if (egresosMesActual.length > 0) {
        var detEgr = 'EGRESOS DETALLADOS DEL MES ACTUAL (' + mesActualKey + '):';
        egresosMesActual.sort(function(a,b) { return a.fecha < b.fecha ? -1 : 1; });
        for (var k2 = 0; k2 < Math.min(egresosMesActual.length, 50); k2++) {
          var y = egresosMesActual[k2];
          detEgr += '\n  ' + y.fecha + ' | $' + Math.round(y.valor).toLocaleString('es-CO') + ' | ' + y.tipo + ' | ' + y.medio + ' | ' + y.desc;
        }
        parts.push(detEgr);
      }
    }
  } catch(e) { parts.push('CONTABILIDAD: error - ' + e.message); }

  // ═════ RESERVAS COMPLETAS (incluye INTIWASI) ═════
  try {
    var reservas = ss.getSheetByName('RESERVAS');
    if (reservas && reservas.getLastRow() > 1) {
      var rrows = reservas.getDataRange().getValues();

      var totalActivas = 0, apuActivas = 0, salvActivas = 0, intiActivas = 0;
      var totalHist = 0, apuHist = 0, salvHist = 0, intiHist = 0;
      var totalPendientes = 0, totalCanceladas = 0, totalValidadas = 0;
      var porMes = {};
      var ingresoEsperadoMes = {};
      var detalleActivas = [];
      var detallePorMes  = {};

      for (var r = 1; r < rrows.length; r++) {
        if (!rrows[r][0]) continue;
        var cabana     = String(rrows[r][1] || '').toUpperCase();
        var fechaIniObj = (rrows[r][2] instanceof Date) ? rrows[r][2] : new Date(rrows[r][2]);
        var fechaFinObj = (rrows[r][3] instanceof Date) ? rrows[r][3] : new Date(rrows[r][3]);
        if (isNaN(fechaIniObj.getTime())) continue;
        var nombre    = String(rrows[r][4] || '');
        var nPersonas = parseInt(rrows[r][8]) || 1;
        var abono     = parseFloat(rrows[r][10]) || 0;
        var total     = parseFloat(rrows[r][11]) || 0;
        var saldo     = parseFloat(rrows[r][12]) || 0;
        var estado    = String(rrows[r][13] || '').toUpperCase();
        var frecuente = String(rrows[r][14] || '').toUpperCase() === 'SI';

        if (estado === 'CANCELADO') { totalCanceladas++; continue; }
        if (estado === 'PENDIENTE') totalPendientes++;
        if (estado === 'VALIDADO')  totalValidadas++;

        var mesKey2    = fechaIniObj.getFullYear() + '-' + ('0'+(fechaIniObj.getMonth()+1)).slice(-2);
        var fechaIniStr = fechaIniObj.getFullYear() + '-' + ('0'+(fechaIniObj.getMonth()+1)).slice(-2) + '-' + ('0'+fechaIniObj.getDate()).slice(-2);
        var fechaFinStr = fechaFinObj.getFullYear() + '-' + ('0'+(fechaFinObj.getMonth()+1)).slice(-2) + '-' + ('0'+fechaFinObj.getDate()).slice(-2);

        if (!porMes[mesKey2]) porMes[mesKey2] = {APUNAYNI:0, SALVAWASI:0, INTIWASI:0};
        if (porMes[mesKey2][cabana] !== undefined) porMes[mesKey2][cabana]++;

        if (!ingresoEsperadoMes[mesKey2]) ingresoEsperadoMes[mesKey2] = {
          APUNAYNI:  {total:0, abono:0, saldo:0, count:0, pendientes:0},
          SALVAWASI: {total:0, abono:0, saldo:0, count:0, pendientes:0},
          INTIWASI:  {total:0, abono:0, saldo:0, count:0, pendientes:0}
        };
        if (ingresoEsperadoMes[mesKey2][cabana]) {
          ingresoEsperadoMes[mesKey2][cabana].total += total;
          ingresoEsperadoMes[mesKey2][cabana].abono += abono;
          ingresoEsperadoMes[mesKey2][cabana].saldo += saldo;
          ingresoEsperadoMes[mesKey2][cabana].count++;
          if (estado === 'PENDIENTE') ingresoEsperadoMes[mesKey2][cabana].pendientes++;
        }

        if (!detallePorMes[mesKey2]) detallePorMes[mesKey2] = [];
        detallePorMes[mesKey2].push({
          inicio: fechaIniStr, fin: fechaFinStr, cabana: cabana,
          nombre: nombre, nPersonas: nPersonas,
          total: total, abono: abono, saldo: saldo,
          estado: estado, frecuente: frecuente
        });

        if (fechaFinObj >= hoy) {
          totalActivas++;
          if (cabana === 'APUNAYNI')  apuActivas++;
          else if (cabana === 'SALVAWASI') salvActivas++;
          else if (cabana === 'INTIWASI')  intiActivas++;
          if (detalleActivas.length < 30) {
            detalleActivas.push({
              inicio: fechaIniStr, fin: fechaFinStr, cabana: cabana,
              nombre: nombre, total: total, abono: abono, saldo: saldo, estado: estado
            });
          }
        } else {
          totalHist++;
          if (cabana === 'APUNAYNI')  apuHist++;
          else if (cabana === 'SALVAWASI') salvHist++;
          else if (cabana === 'INTIWASI')  intiHist++;
        }
      }

      detalleActivas.sort(function(a,b) { return a.inicio < b.inicio ? -1 : 1; });

      var resLines = 'RESERVAS - RESUMEN:';
      resLines += '\n  TOTAL ACTIVAS: ' + totalActivas + ' (APUNAYNI: ' + apuActivas + ', SALVAWASI: ' + salvActivas + ', INTIWASI: ' + intiActivas + ')';
      resLines += '\n  TOTAL HISTORICAS: ' + totalHist + ' (APUNAYNI: ' + apuHist + ', SALVAWASI: ' + salvHist + ', INTIWASI: ' + intiHist + ')';
      resLines += '\n  PENDIENTES de validar pago: ' + totalPendientes;
      resLines += '\n  VALIDADAS (pago confirmado): ' + totalValidadas;
      resLines += '\n  CANCELADAS: ' + totalCanceladas;
      parts.push(resLines);

      var porMesLines = 'RESERVAS POR MES Y CABANA:';
      Object.keys(porMes).sort().forEach(function(mk) {
        var pm  = porMes[mk];
        var tot = pm.APUNAYNI + pm.SALVAWASI + pm.INTIWASI;
        if (tot > 0) porMesLines += '\n  ' + mk + ': ' + tot + ' total (APUNAYNI: ' + pm.APUNAYNI + ', SALVAWASI: ' + pm.SALVAWASI + ', INTIWASI: ' + pm.INTIWASI + ')';
      });
      parts.push(porMesLines);

      var espLines = 'DINERO POR MES Y CABANA (de reservas no canceladas):';
      Object.keys(ingresoEsperadoMes).sort().forEach(function(mk) {
        var iem = ingresoEsperadoMes[mk];
        ['APUNAYNI','SALVAWASI','INTIWASI'].forEach(function(cab) {
          if (iem[cab] && iem[cab].count > 0) {
            espLines += '\n  ' + mk + ' - ' + cab + ' (' + iem[cab].count + ' reservas, ' + iem[cab].pendientes + ' pendientes):';
            espLines += '\n    Total esperado: $' + Math.round(iem[cab].total).toLocaleString('es-CO');
            espLines += '\n    Abono recibido: $' + Math.round(iem[cab].abono).toLocaleString('es-CO');
            espLines += '\n    Saldo por cobrar: $' + Math.round(iem[cab].saldo).toLocaleString('es-CO');
          }
        });
      });
      parts.push(espLines);

      if (detallePorMes[mesActualKey]) {
        var detMes = 'DETALLE DE TODAS LAS RESERVAS DEL MES ACTUAL (' + mesActualKey + '):';
        detallePorMes[mesActualKey].sort(function(a,b) { return a.inicio < b.inicio ? -1 : 1; });
        detallePorMes[mesActualKey].forEach(function(rr) {
          detMes += '\n  ' + rr.inicio + ' a ' + rr.fin + ' | ' + rr.cabana + ' | ' + rr.nombre + ' (' + rr.nPersonas + 'p)';
          detMes += ' | Total $' + Math.round(rr.total).toLocaleString('es-CO');
          detMes += ' | Abono $' + Math.round(rr.abono).toLocaleString('es-CO');
          detMes += ' | Saldo $' + Math.round(rr.saldo).toLocaleString('es-CO');
          detMes += ' | ' + rr.estado;
          if (rr.frecuente) detMes += ' | FRECUENTE';
        });
        parts.push(detMes);
      }

      if (detalleActivas.length > 0) {
        var proxLines = 'PROXIMAS RESERVAS ACTIVAS (futuras, ordenadas por fecha):';
        detalleActivas.forEach(function(d) {
          proxLines += '\n  ' + d.inicio + ' a ' + d.fin + ' | ' + d.cabana + ' | ' + d.nombre;
          proxLines += ' | Total $' + Math.round(d.total).toLocaleString('es-CO');
          proxLines += ' | Abono $' + Math.round(d.abono).toLocaleString('es-CO');
          proxLines += ' | Saldo $' + Math.round(d.saldo).toLocaleString('es-CO');
          proxLines += ' | ' + d.estado;
        });
        parts.push(proxLines);
      }
    } else {
      parts.push('RESERVAS: sin datos aun (hoja vacia)');
    }
  } catch(e) { parts.push('RESERVAS: error - ' + e.message); }

  return parts.join('\n\n');
}

// ── TOOL EXECUTOR ─────────────────────────────────────────────
function executeTool(toolCall) {
  if (toolCall.tool === 'crear_reserva')   return handleCreateReserva(toolCall);
  if (toolCall.tool === 'calcular_precio') return calcularPrecio(toolCall.n_personas, toolCall.cliente_frecuente);
  if (toolCall.tool === 'buscar_reservas') return buscarReservas(toolCall.filtro);
  return {error: 'Tool desconocida: ' + toolCall.tool};
}

// ── RESERVAS: crear ───────────────────────────────────────────
function handleCreateReserva(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('RESERVAS');
  if (!ws) {
    ws = ss.insertSheet('RESERVAS');
    ws.appendRow(['ID','CABANA','FECHA_INICIO','FECHA_FIN','NOMBRE','CEDULA',
      'TELEFONO','EMAIL','N_PERSONAS','VISITANTES_EXTRA','ABONO',
      'TOTAL_ALQUILER','SALDO','ESTADO_PAGO','CLIENTE_FRECUENTE',
      'COMPROBANTE_URL','FECHA_REGISTRO','NOTAS']);
    ws.setFrozenRows(1);
  }
  var id    = 'R' + Date.now();
  var abono = parseFloat(data.abono) || 0;
  // Acepta tanto 'total' (enviado por el form) como 'total_alquiler' (legado)
  var total = parseFloat(data.total || data.total_alquiler) || 0;
  var saldo = parseFloat(data.saldo) || Math.max(0, total - abono);
  // Forzar columnas de fecha como TEXTO (evita conversión a Date con timezone shift)
  var nextRow = ws.getLastRow() + 1;
  ws.getRange(nextRow, 3, 1, 2).setNumberFormat('@'); // FECHA_INICIO y FECHA_FIN como texto
  ws.appendRow([
    id,
    (data.cabana || '').toString().toUpperCase(),
    parseFechaInput(data.fecha_inicio),
    parseFechaInput(data.fecha_fin),
    data.nombre       || '',
    data.cedula       || '',
    data.telefono     || '',
    data.email        || '',
    data.n_personas   || 1,
    JSON.stringify(data.visitantes_extra || []),
    abono,
    total,
    saldo,
    'PENDIENTE',
    data.cliente_frecuente ? 'SI' : 'NO',
    data.comprobante_url || '',
    new Date().toLocaleDateString('es-CO'),
    data.notas || ''
  ]);
  return jsonResponse({ok: true, id: id});
}

// ── RESERVAS: actualizar ──────────────────────────────────────
function handleUpdateReserva(data) {
  var ss   = SpreadsheetApp.openById(SHEET_ID);
  var ws   = ss.getSheetByName('RESERVAS');
  if (!ws) return jsonResponse({error: 'Hoja RESERVAS no encontrada'});
  var rows = ws.getDataRange().getValues();
  var targetId = String(data.id || '').trim();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === targetId) {
      var rowNum = i + 1; // 1-based row

      // === Campos editables completos ===
      // Col 2: CABAÑA
      if (data.cabana !== undefined && data.cabana !== null && data.cabana !== '') {
        ws.getRange(rowNum, 2).setValue(String(data.cabana).toUpperCase());
      }
      // Col 3: FECHA_INICIO — fuerza texto
      if (data.fecha_inicio !== undefined && data.fecha_inicio !== null && data.fecha_inicio !== '') {
        ws.getRange(rowNum, 3).setNumberFormat('@').setValue(parseFechaInput(data.fecha_inicio));
      }
      // Col 4: FECHA_FIN — fuerza texto
      if (data.fecha_fin !== undefined && data.fecha_fin !== null && data.fecha_fin !== '') {
        ws.getRange(rowNum, 4).setNumberFormat('@').setValue(parseFechaInput(data.fecha_fin));
      }
      // Col 5: NOMBRE
      if (data.nombre !== undefined && data.nombre !== null) {
        ws.getRange(rowNum, 5).setValue(data.nombre);
      }
      // Col 6: CEDULA
      if (data.cedula !== undefined && data.cedula !== null) {
        ws.getRange(rowNum, 6).setValue(data.cedula);
      }
      // Col 7: TELEFONO
      if (data.telefono !== undefined && data.telefono !== null) {
        ws.getRange(rowNum, 7).setValue(data.telefono);
      }
      // Col 8: EMAIL
      if (data.email !== undefined && data.email !== null) {
        ws.getRange(rowNum, 8).setValue(data.email);
      }
      // Col 9: N_PERSONAS
      if (data.n_personas !== undefined && data.n_personas !== null) {
        ws.getRange(rowNum, 9).setValue(parseInt(data.n_personas) || 1);
      }
      // Col 10: VISITANTES_EXTRA (JSON string)
      if (data.personas_extra !== undefined && data.personas_extra !== null) {
        ws.getRange(rowNum, 10).setValue(data.personas_extra);
      }
      // Col 14: ESTADO_PAGO
      if (data.estado_pago) ws.getRange(rowNum, 14).setValue(data.estado_pago);

      // === Lógica de montos ===
      var newAbono = (data.abono !== undefined && data.abono !== null)
        ? parseFloat(data.abono) : parseFloat(rows[i][10]) || 0;
      var newTotal = (data.total_alquiler !== undefined && data.total_alquiler !== null)
        ? parseFloat(data.total_alquiler)
        : (data.total !== undefined && data.total !== null)
          ? parseFloat(data.total)
          : parseFloat(rows[i][11]) || 0;

      if (data.abono          !== undefined && data.abono          !== null) ws.getRange(rowNum, 11).setValue(newAbono);
      if (data.total_alquiler !== undefined && data.total_alquiler !== null) ws.getRange(rowNum, 12).setValue(newTotal);
      if (data.total          !== undefined && data.total          !== null) ws.getRange(rowNum, 12).setValue(newTotal);

      // Recalcular saldo salvo que se haya enviado explícitamente
      if (data.saldo !== undefined && data.saldo !== null) {
        ws.getRange(rowNum, 13).setValue(parseFloat(data.saldo));
      } else if ((data.abono !== undefined) || (data.total_alquiler !== undefined) || (data.total !== undefined)) {
        ws.getRange(rowNum, 13).setValue(Math.max(0, newTotal - newAbono));
      }

      // Col 18: NOTAS
      if (data.notas !== undefined && data.notas !== null) ws.getRange(rowNum, 18).setValue(data.notas);

      // ── Registrar movimientos en FLUJO según estado ──
      var flujoMsg = '';

      // VALIDADO → registrar INGRESO (solo si no existe ya)
      if (data.estado_pago === 'VALIDADO' && newAbono > 0) {
        var flujo = ss.getSheetByName('FLUJO');
        if (flujo) {
          var cabana   = data.cabana ? String(data.cabana).toUpperCase() : rows[i][1];
          var nombre   = data.nombre || rows[i][4];
          var fechaIni = data.fecha_inicio ? parseFechaInput(data.fecha_inicio) : parseFechaInput(rows[i][2]);
          var markerIng = 'INGRESO · ID ' + targetId;
          // Buscar duplicado por marcador exacto en NOTAS (col H, index 7)
          var flujoRows = flujo.getDataRange().getValues();
          var alreadyLogged = false;
          for (var k = 3; k < flujoRows.length; k++) {
            var notasFlujo = String(flujoRows[k][7] || '');
            if (notasFlujo.indexOf(markerIng) !== -1) { alreadyLogged = true; break; }
          }
          if (!alreadyLogged) {
            // Calcular el next row y escribir manualmente preservando la fórmula de saldos en col G
            var nextRow = flujo.getLastRow() + 1;
            var prevRow = nextRow - 1;
            // Forzar formato texto en celda de fecha
            flujo.getRange(nextRow, 1).setNumberFormat('@');
            // Escribir A:F (datos) en una sola operación
            flujo.getRange(nextRow, 1, 1, 6).setValues([[fechaIni, newAbono, 'RESERVA - ' + nombre, 'BANCOLOMBIA', cabana, 'INGRESO']]);
            // Col G: fórmula de saldo acumulado (suma valor anterior + valor actual)
            flujo.getRange(nextRow, 7).setFormula('=G' + prevRow + '+B' + nextRow);
            // Col H: notas con marcador
            flujo.getRange(nextRow, 8).setValue(markerIng);
            flujoMsg = 'Ingreso registrado en FLUJO';
          } else {
            flujoMsg = 'Ingreso ya estaba registrado (no se duplicó)';
          }
        }
      }

      // CANCELADO → registrar EGRESO de reembolso si abono existía
      if (data.estado_pago === 'CANCELADO') {
        var abonoOriginal = parseFloat(rows[i][10]) || 0;
        // Calcular % de reembolso desde las notas si están presentes (formato: "Reembolso: X%")
        var pctReembolso = 0;
        var notasStr = String(data.notas || '');
        var pctMatch = notasStr.match(/Reembolso:\s*(\d+(?:\.\d+)?)\s*%/i);
        if (pctMatch) {
          pctReembolso = parseFloat(pctMatch[1]) || 0;
        }
        var montoReembolso = Math.round(abonoOriginal * pctReembolso / 100);

        if (montoReembolso > 0) {
          var flujoEg = ss.getSheetByName('FLUJO');
          if (flujoEg) {
            var cabanaEg = data.cabana ? String(data.cabana).toUpperCase() : rows[i][1];
            var nombreEg = data.nombre || rows[i][4];
            var markerEg = 'REEMBOLSO · ID ' + targetId;
            var flujoRowsEg = flujoEg.getDataRange().getValues();
            var alreadyRefunded = false;
            for (var kk = 3; kk < flujoRowsEg.length; kk++) {
              var notasEg = String(flujoRowsEg[kk][7] || '');
              if (notasEg.indexOf(markerEg) !== -1) { alreadyRefunded = true; break; }
            }
            if (!alreadyRefunded) {
              var hoyStr = new Date().toISOString().slice(0, 10);
              var nextRowEg = flujoEg.getLastRow() + 1;
              var prevRowEg = nextRowEg - 1;
              flujoEg.getRange(nextRowEg, 1).setNumberFormat('@');
              // Escribir A:F en una sola operación
              flujoEg.getRange(nextRowEg, 1, 1, 6).setValues([[hoyStr, -montoReembolso, 'REEMBOLSO - ' + nombreEg, 'BANCOLOMBIA', 'REEMBOLSO', 'EGRESO']]);
              // Col G: fórmula de saldo
              flujoEg.getRange(nextRowEg, 7).setFormula('=G' + prevRowEg + '+B' + nextRowEg);
              // Col H: notas
              flujoEg.getRange(nextRowEg, 8).setValue(markerEg + ' (' + pctReembolso + '%)');
              flujoMsg = 'Reembolso de $' + montoReembolso + ' registrado en FLUJO';
            } else {
              flujoMsg = 'Reembolso ya estaba registrado';
            }
          }
        } else {
          flujoMsg = 'Cancelado sin reembolso (0%)';
        }
      }

      return jsonResponse({ok: true, flujoMsg: flujoMsg});
    }
  }
  return jsonResponse({error: 'Reserva no encontrada: ' + data.id});
}

// Returns "YYYY-MM-DD" STRING — keeps Sheet consistent and avoids any timezone shift
function parseFechaInput(input) {
  if (!input) return '';
  if (input instanceof Date) {
    var y = input.getFullYear();
    var m = String(input.getMonth() + 1);
    if (m.length < 2) m = '0' + m;
    var d = String(input.getDate());
    if (d.length < 2) d = '0' + d;
    return y + '-' + m + '-' + d;
  }
  var str = String(input);
  // Already YYYY-MM-DD?
  var match = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return match[0];
  return str;
}

// ── DISPONIBILIDAD ────────────────────────────────────────────
function handleGetDisponibilidad(data) {
  var ss     = SpreadsheetApp.openById(SHEET_ID);
  var ws     = ss.getSheetByName('RESERVAS');
  var cabana = String(data.cabana || '').toUpperCase();
  if (!ws || ws.getLastRow() <= 1) return jsonResponse({cabana: cabana, ocupadas: []});
  var rows    = ws.getDataRange().getValues();
  var hoy2    = new Date();
  var ocupadas = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toUpperCase() === cabana &&
        rows[i][13] !== 'CANCELADO' &&
        new Date(rows[i][3]) >= hoy2) {
      ocupadas.push({inicio: rows[i][2], fin: rows[i][3], nombre: rows[i][4]});
    }
  }
  return jsonResponse({cabana: cabana, ocupadas: ocupadas});
}

// ── OBSERVABILIDAD ────────────────────────────────────────────
function logObservabilidad(entry) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('LOGS_IA');
  if (!ws) {
    ws = ss.insertSheet('LOGS_IA');
    ws.appendRow(['TIMESTAMP','SESSION_ID','MODELO','INPUT_TOKENS',
      'OUTPUT_TOKENS','LATENCY_MS','COSTO_USD','ERROR']);
    ws.setFrozenRows(1);
  }
  ws.appendRow([
    new Date().toISOString(), entry.sessionId, entry.model,
    entry.inputTokens, entry.outputTokens, entry.latency,
    estimateCost({input_tokens: entry.inputTokens, output_tokens: entry.outputTokens}),
    entry.error || ''
  ]);
}

function estimateCost(usage) {
  var input  = ((usage.input_tokens  || 0) / 1000000) * 3;
  var output = ((usage.output_tokens || 0) / 1000000) * 15;
  return parseFloat((input + output).toFixed(6));
}

// ── HELPERS ───────────────────────────────────────────────────
function calcularPrecio(nPersonas, clienteFrecuente) {
  var n      = parseInt(nPersonas) || 1;
  var precio = n === 1 ? 150000 : 200000 + Math.max(0, n - 2) * 20000;
  if (clienteFrecuente) precio = Math.round(precio * 0.9);
  return {precio: precio, nPersonas: n, descuento: clienteFrecuente ? '10%' : 'ninguno'};
}

function buscarReservas(filtro) {
  var ss   = SpreadsheetApp.openById(SHEET_ID);
  var ws   = ss.getSheetByName('RESERVAS');
  if (!ws || ws.getLastRow() <= 1) return [];
  var rows = ws.getDataRange().getValues();
  var f    = String(filtro || '').toLowerCase();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    for (var j = 0; j < rows[i].length; j++) {
      if (String(rows[i][j]).toLowerCase().indexOf(f) !== -1) {
        result.push({id:rows[i][0], cabana:rows[i][1], inicio:rows[i][2],
          fin:rows[i][3], nombre:rows[i][4], estado:rows[i][13],
          abono:rows[i][10], total:rows[i][11]});
        break;
      }
    }
    if (result.length >= 10) break;
  }
  return result;
}

// ── VERIFY USER ───────────────────────────────────────────────
function handleVerifyUser(data) {
  var email = String(data.email || '').toLowerCase().trim();
  if (!email) return jsonResponse({authorized: false, error: 'email vacio'});
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('USUARIOS');
  if (!ws) return jsonResponse({authorized: false, error: 'Hoja USUARIOS no existe'});
  var rows = ws.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowEmail  = String(rows[i][0] || '').toLowerCase().trim();
    var rowActivo = String(rows[i][3] || '').toUpperCase().trim();
    if (rowEmail === email && rowActivo === 'SI') {
      return jsonResponse({
        authorized: true,
        nombre: rows[i][1] || '',
        rol:    rows[i][2] || 'OPERADOR'
      });
    }
  }
  return jsonResponse({authorized: false});
}

// ── AUTH CHECK ────────────────────────────────────────────────
function isAuthorizedEmail(email) {
  if (!email) return false;
  var normEmail = String(email).toLowerCase().trim();
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName('USUARIOS');
    if (!ws) return false;
    var rows = ws.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var rowEmail  = String(rows[i][0] || '').toLowerCase().trim();
      var rowActivo = String(rows[i][3] || '').toUpperCase().trim();
      if (rowEmail === normEmail && rowActivo === 'SI') return true;
    }
  } catch(e) {}
  return false;
}

// ── READ FLUJO ────────────────────────────────────────────────
function handleReadFlujo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('FLUJO');
  if (!ws) return jsonResponse({rows: []});
  var lastRow = ws.getLastRow();
  if (lastRow < 4) return jsonResponse({rows: []});
  var range = ws.getRange(4, 1, lastRow - 3, 8).getValues();
  var rows = [];
  for (var i = 0; i < range.length; i++) {
    var row = range[i];
    if (!row[0] || (row[1] === '' || row[1] === null)) continue;
    rows.push({
      fecha:  formatDateValue(row[0]),
      valor:  parseFloat(row[1]) || 0,
      desc:   String(row[2] || '').trim(),
      medio:  String(row[3] || '').trim().toUpperCase(),
      tipo:   String(row[4] || '').trim().toUpperCase(),
      ie:     String(row[5] || '').trim().toUpperCase(),
      notas:  String(row[7] || '')
    });
  }
  return jsonResponse({rows: rows});
}

// ── READ RESERVAS ─────────────────────────────────────────────
function handleReadReservas(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('RESERVAS');
  if (!ws || ws.getLastRow() < 2) return jsonResponse({rows: []});
  var range = ws.getRange(2, 1, ws.getLastRow() - 1, 18).getValues();
  var rows  = [];
  for (var i = 0; i < range.length; i++) {
    var r = range[i];
    if (!r[0]) continue;
    rows.push({
      id:        String(r[0]),
      cabana:    String(r[1] || '').toUpperCase(),
      inicio:    formatDateValue(r[2]),
      fin:       formatDateValue(r[3]),
      nombre:    String(r[4]  || ''),
      cedula:    String(r[5]  || ''),
      telefono:  String(r[6]  || ''),
      email:     String(r[7]  || ''),
      nPersonas: parseInt(r[8]) || 1,
      abono:     parseFloat(r[10]) || 0,
      total:     parseFloat(r[11]) || 0,
      saldo:     parseFloat(r[12]) || 0,
      estado:    String(r[13] || 'PENDIENTE').toUpperCase(),
      frecuente: String(r[14] || 'NO').toUpperCase() === 'SI',
      notas:     String(r[17] || '')
    });
  }
  return jsonResponse({rows: rows});
}

// ── READ DASHBOARD ────────────────────────────────────────────
function handleReadDashboard(data) {
  var flujoRes     = handleReadFlujo(data);
  var reservasRes  = handleReadReservas(data);
  var flujoData    = JSON.parse(flujoRes.getContent());
  var reservasData = JSON.parse(reservasRes.getContent());
  return jsonResponse({
    flujo:    flujoData.rows    || [],
    reservas: reservasData.rows || []
  });
}

// ── FORMAT DATE ───────────────────────────────────────────────
function formatDateValue(v) {
  if (!v) return '';
  if (v instanceof Date && !isNaN(v.getTime())) {
    var y = v.getFullYear();
    var m = String(v.getMonth() + 1); if (m.length < 2) m = '0' + m;
    var d = String(v.getDate());      if (d.length < 2) d = '0' + d;
    return y + '-' + m + '-' + d;
  }
  if (typeof v === 'object' && v !== null && typeof v.getFullYear === 'function') {
    var y2 = v.getFullYear();
    var m2 = String(v.getMonth() + 1); if (m2.length < 2) m2 = '0' + m2;
    var d2 = String(v.getDate());       if (d2.length < 2) d2 = '0' + d2;
    return y2 + '-' + m2 + '-' + d2;
  }
  var str = String(v);
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0, 10);
  try {
    var parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      var y3 = parsed.getFullYear();
      var m3 = String(parsed.getMonth() + 1); if (m3.length < 2) m3 = '0' + m3;
      var d3 = String(parsed.getDate());       if (d3.length < 2) d3 = '0' + d3;
      return y3 + '-' + m3 + '-' + d3;
    }
  } catch(e) {}
  return str;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
