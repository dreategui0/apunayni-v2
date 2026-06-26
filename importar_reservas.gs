/**
 * importar_reservas.gs
 * Ejecutar UNA sola vez desde el editor de Apps Script:
 *   Selecciona "importarReservasHistoricas" → clic Ejecutar
 *
 * Importa todas las reservas de Junio–Agosto 2026 extraídas
 * del historial de disponibilidad. Sin cédula, teléfono ni abono
 * (se completan luego desde la app con el botón Editar).
 */

function importarReservasHistoricas() {
  var ss = SpreadsheetApp.openById('1ueAv6JhzTYhJ7y3ug-N6-IbPBOfuuGL4rUKUE1WuHaQ');
  var ws = ss.getSheetByName('RESERVAS');

  // Crear hoja si no existe
  if (!ws) {
    ws = ss.insertSheet('RESERVAS');
    ws.appendRow(['ID','CABANA','FECHA_INICIO','FECHA_FIN','NOMBRE','CEDULA',
      'TELEFONO','EMAIL','N_PERSONAS','VISITANTES_EXTRA','ABONO',
      'TOTAL_ALQUILER','SALDO','ESTADO_PAGO','CLIENTE_FRECUENTE',
      'COMPROBANTE_URL','FECHA_REGISTRO','NOTAS']);
    ws.setFrozenRows(1);
  }

  // Verificar si ya hay datos (evitar duplicados)
  if (ws.getLastRow() > 1) {
    var confirmar = Browser.msgBox(
      'Advertencia',
      'La hoja RESERVAS ya tiene datos. ¿Agregar igualmente?',
      Browser.Buttons.YES_NO
    );
    if (confirmar !== 'yes') return;
  }

  var hoy = new Date();
  var fechaReg = Utilities.formatDate(hoy, 'America/Bogota', 'dd/MM/yyyy');

  // Estado según si la reserva ya terminó
  function estado(fechaFin) {
    return new Date(fechaFin) < hoy ? 'VALIDADO' : 'PENDIENTE';
  }

  // [ cabana, fecha_inicio, fecha_fin, nombre ]
  var reservas = [
    // ── APUNAYNI ──────────────────────────────────────────────
    ['APUNAYNI', '2026-06-03', '2026-06-04', 'Edward'],
    ['APUNAYNI', '2026-06-04', '2026-06-05', 'Mateo'],
    ['APUNAYNI', '2026-06-06', '2026-06-08', 'Andreita'],
    ['APUNAYNI', '2026-06-14', '2026-06-15', 'Marcela'],
    ['APUNAYNI', '2026-06-15', '2026-06-16', 'Manuela'],
    ['APUNAYNI', '2026-06-17', '2026-06-18', 'Daniel Guia'],
    ['APUNAYNI', '2026-06-21', '2026-06-22', 'Ana Tapias'],
    ['APUNAYNI', '2026-06-22', '2026-06-23', 'Lorena'],
    ['APUNAYNI', '2026-06-23', '2026-06-25', 'Monica Hernandez'],
    ['APUNAYNI', '2026-06-25', '2026-06-27', 'Claudia Marcela'],
    ['APUNAYNI', '2026-06-27', '2026-06-28', 'Julieth'],
    ['APUNAYNI', '2026-06-28', '2026-06-29', 'Caterine Aguirre'],
    ['APUNAYNI', '2026-07-03', '2026-07-04', 'Karen'],
    ['APUNAYNI', '2026-07-04', '2026-07-05', 'Daniela'],
    ['APUNAYNI', '2026-07-05', '2026-07-06', 'Diana Zuluaga'],
    ['APUNAYNI', '2026-07-10', '2026-07-11', 'Monica'],
    ['APUNAYNI', '2026-07-11', '2026-07-12', 'Natalia González'],
    ['APUNAYNI', '2026-07-12', '2026-07-13', 'Joan Reyes'],
    ['APUNAYNI', '2026-07-16', '2026-07-18', 'LeZ'],
    ['APUNAYNI', '2026-07-18', '2026-07-20', 'Diana Fernández'],
    ['APUNAYNI', '2026-07-25', '2026-07-26', 'Catalina Salazar'],
    ['APUNAYNI', '2026-08-01', '2026-08-02', 'Deisi'],
    ['APUNAYNI', '2026-08-07', '2026-08-09', 'Monica'],
    ['APUNAYNI', '2026-08-15', '2026-08-16', 'Camilo Gomes'],
    ['APUNAYNI', '2026-08-16', '2026-08-17', 'Duverney'],
    ['APUNAYNI', '2026-08-22', '2026-08-23', 'Alejandra Ceballos'],
    ['APUNAYNI', '2026-08-29', '2026-08-30', 'Jennifer'],
    // ── SALVAWASI ─────────────────────────────────────────────
    ['SALVAWASI', '2026-06-05', '2026-06-07', 'Yolanda Preciado'],
    ['SALVAWASI', '2026-06-07', '2026-06-08', 'Melanie Fix'],
    ['SALVAWASI', '2026-06-12', '2026-06-13', 'Andy'],
    ['SALVAWASI', '2026-06-13', '2026-06-14', 'Simon'],
    ['SALVAWASI', '2026-06-16', '2026-06-18', 'Deisi J'],
    ['SALVAWASI', '2026-06-19', '2026-06-20', 'Daniel'],
    ['SALVAWASI', '2026-06-20', '2026-06-22', 'Juan Torres'],
    ['SALVAWASI', '2026-06-22', '2026-06-23', 'Ana Tapia'],
    ['SALVAWASI', '2026-06-23', '2026-06-24', 'Mariana Conde'],
    ['SALVAWASI', '2026-06-26', '2026-06-27', 'Idali'],
    ['SALVAWASI', '2026-06-27', '2026-06-28', 'Luis Hernández'],
    ['SALVAWASI', '2026-06-28', '2026-06-29', 'Andrew'],
    ['SALVAWASI', '2026-07-03', '2026-07-04', 'Edwin Lopez'],
    ['SALVAWASI', '2026-07-04', '2026-07-05', 'Fernanda'],
    ['SALVAWASI', '2026-07-11', '2026-07-12', 'Suly'],
    ['SALVAWASI', '2026-07-18', '2026-07-20', 'Juli'],
    ['SALVAWASI', '2026-07-25', '2026-07-26', 'Elian Gomez'],
    ['SALVAWASI', '2026-07-31', '2026-08-02', 'Kt Castro'],
    ['SALVAWASI', '2026-08-02', '2026-08-03', 'Hernández Marcel'],
    ['SALVAWASI', '2026-08-07', '2026-08-09', 'Melissa'],
    ['SALVAWASI', '2026-08-15', '2026-08-16', 'Luisa Fernanda']
  ];

  var contador = 0;
  reservas.forEach(function(r, idx) {
    var id       = 'RES-' + String(idx + 1).padStart(3, '0');
    var cabana   = r[0];
    var fechaIni = r[1];
    var fechaFin = r[2];
    var nombre   = r[3];
    var est      = estado(fechaFin);

    ws.appendRow([
      id,          // A: ID
      cabana,      // B: CABANA
      fechaIni,    // C: FECHA_INICIO
      fechaFin,    // D: FECHA_FIN
      nombre,      // E: NOMBRE
      '',          // F: CEDULA
      '',          // G: TELEFONO
      '',          // H: EMAIL
      1,           // I: N_PERSONAS
      '[]',        // J: VISITANTES_EXTRA
      0,           // K: ABONO
      0,           // L: TOTAL_ALQUILER
      0,           // M: SALDO
      est,         // N: ESTADO_PAGO
      'NO',        // O: CLIENTE_FRECUENTE
      '',          // P: COMPROBANTE_URL
      fechaReg,    // Q: FECHA_REGISTRO
      'Importado desde historial de disponibilidad'  // R: NOTAS
    ]);
    contador++;
  });

  SpreadsheetApp.flush();
  Browser.msgBox('✅ Importación completa: ' + contador + ' reservas agregadas.\n\nRecuerda completar cédula, teléfono y montos usando el botón Editar en la app.');
}
