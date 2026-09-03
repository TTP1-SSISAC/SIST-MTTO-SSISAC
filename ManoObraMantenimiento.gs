// ============================================================
//  ManoObraMantenimiento.gs  —  SIST-MTTO-SSISAC
//  Registro de mano de obra por Orden de Trabajo
// ============================================================

var SHEET_MO = 'MANO_OBRA_OT';

var MC_MO = {
  ID:           0,   // MO-YYYY-NNNN
  FECHA_REG:    1,
  ID_OT:        2,
  ID_VEHICULO:  3,
  PLACA:        4,
  TECNICO:      5,
  CARGO:        6,   // Mecánico | Electricista | Especialista
  FECHA_INICIO: 7,
  HORA_INICIO:  8,   // HH:MM (texto)
  FECHA_FIN:    9,
  HORA_FIN:     10,
  HORAS_TOTALES:11,
  TARIFA_HORA:  12,
  COSTO_TOTAL:  13,
  TIPO_HRS:     14,  // NORMALES | EXTRAS | GUARDIA
  DESCRIPCION:  15,
  OBSERVACIONES:16,
  USUARIO:      17,
  TIMESTAMP:    18
};

var CABECERA_MO = [
  'ID_MO','FECHA_REGISTRO','ID_OT','ID_VEHICULO','PLACA',
  'TECNICO','CARGO','FECHA_INICIO','HORA_INICIO','FECHA_FIN','HORA_FIN',
  'HORAS_TOTALES','TARIFA_HORA','COSTO_TOTAL','TIPO_HORAS',
  'DESCRIPCION_TRABAJO','OBSERVACIONES','USUARIO','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetMO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MO);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MO);
    sh.appendRow(CABECERA_MO);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_MO.length)
      .setBackground('#3A0A4A').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdMO() {
  var sh   = _sheetMO();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'MO-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('MO-' + yr); });
  if (!ids.length) return 'MO-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'MO-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _rowToMO(row) {
  return {
    id:           row[MC_MO.ID],
    fechaReg:     row[MC_MO.FECHA_REG],
    idOT:         row[MC_MO.ID_OT],
    idVehiculo:   row[MC_MO.ID_VEHICULO],
    placa:        row[MC_MO.PLACA],
    tecnico:      row[MC_MO.TECNICO],
    cargo:        row[MC_MO.CARGO],
    fechaInicio:  row[MC_MO.FECHA_INICIO],
    horaInicio:   row[MC_MO.HORA_INICIO],
    fechaFin:     row[MC_MO.FECHA_FIN],
    horaFin:      row[MC_MO.HORA_FIN],
    horasTotales: row[MC_MO.HORAS_TOTALES],
    tarifaHora:   row[MC_MO.TARIFA_HORA],
    costoTotal:   row[MC_MO.COSTO_TOTAL],
    tipoHrs:      row[MC_MO.TIPO_HRS],
    descripcion:  row[MC_MO.DESCRIPCION],
    observaciones:row[MC_MO.OBSERVACIONES],
    usuario:      row[MC_MO.USUARIO],
    timestamp:    row[MC_MO.TIMESTAMP]
  };
}

// ─────────────────────────────────────────────────────────────
//  calcularHoras(fechaInicio, horaInicio, fechaFin, horaFin)
//  Retorna horas con 1 decimal
// ─────────────────────────────────────────────────────────────
function calcularHoras(fechaInicio, horaInicio, fechaFin, horaFin) {
  try {
    if (!fechaInicio || !horaInicio || !fechaFin || !horaFin) {
      throw new Error('Se requieren fechas y horas de inicio y fin');
    }
    var ini = new Date(fechaInicio + 'T' + horaInicio + ':00');
    var fin = new Date(fechaFin    + 'T' + horaFin    + ':00');
    if (isNaN(ini.getTime()) || isNaN(fin.getTime())) {
      throw new Error('Formato de fecha u hora inválido (use yyyy-MM-dd y HH:MM)');
    }
    var horas = Math.round((fin - ini) / 3600000 * 10) / 10;
    if (horas < 0) throw new Error('La fecha/hora de fin debe ser posterior al inicio');
    return { ok: true, data: { horas: horas } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularCostoManoObra(horas, tarifaHora, tipoHrs)
//  Aplica recargo de 1.5× en extras y 2× en guardia
// ─────────────────────────────────────────────────────────────
function calcularCostoManoObra(horas, tarifaHora, tipoHrs) {
  try {
    horas      = parseFloat(horas)      || 0;
    tarifaHora = parseFloat(tarifaHora) || 0;
    tipoHrs    = String(tipoHrs || 'NORMALES').toUpperCase();

    var factor = 1;
    if (tipoHrs === 'EXTRAS')  factor = 1.5;
    if (tipoHrs === 'GUARDIA') factor = 2.0;

    var costo = Math.round(horas * tarifaHora * factor * 100) / 100;
    return { ok: true, data: { horas: horas, tarifaHora: tarifaHora, factor: factor, costo: costo } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  registrarManoObra(params)
//  params: idOT, idVehiculo, placa, tecnico, cargo,
//          fechaInicio, horaInicio, fechaFin, horaFin,
//          tarifaHora, tipoHrs, descripcion, usuario
// ─────────────────────────────────────────────────────────────
function registrarManoObra(params) {
  try {
    if (!params.idOT)   throw new Error('ID de OT requerido');
    if (!params.tecnico) throw new Error('Técnico requerido');

    var horas     = 0;
    var costoTotal = 0;

    // Calcular horas si se proveen fechas/horas
    if (params.fechaInicio && params.horaInicio && params.fechaFin && params.horaFin) {
      var hRes = calcularHoras(params.fechaInicio, params.horaInicio,
                               params.fechaFin,    params.horaFin);
      if (!hRes.ok) throw new Error(hRes.error);
      horas = hRes.data.horas;
    } else {
      horas = parseFloat(params.horasTotales) || 0;
    }

    // Calcular costo
    if (horas && params.tarifaHora) {
      var cRes = calcularCostoManoObra(horas, params.tarifaHora, params.tipoHrs);
      if (cRes.ok) costoTotal = cRes.data.costo;
    }

    var sh  = _sheetMO();
    var id  = _nextIdMO();
    var now = new Date();
    var hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    sh.appendRow([
      id,
      hoy,
      params.idOT,
      params.idVehiculo  || '',
      String(params.placa || '').toUpperCase(),
      params.tecnico,
      params.cargo       || 'Mecánico',
      params.fechaInicio || hoy,
      params.horaInicio  || '',
      params.fechaFin    || '',
      params.horaFin     || '',
      horas,
      parseFloat(params.tarifaHora) || 0,
      costoTotal,
      String(params.tipoHrs || 'NORMALES').toUpperCase(),
      params.descripcion   || '',
      params.observaciones || '',
      params.usuario || 'SISTEMA',
      now.toISOString()
    ]);

    return { ok: true, data: { id: id, horas: horas, costoTotal: costoTotal,
                                mensaje: 'Mano de obra registrada: ' + id } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  obtenerManoObraPorOT(idOT)
// ─────────────────────────────────────────────────────────────
function obtenerManoObraPorOT(idOT) {
  try {
    if (!idOT) throw new Error('ID de OT requerido');
    var data      = _sheetMO().getDataRange().getValues();
    var resultado = [];
    var totalH    = 0;
    var totalC    = 0;

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_MO.ID_OT] == idOT) {
        var mo = _rowToMO(data[i]);
        totalH += parseFloat(mo.horasTotales) || 0;
        totalC += parseFloat(mo.costoTotal)   || 0;
        resultado.push(mo);
      }
    }

    return { ok: true, data: resultado, total: resultado.length,
             resumen: { horasTotales: Math.round(totalH * 10) / 10,
                        costoTotal:   Math.round(totalC * 100) / 100 } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
