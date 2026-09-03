// ============================================================
//  Intervenciones.gs  —  SIST-MTTO-SSISAC
//  Registro de intervenciones técnicas dentro de una OT
// ============================================================

var SHEET_INT = 'INTERVENCIONES';

var MC_INT = {
  ID:           0,   // INT-YYYY-NNNN
  FECHA_REG:    1,
  ID_OT:        2,
  ID_VEHICULO:  3,
  PLACA:        4,
  SISTEMA:      5,   // Motor | Frenos | Suspensión | ...
  COMPONENTE:   6,
  FALLA:        7,
  ACCION:       8,   // Descripción de lo que se hizo
  ESTADO:       9,   // ABIERTA | EN_PROCESO | CERRADA
  TECNICO:      10,
  FECHA_INICIO: 11,
  FECHA_FIN:    12,
  DURACION_H:   13,
  OBSERVACIONES:14,
  USUARIO:      15,
  TIMESTAMP:    16
};

var CABECERA_INT = [
  'ID_INTERVENCION','FECHA_REGISTRO','ID_OT','ID_VEHICULO','PLACA',
  'SISTEMA','COMPONENTE','DESCRIPCION_FALLA','ACCION_REALIZADA',
  'ESTADO','TECNICO','FECHA_INICIO','FECHA_FIN','DURACION_HORAS',
  'OBSERVACIONES','USUARIO','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetInt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_INT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_INT);
    sh.appendRow(CABECERA_INT);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_INT.length)
      .setBackground('#4A1D00').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdInt() {
  var sh   = _sheetInt();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'INT-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('INT-' + yr); });
  if (!ids.length) return 'INT-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'INT-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _rowToInt(row) {
  return {
    id:           row[MC_INT.ID],
    fechaReg:     row[MC_INT.FECHA_REG],
    idOT:         row[MC_INT.ID_OT],
    idVehiculo:   row[MC_INT.ID_VEHICULO],
    placa:        row[MC_INT.PLACA],
    sistema:      row[MC_INT.SISTEMA],
    componente:   row[MC_INT.COMPONENTE],
    falla:        row[MC_INT.FALLA],
    accion:       row[MC_INT.ACCION],
    estado:       row[MC_INT.ESTADO],
    tecnico:      row[MC_INT.TECNICO],
    fechaInicio:  row[MC_INT.FECHA_INICIO],
    fechaFin:     row[MC_INT.FECHA_FIN],
    duracionH:    row[MC_INT.DURACION_H],
    observaciones:row[MC_INT.OBSERVACIONES],
    usuario:      row[MC_INT.USUARIO],
    timestamp:    row[MC_INT.TIMESTAMP]
  };
}

// ─────────────────────────────────────────────────────────────
//  registrarIntervencion(params)
//  params: idOT, idVehiculo, placa, sistema, componente,
//          falla, accion, tecnico, fechaInicio, usuario
// ─────────────────────────────────────────────────────────────
function registrarIntervencion(params) {
  try {
    if (!params.idOT)      throw new Error('ID de OT requerido');
    if (!params.placa)     throw new Error('Placa requerida');
    if (!params.sistema)   throw new Error('Sistema requerido');
    if (!params.componente) throw new Error('Componente requerido');

    var sh  = _sheetInt();
    var id  = _nextIdInt();
    var now = new Date();
    var hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    sh.appendRow([
      id,
      hoy,
      params.idOT,
      params.idVehiculo || '',
      String(params.placa).toUpperCase(),
      params.sistema,
      params.componente,
      params.falla       || '',
      params.accion      || '',
      'ABIERTA',
      params.tecnico     || '',
      params.fechaInicio || hoy,
      '',
      '',
      params.observaciones || '',
      params.usuario || 'SISTEMA',
      now.toISOString()
    ]);

    return { ok: true, data: { id: id, mensaje: 'Intervención registrada: ' + id } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  obtenerIntervencionesPorOT(idOT)
// ─────────────────────────────────────────────────────────────
function obtenerIntervencionesPorOT(idOT) {
  try {
    if (!idOT) throw new Error('ID de OT requerido');
    var data      = _sheetInt().getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_INT.ID_OT] == idOT) {
        resultado.push(_rowToInt(data[i]));
      }
    }
    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  cerrarIntervencion(params)
//  params: id, accion, fechaFin, observaciones, usuario
// ─────────────────────────────────────────────────────────────
function cerrarIntervencion(params) {
  try {
    if (!params.id) throw new Error('ID requerido');
    var sh   = _sheetInt();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_INT.ID] == params.id) {
        var fila   = i + 1;
        var hoy    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var fFin   = params.fechaFin || hoy;
        var fIni   = data[i][MC_INT.FECHA_INICIO];

        // Calcular duración
        var duracion = '';
        if (fIni && fFin) {
          var d1  = new Date(fIni);
          var d2  = new Date(fFin);
          var hrs = Math.round((d2 - d1) / 3600000 * 10) / 10;
          if (hrs >= 0) duracion = hrs;
        }

        sh.getRange(fila, MC_INT.ESTADO      + 1).setValue('CERRADA');
        sh.getRange(fila, MC_INT.FECHA_FIN   + 1).setValue(fFin);
        sh.getRange(fila, MC_INT.DURACION_H  + 1).setValue(duracion);
        if (params.accion) sh.getRange(fila, MC_INT.ACCION + 1).setValue(params.accion);
        if (params.observaciones) sh.getRange(fila, MC_INT.OBSERVACIONES + 1).setValue(params.observaciones);
        sh.getRange(fila, MC_INT.TIMESTAMP   + 1).setValue(new Date().toISOString());

        return { ok: true, data: { id: params.id, duracionH: duracion, mensaje: 'Intervención cerrada' } };
      }
    }
    throw new Error('Intervención no encontrada: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
