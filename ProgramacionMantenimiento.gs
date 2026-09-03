// ============================================================
//  ProgramacionMantenimiento.gs  —  SIST-MTTO-SSISAC
//  Generación y seguimiento de programaciones periódicas
// ============================================================

var SHEET_PROG = 'PROGRAMACION_MTTO';

var MC_PROG = {
  ID:            0,   // PROG-YYYY-NNNN
  FECHA_REG:     1,
  ID_PLAN:       2,   // Referencia al plan preventivo
  ID_VEHICULO:   3,
  PLACA:         4,
  TIPO_SERVICIO: 5,
  DESCRIPCION:   6,
  FECHA_PROG:    7,   // Fecha programada
  KM_PROG:       8,   // KM programado
  ESTADO:        9,   // PENDIENTE | REALIZADO | VENCIDO | REPROGRAMADO
  FECHA_REAL:    10,  // Fecha en que se realizó
  KM_REAL:       11,
  ID_OT:         12,  // OT generada
  TECNICO:       13,
  OBSERVACIONES: 14,
  USUARIO:       15,
  TIMESTAMP:     16
};

var CABECERA_PROG = [
  'ID_PROG','FECHA_REGISTRO','ID_PLAN','ID_VEHICULO','PLACA',
  'TIPO_SERVICIO','DESCRIPCION','FECHA_PROGRAMADA','KM_PROGRAMADO',
  'ESTADO','FECHA_REALIZADO','KM_REALIZADO','ID_OT',
  'TECNICO','OBSERVACIONES','USUARIO','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetProg() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PROG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PROG);
    sh.appendRow(CABECERA_PROG);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_PROG.length)
      .setBackground('#003B1A').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdProg() {
  var sh   = _sheetProg();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'PROG-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('PROG-' + yr); });
  if (!ids.length) return 'PROG-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'PROG-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _rowToProg(row) {
  return {
    id:           row[MC_PROG.ID],
    fechaReg:     row[MC_PROG.FECHA_REG],
    idPlan:       row[MC_PROG.ID_PLAN],
    idVehiculo:   row[MC_PROG.ID_VEHICULO],
    placa:        row[MC_PROG.PLACA],
    tipoServicio: row[MC_PROG.TIPO_SERVICIO],
    descripcion:  row[MC_PROG.DESCRIPCION],
    fechaProg:    row[MC_PROG.FECHA_PROG],
    kmProg:       row[MC_PROG.KM_PROG],
    estado:       row[MC_PROG.ESTADO],
    fechaReal:    row[MC_PROG.FECHA_REAL],
    kmReal:       row[MC_PROG.KM_REAL],
    idOT:         row[MC_PROG.ID_OT],
    tecnico:      row[MC_PROG.TECNICO],
    observaciones:row[MC_PROG.OBSERVACIONES],
    usuario:      row[MC_PROG.USUARIO],
    timestamp:    row[MC_PROG.TIMESTAMP]
  };
}

// ─────────────────────────────────────────────────────────────
//  generarProgramaciones(placa, mesesAdelante)
//  Genera programaciones futuras a partir de los planes activos
// ─────────────────────────────────────────────────────────────
function generarProgramaciones(placa, mesesAdelante) {
  try {
    mesesAdelante = mesesAdelante || 3;

    // Obtener planes activos
    var planesRes = listarPlanesPorVehiculo(placa);
    if (!planesRes.ok) throw new Error(planesRes.error);

    var sh        = _sheetProg();
    var hoy       = new Date();
    var limite    = new Date();
    limite.setMonth(limite.getMonth() + mesesAdelante);
    var generadas = [];

    planesRes.data.forEach(function(plan) {
      if (plan.estado !== 'ACTIVO') return;
      if (!plan.intervaloDias || !plan.fechaProxima) return;

      // Verificar si ya existe una programación pendiente para este plan
      var existente = _buscarProgPendientePorPlan(plan.id);
      if (existente) return; // Ya hay una programación activa

      var fechaProg = new Date(plan.fechaProxima);
      if (isNaN(fechaProg.getTime())) return;
      if (fechaProg > limite) return; // Muy lejos en el futuro

      var id  = _nextIdProg();
      var now = new Date();

      sh.appendRow([
        id,
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        plan.id,
        plan.idVehiculo,
        plan.placa,
        plan.tipoServicio,
        plan.descripcion,
        Utilities.formatDate(fechaProg, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        plan.kmProximo || '',
        'PENDIENTE',
        '', '', '',
        plan.tecnico || '',
        '',
        'SISTEMA',
        now.toISOString()
      ]);

      generadas.push({ id: id, tipoServicio: plan.tipoServicio, fechaProg: plan.fechaProxima });
    });

    return { ok: true, data: generadas, total: generadas.length,
             mensaje: 'Se generaron ' + generadas.length + ' programaciones' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function _buscarProgPendientePorPlan(idPlan) {
  var data = _sheetProg().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][MC_PROG.ID_PLAN] === idPlan && data[i][MC_PROG.ESTADO] === 'PENDIENTE') {
      return _rowToProg(data[i]);
    }
  }
  return null;
}

// ─────────────────────────────────────────────────────────────
//  obtenerProximosMantenimientos(filtros)
//  filtros: placa, diasAdelante (default 30), tipoServicio
// ─────────────────────────────────────────────────────────────
function obtenerProximosMantenimientos(filtros) {
  try {
    filtros       = filtros || {};
    var diasAd    = filtros.diasAdelante || 30;
    var hoy       = new Date();
    var limite    = new Date();
    limite.setDate(limite.getDate() + diasAd);

    var data      = _sheetProg().getDataRange().getValues();
    var resultado = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[MC_PROG.ID]) continue;
      if (row[MC_PROG.ESTADO] !== 'PENDIENTE') continue;

      var fProg = new Date(row[MC_PROG.FECHA_PROG]);
      if (isNaN(fProg.getTime())) continue;
      if (fProg > limite) continue;

      if (filtros.placa && row[MC_PROG.PLACA] !== filtros.placa.toUpperCase()) continue;
      if (filtros.tipoServicio && row[MC_PROG.TIPO_SERVICIO] !== filtros.tipoServicio.toUpperCase()) continue;

      var prog       = _rowToProg(row);
      var diasRest   = Math.ceil((fProg - hoy) / 86400000);
      prog.diasRestantes = diasRest;
      prog.vencido       = diasRest < 0;

      resultado.push(prog);
    }

    resultado.sort(function(a, b) { return a.diasRestantes - b.diasRestantes; });
    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  marcarComoRealizado(params)
//  params: id, fechaReal, kmReal, idOT, usuario, observaciones
// ─────────────────────────────────────────────────────────────
function marcarComoRealizado(params) {
  try {
    if (!params.id) throw new Error('ID requerido');
    var sh   = _sheetProg();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_PROG.ID] == params.id) {
        var fila = i + 1;
        var hoy  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

        sh.getRange(fila, MC_PROG.ESTADO     + 1).setValue('REALIZADO');
        sh.getRange(fila, MC_PROG.FECHA_REAL + 1).setValue(params.fechaReal || hoy);
        sh.getRange(fila, MC_PROG.KM_REAL    + 1).setValue(params.kmReal    || '');
        sh.getRange(fila, MC_PROG.ID_OT      + 1).setValue(params.idOT      || '');
        if (params.observaciones) sh.getRange(fila, MC_PROG.OBSERVACIONES + 1).setValue(params.observaciones);
        sh.getRange(fila, MC_PROG.TIMESTAMP  + 1).setValue(new Date().toISOString());

        // Actualizar el plan con el nuevo kmUltimo y fechaUltimo
        var idPlan = data[i][MC_PROG.ID_PLAN];
        if (idPlan) {
          actualizarPlanMantenimiento({
            id:          idPlan,
            kmUltimo:    params.kmReal    || data[i][MC_PROG.KM_PROG],
            fechaUltimo: params.fechaReal || hoy
          });
        }

        return { ok: true, data: { id: params.id, mensaje: 'Programación marcada como realizada' } };
      }
    }
    throw new Error('Programación no encontrada: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  reprogramarMantenimiento(params)
//  params: id, nuevaFecha, motivo, usuario
// ─────────────────────────────────────────────────────────────
function reprogramarMantenimiento(params) {
  try {
    if (!params.id)        throw new Error('ID requerido');
    if (!params.nuevaFecha) throw new Error('Nueva fecha requerida');

    var sh   = _sheetProg();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_PROG.ID] == params.id) {
        var fila    = i + 1;
        var motivo  = 'Reprogramado: ' + (params.motivo || 'Sin motivo especificado');
        var obs     = data[i][MC_PROG.OBSERVACIONES];
        var obsNueva = obs ? obs + ' | ' + motivo : motivo;

        sh.getRange(fila, MC_PROG.FECHA_PROG    + 1).setValue(params.nuevaFecha);
        sh.getRange(fila, MC_PROG.ESTADO        + 1).setValue('REPROGRAMADO');
        sh.getRange(fila, MC_PROG.OBSERVACIONES + 1).setValue(obsNueva);
        sh.getRange(fila, MC_PROG.TIMESTAMP     + 1).setValue(new Date().toISOString());

        return { ok: true, data: { id: params.id, nuevaFecha: params.nuevaFecha,
                                   mensaje: 'Mantenimiento reprogramado' } };
      }
    }
    throw new Error('Programación no encontrada: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  detectarVencidos()
//  Marca como VENCIDO todas las programaciones pendientes
//  cuya fecha ya pasó. Diseñado para correr como trigger diario.
// ─────────────────────────────────────────────────────────────
function detectarVencidos() {
  try {
    var sh       = _sheetProg();
    var data     = sh.getDataRange().getValues();
    var hoy      = new Date();
    var vencidos = 0;

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_PROG.ESTADO] !== 'PENDIENTE') continue;
      var fProg = new Date(data[i][MC_PROG.FECHA_PROG]);
      if (isNaN(fProg.getTime())) continue;
      if (fProg < hoy) {
        sh.getRange(i + 1, MC_PROG.ESTADO    + 1).setValue('VENCIDO');
        sh.getRange(i + 1, MC_PROG.TIMESTAMP + 1).setValue(new Date().toISOString());
        vencidos++;
      }
    }

    return { ok: true, data: { vencidos: vencidos, mensaje: vencidos + ' programaciones marcadas como vencidas' } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
