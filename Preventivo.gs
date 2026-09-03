// ============================================================
//  Preventivo.gs  —  SIST-MTTO-SSISAC
//  Planes de Mantenimiento Preventivo
// ============================================================

var SHEET_PREV = 'PLANES_PREVENTIVOS';

var MC_PREV = {
  ID:            0,   // PLAN-YYYY-NNNN
  FECHA_REG:     1,
  ID_VEHICULO:   2,
  PLACA:         3,
  TIPO_SERVICIO: 4,   // ACEITE | FILTROS | FRENOS | LLANTAS | REVISION_GENERAL | ...
  DESCRIPCION:   5,
  INTERVALO_KM:  6,   // Cada cuántos KM
  INTERVALO_DIAS:7,   // Cada cuántos días
  KM_ULTIMO:     8,   // KM en el último servicio
  FECHA_ULTIMO:  9,   // Fecha del último servicio
  KM_PROXIMO:    10,  // KM_ULTIMO + INTERVALO_KM
  FECHA_PROXIMA: 11,  // FECHA_ULTIMO + INTERVALO_DIAS
  ESTADO:        12,  // ACTIVO | INACTIVO | VENCIDO
  TECNICO:       13,
  OBSERVACIONES: 14,
  USUARIO:       15,
  TIMESTAMP:     16
};

var CABECERA_PREV = [
  'ID_PLAN','FECHA_REGISTRO','ID_VEHICULO','PLACA','TIPO_SERVICIO',
  'DESCRIPCION','INTERVALO_KM','INTERVALO_DIAS','KM_ULTIMO_SERVICIO',
  'FECHA_ULTIMO_SERVICIO','KM_PROXIMO_SERVICIO','FECHA_PROXIMA_SERVICIO',
  'ESTADO','TECNICO_ASIGNADO','OBSERVACIONES','USUARIO','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetPrev() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PREV);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PREV);
    sh.appendRow(CABECERA_PREV);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_PREV.length)
      .setBackground('#1A3C00').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdPrev() {
  var sh   = _sheetPrev();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'PLAN-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('PLAN-' + yr); });
  if (!ids.length) return 'PLAN-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'PLAN-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _rowToPrev(row) {
  return {
    id:           row[MC_PREV.ID],
    fechaReg:     row[MC_PREV.FECHA_REG],
    idVehiculo:   row[MC_PREV.ID_VEHICULO],
    placa:        row[MC_PREV.PLACA],
    tipoServicio: row[MC_PREV.TIPO_SERVICIO],
    descripcion:  row[MC_PREV.DESCRIPCION],
    intervaloKm:  row[MC_PREV.INTERVALO_KM],
    intervaloDias:row[MC_PREV.INTERVALO_DIAS],
    kmUltimo:     row[MC_PREV.KM_ULTIMO],
    fechaUltimo:  row[MC_PREV.FECHA_ULTIMO],
    kmProximo:    row[MC_PREV.KM_PROXIMO],
    fechaProxima: row[MC_PREV.FECHA_PROXIMA],
    estado:       row[MC_PREV.ESTADO],
    tecnico:      row[MC_PREV.TECNICO],
    observaciones:row[MC_PREV.OBSERVACIONES],
    usuario:      row[MC_PREV.USUARIO],
    timestamp:    row[MC_PREV.TIMESTAMP]
  };
}

// ─────────────────────────────────────────────────────────────
//  crearPlanMantenimiento(params)
//  params: idVehiculo, placa, tipoServicio, descripcion,
//          intervaloKm, intervaloDias, kmUltimo, fechaUltimo,
//          tecnico, usuario
// ─────────────────────────────────────────────────────────────
function crearPlanMantenimiento(params) {
  try {
    if (!params.placa)       throw new Error('Placa requerida');
    if (!params.tipoServicio) throw new Error('Tipo de servicio requerido');

    var sh   = _sheetPrev();
    var id   = _nextIdPrev();
    var now  = new Date();
    var hoy  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    var kmUltimo  = parseFloat(params.kmUltimo)   || 0;
    var intKm     = parseFloat(params.intervaloKm) || 0;
    var intDias   = parseInt(params.intervaloDias) || 0;
    var kmProximo = intKm ? kmUltimo + intKm : '';

    // Calcular fecha próxima
    var fechaProxima = '';
    var fechaUltimo  = params.fechaUltimo || hoy;
    if (intDias > 0) {
      var base = new Date(fechaUltimo);
      base.setDate(base.getDate() + intDias);
      fechaProxima = Utilities.formatDate(base, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    sh.appendRow([
      id,
      hoy,
      params.idVehiculo || '',
      String(params.placa).toUpperCase(),
      String(params.tipoServicio).toUpperCase(),
      params.descripcion || '',
      intKm,
      intDias,
      kmUltimo,
      fechaUltimo,
      kmProximo,
      fechaProxima,
      'ACTIVO',
      params.tecnico || '',
      params.observaciones || '',
      params.usuario || 'SISTEMA',
      now.toISOString()
    ]);

    return { ok: true, data: { id: id, kmProximo: kmProximo, fechaProxima: fechaProxima,
                                mensaje: 'Plan de mantenimiento creado: ' + id } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  obtenerPlanMantenimiento(id)
// ─────────────────────────────────────────────────────────────
function obtenerPlanMantenimiento(id) {
  try {
    if (!id) throw new Error('ID requerido');
    var data = _sheetPrev().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_PREV.ID] == id) {
        return { ok: true, data: _rowToPrev(data[i]) };
      }
    }
    throw new Error('Plan no encontrado: ' + id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  actualizarPlanMantenimiento(params)
//  Actualiza un plan y recalcula km/fecha próximos
// ─────────────────────────────────────────────────────────────
function actualizarPlanMantenimiento(params) {
  try {
    if (!params.id) throw new Error('ID requerido');
    var sh   = _sheetPrev();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_PREV.ID] == params.id) {
        var fila = i + 1;

        var campos = {
          tipoServicio: MC_PREV.TIPO_SERVICIO,
          descripcion:  MC_PREV.DESCRIPCION,
          intervaloKm:  MC_PREV.INTERVALO_KM,
          intervaloDias:MC_PREV.INTERVALO_DIAS,
          kmUltimo:     MC_PREV.KM_ULTIMO,
          fechaUltimo:  MC_PREV.FECHA_ULTIMO,
          tecnico:      MC_PREV.TECNICO,
          estado:       MC_PREV.ESTADO,
          observaciones:MC_PREV.OBSERVACIONES
        };

        for (var campo in campos) {
          if (params[campo] !== undefined && params[campo] !== null) {
            sh.getRange(fila, campos[campo] + 1).setValue(params[campo]);
          }
        }

        // Recalcular próximos
        var intKm   = parseFloat(sh.getRange(fila, MC_PREV.INTERVALO_KM + 1).getValue()) || 0;
        var intDias = parseInt(sh.getRange(fila, MC_PREV.INTERVALO_DIAS + 1).getValue())  || 0;
        var kmUlt   = parseFloat(sh.getRange(fila, MC_PREV.KM_ULTIMO     + 1).getValue()) || 0;
        var fUlt    = sh.getRange(fila, MC_PREV.FECHA_ULTIMO + 1).getValue();

        if (intKm)   sh.getRange(fila, MC_PREV.KM_PROXIMO + 1).setValue(kmUlt + intKm);
        if (intDias && fUlt) {
          var base = new Date(fUlt);
          base.setDate(base.getDate() + intDias);
          sh.getRange(fila, MC_PREV.FECHA_PROXIMA + 1)
            .setValue(Utilities.formatDate(base, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
        }

        sh.getRange(fila, MC_PREV.TIMESTAMP + 1).setValue(new Date().toISOString());
        return { ok: true, data: { id: params.id, mensaje: 'Plan actualizado correctamente' } };
      }
    }
    throw new Error('Plan no encontrado: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularProximoMantenimiento(placa, kmActual)
//  Retorna todos los planes activos de una placa con
//  días/km restantes y alerta si está próximo o vencido
// ─────────────────────────────────────────────────────────────
function calcularProximoMantenimiento(placa, kmActual) {
  try {
    if (!placa) throw new Error('Placa requerida');

    var data     = _sheetPrev().getDataRange().getValues();
    var hoy      = new Date();
    var resultado = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[MC_PREV.ID]) continue;
      if (row[MC_PREV.PLACA]  !== String(placa).toUpperCase()) continue;
      if (row[MC_PREV.ESTADO] === 'INACTIVO') continue;

      var plan = _rowToPrev(row);

      // Distancia restante
      var kmRestante = null;
      var alertaKm   = false;
      if (plan.kmProximo && kmActual) {
        kmRestante = parseFloat(plan.kmProximo) - parseFloat(kmActual);
        alertaKm   = kmRestante <= 500; // alerta si queda 500 km o menos
      }

      // Días restantes
      var diasRestantes = null;
      var alertaDias    = false;
      if (plan.fechaProxima) {
        var fProx      = new Date(plan.fechaProxima);
        var diff       = fProx - hoy;
        diasRestantes  = Math.ceil(diff / 86400000);
        alertaDias     = diasRestantes <= 7; // alerta si queda 1 semana o menos
      }

      var estado = 'OK';
      if ((kmRestante !== null && kmRestante < 0) || (diasRestantes !== null && diasRestantes < 0)) {
        estado = 'VENCIDO';
      } else if (alertaKm || alertaDias) {
        estado = 'PROXIMO';
      }

      resultado.push({
        id:           plan.id,
        tipoServicio: plan.tipoServicio,
        descripcion:  plan.descripcion,
        kmProximo:    plan.kmProximo,
        fechaProxima: plan.fechaProxima,
        kmRestante:   kmRestante,
        diasRestantes:diasRestantes,
        alertaKm:     alertaKm,
        alertaDias:   alertaDias,
        estadoAlerta: estado
      });
    }

    // Ordenar: VENCIDO primero, luego PROXIMO, luego OK
    var ord = { 'VENCIDO': 0, 'PROXIMO': 1, 'OK': 2 };
    resultado.sort(function(a, b) {
      return (ord[a.estadoAlerta] || 9) - (ord[b.estadoAlerta] || 9);
    });

    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  listarPlanesPorVehiculo(placa)
// ─────────────────────────────────────────────────────────────
function listarPlanesPorVehiculo(placa) {
  try {
    if (!placa) throw new Error('Placa requerida');
    var data      = _sheetPrev().getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[MC_PREV.ID]) continue;
      if (row[MC_PREV.PLACA] !== String(placa).toUpperCase()) continue;
      resultado.push(_rowToPrev(row));
    }
    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
