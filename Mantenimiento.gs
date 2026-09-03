// ============================================================
//  Mantenimiento.gs  —  SIST-MTTO-SSISAC
//  Funciones generales del módulo de mantenimiento
// ============================================================

// ── Nombre de la hoja ──
var SHEET_MANT = 'MANTENIMIENTOS';

// ── Índices de columnas (0-based) ──
var MC_MANT = {
  ID:           0,   // ID_MANTENIMIENTO  (MANT-YYYY-NNNN)
  FECHA_REG:    1,   // FECHA_REGISTRO
  TIPO:         2,   // CORRECTIVO | PREVENTIVO | PREDICTIVO
  ID_VEHICULO:  3,   // ID interno del vehículo
  PLACA:        4,
  KM:           5,   // KM_ACTUAL al registrar
  ID_OT:        6,   // OT asociada (puede estar vacía al inicio)
  ESTADO:       7,   // PENDIENTE | EN_PROCESO | COMPLETADO | CANCELADO
  DESCRIPCION:  8,
  FECHA_INICIO: 9,
  FECHA_FIN:    10,
  DURACION_H:   11,  // Duración en horas (número)
  TECNICO:      12,
  TALLER:       13,
  COSTO_REP:    14,  // Costo repuestos
  COSTO_MO:     15,  // Costo mano de obra
  COSTO_TOTAL:  16,
  OBSERVACIONES:17,
  USUARIO:      18,
  TIMESTAMP:    19
};

var CABECERA_MANT = [
  'ID_MANTENIMIENTO','FECHA_REGISTRO','TIPO','ID_VEHICULO','PLACA',
  'KM_ACTUAL','ID_OT','ESTADO','DESCRIPCION','FECHA_INICIO','FECHA_FIN',
  'DURACION_HORAS','TECNICO','TALLER','COSTO_REPUESTOS','COSTO_MO',
  'COSTO_TOTAL','OBSERVACIONES','USUARIO','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetMant() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MANT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MANT);
    sh.appendRow(CABECERA_MANT);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_MANT.length)
      .setBackground('#5B1D00').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdMant() {
  var sh   = _sheetMant();
  var last = sh.getLastRow();
  var yr   = new Date().getFullYear();
  if (last < 2) return 'MANT-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('MANT-' + yr); });
  if (!ids.length) return 'MANT-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  var next = Math.max.apply(null, nums) + 1;
  return 'MANT-' + yr + '-' + String(next).padStart(4, '0');
}

function _rowToMant(row) {
  return {
    id:           row[MC_MANT.ID],
    fechaRegistro:row[MC_MANT.FECHA_REG],
    tipo:         row[MC_MANT.TIPO],
    idVehiculo:   row[MC_MANT.ID_VEHICULO],
    placa:        row[MC_MANT.PLACA],
    km:           row[MC_MANT.KM],
    idOT:         row[MC_MANT.ID_OT],
    estado:       row[MC_MANT.ESTADO],
    descripcion:  row[MC_MANT.DESCRIPCION],
    fechaInicio:  row[MC_MANT.FECHA_INICIO],
    fechaFin:     row[MC_MANT.FECHA_FIN],
    duracionH:    row[MC_MANT.DURACION_H],
    tecnico:      row[MC_MANT.TECNICO],
    taller:       row[MC_MANT.TALLER],
    costoRep:     row[MC_MANT.COSTO_REP],
    costoMO:      row[MC_MANT.COSTO_MO],
    costoTotal:   row[MC_MANT.COSTO_TOTAL],
    observaciones:row[MC_MANT.OBSERVACIONES],
    usuario:      row[MC_MANT.USUARIO],
    timestamp:    row[MC_MANT.TIMESTAMP]
  };
}

// ─────────────────────────────────────────────────────────────
//  registrarMantenimiento(params)
//  params: tipo, idVehiculo, placa, km, descripcion,
//          fechaInicio, tecnico, taller, observaciones, usuario
// ─────────────────────────────────────────────────────────────
function registrarMantenimiento(params) {
  try {
    if (!params.placa)       throw new Error('Placa requerida');
    if (!params.tipo)        throw new Error('Tipo de mantenimiento requerido');
    if (!params.descripcion) throw new Error('Descripción requerida');

    var sh   = _sheetMant();
    var id   = _nextIdMant();
    var now  = new Date();
    var fReg = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    sh.appendRow([
      id,
      fReg,
      String(params.tipo).toUpperCase(),
      params.idVehiculo || '',
      String(params.placa).toUpperCase(),
      params.km || '',
      params.idOT || '',
      params.estado || 'PENDIENTE',
      params.descripcion,
      params.fechaInicio || fReg,
      params.fechaFin || '',
      params.duracionH || '',
      params.tecnico || '',
      params.taller || '',
      0, 0, 0,
      params.observaciones || '',
      params.usuario || 'SISTEMA',
      now.toISOString()
    ]);

    // Generar OT automáticamente si no se provee una
    if (!params.idOT) {
      var ot = crearOrdenTrabajo({
        idVehiculo:  params.idVehiculo,
        placa:       params.placa,
        tipo:        params.tipo,
        descripcion: params.descripcion,
        tecnico:     params.tecnico,
        usuario:     params.usuario,
        idMant:      id
      });
      if (ot.ok) {
        actualizarMantenimiento({ id: id, idOT: ot.data.id });
      }
    }

    return { ok: true, data: { id: id, mensaje: 'Mantenimiento registrado correctamente' } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  obtenerMantenimiento(id)
// ─────────────────────────────────────────────────────────────
function obtenerMantenimiento(id) {
  try {
    if (!id) throw new Error('ID requerido');
    var sh   = _sheetMant();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_MANT.ID] == id) {
        return { ok: true, data: _rowToMant(data[i]) };
      }
    }
    throw new Error('Mantenimiento no encontrado: ' + id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  actualizarMantenimiento(params)
//  params: id  + cualquier campo a actualizar
// ─────────────────────────────────────────────────────────────
function actualizarMantenimiento(params) {
  try {
    if (!params.id) throw new Error('ID requerido');
    var sh   = _sheetMant();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_MANT.ID] == params.id) {
        var fila = i + 1; // 1-based row

        var campos = {
          tipo:         MC_MANT.TIPO,
          placa:        MC_MANT.PLACA,
          km:           MC_MANT.KM,
          idOT:         MC_MANT.ID_OT,
          estado:       MC_MANT.ESTADO,
          descripcion:  MC_MANT.DESCRIPCION,
          fechaInicio:  MC_MANT.FECHA_INICIO,
          fechaFin:     MC_MANT.FECHA_FIN,
          duracionH:    MC_MANT.DURACION_H,
          tecnico:      MC_MANT.TECNICO,
          taller:       MC_MANT.TALLER,
          costoRep:     MC_MANT.COSTO_REP,
          costoMO:      MC_MANT.COSTO_MO,
          costoTotal:   MC_MANT.COSTO_TOTAL,
          observaciones:MC_MANT.OBSERVACIONES
        };

        for (var campo in campos) {
          if (params[campo] !== undefined && params[campo] !== null) {
            sh.getRange(fila, campos[campo] + 1).setValue(params[campo]);
          }
        }

        // Recalcular costo total si se actualizan costos parciales
        if (params.costoRep !== undefined || params.costoMO !== undefined) {
          var cRep = parseFloat(sh.getRange(fila, MC_MANT.COSTO_REP + 1).getValue()) || 0;
          var cMO  = parseFloat(sh.getRange(fila, MC_MANT.COSTO_MO  + 1).getValue()) || 0;
          sh.getRange(fila, MC_MANT.COSTO_TOTAL + 1).setValue(cRep + cMO);
        }

        // Calcular duración automáticamente si se dan ambas fechas
        if (params.fechaInicio && params.fechaFin) {
          var d1 = new Date(params.fechaInicio);
          var d2 = new Date(params.fechaFin);
          var horas = Math.round((d2 - d1) / 3600000 * 10) / 10;
          if (horas >= 0) sh.getRange(fila, MC_MANT.DURACION_H + 1).setValue(horas);
        }

        sh.getRange(fila, MC_MANT.TIMESTAMP + 1).setValue(new Date().toISOString());
        return { ok: true, data: { id: params.id, mensaje: 'Actualizado correctamente' } };
      }
    }
    throw new Error('Mantenimiento no encontrado: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  listarMantenimientos(filtros)
//  filtros: placa, tipo, estado, fechaDesde, fechaHasta
// ─────────────────────────────────────────────────────────────
function listarMantenimientos(filtros) {
  try {
    filtros = filtros || {};
    var sh   = _sheetMant();
    var data = sh.getDataRange().getValues();
    var resultado = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[MC_MANT.ID]) continue;

      if (filtros.placa  && row[MC_MANT.PLACA]  !== filtros.placa.toUpperCase()) continue;
      if (filtros.tipo   && row[MC_MANT.TIPO]   !== filtros.tipo.toUpperCase())  continue;
      if (filtros.estado && row[MC_MANT.ESTADO] !== filtros.estado.toUpperCase()) continue;
      if (filtros.fechaDesde && row[MC_MANT.FECHA_REG] < filtros.fechaDesde) continue;
      if (filtros.fechaHasta && row[MC_MANT.FECHA_REG] > filtros.fechaHasta) continue;

      resultado.push(_rowToMant(row));
    }

    resultado.sort(function(a, b) {
      return b.fechaRegistro > a.fechaRegistro ? 1 : -1;
    });

    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  obtenerHistorialVehiculo(placa, limite)
//  Devuelve historial completo de un vehículo ordenado por fecha desc
// ─────────────────────────────────────────────────────────────
function obtenerHistorialVehiculo(placa, limite) {
  try {
    if (!placa) throw new Error('Placa requerida');
    limite = limite || 50;

    var res = listarMantenimientos({ placa: placa });
    if (!res.ok) throw new Error(res.error);

    var historial = res.data.slice(0, limite);

    // Estadísticas rápidas
    var stats = {
      total:       historial.length,
      correctivos: historial.filter(function(m){ return m.tipo === 'CORRECTIVO';  }).length,
      preventivos: historial.filter(function(m){ return m.tipo === 'PREVENTIVO';  }).length,
      predictivos: historial.filter(function(m){ return m.tipo === 'PREDICTIVO';  }).length,
      costoTotal:  historial.reduce(function(s, m){ return s + (parseFloat(m.costoTotal) || 0); }, 0)
    };

    return { ok: true, data: historial, stats: stats };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
