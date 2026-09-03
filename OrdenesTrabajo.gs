// ============================================================
//  OrdenesTrabajo.gs  —  SIST-MTTO-SSISAC
//  Gestión de Órdenes de Trabajo (OT)
// ============================================================

var SHEET_OT = 'ORDENES_TRABAJO';

var MC_OT = {
  ID:           0,   // ID_OT  (OT-YYYY-NNNN)
  FECHA_REG:    1,
  ID_VEHICULO:  2,
  PLACA:        3,
  TIPO:         4,   // CORRECTIVO | PREVENTIVO | PREDICTIVO
  DESCRIPCION:  5,
  ESTADO:       6,   // BORRADOR | ABIERTA | EN_PROCESO | CERRADA | CANCELADA
  PRIORIDAD:    7,   // ALTA | MEDIA | BAJA
  ID_MANT:      8,   // Mantenimiento asociado
  TECNICO:      9,
  FECHA_PROG_I: 10,  // Fecha inicio programada
  FECHA_PROG_F: 11,  // Fecha fin programada
  FECHA_REAL_I: 12,  // Fecha inicio real
  FECHA_CIERRE: 13,
  COSTO_REP:    14,
  COSTO_MO:     15,
  COSTO_TOTAL:  16,
  OBSERVACIONES:17,
  USUARIO:      18,
  TIMESTAMP:    19
};

var CABECERA_OT = [
  'ID_OT','FECHA_REGISTRO','ID_VEHICULO','PLACA','TIPO','DESCRIPCION',
  'ESTADO','PRIORIDAD','ID_MANTENIMIENTO','TECNICO_ASIGNADO',
  'FECHA_PROG_INICIO','FECHA_PROG_FIN','FECHA_REAL_INICIO','FECHA_CIERRE',
  'COSTO_REPUESTOS','COSTO_MO','COSTO_TOTAL','OBSERVACIONES','USUARIO','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetOT() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_OT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_OT);
    sh.appendRow(CABECERA_OT);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_OT.length)
      .setBackground('#8C3400').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdOT() {
  var sh  = _sheetOT();
  var yr  = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'OT-' + yr + '-0001';
  var ids = sh.getRange(2, 1, last - 1, 1).getValues()
              .map(function(r){ return r[0]; })
              .filter(function(v){ return String(v).startsWith('OT-' + yr); });
  if (!ids.length) return 'OT-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'OT-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _rowToOT(row) {
  return {
    id:          row[MC_OT.ID],
    fechaReg:    row[MC_OT.FECHA_REG],
    idVehiculo:  row[MC_OT.ID_VEHICULO],
    placa:       row[MC_OT.PLACA],
    tipo:        row[MC_OT.TIPO],
    descripcion: row[MC_OT.DESCRIPCION],
    estado:      row[MC_OT.ESTADO],
    prioridad:   row[MC_OT.PRIORIDAD],
    idMant:      row[MC_OT.ID_MANT],
    tecnico:     row[MC_OT.TECNICO],
    fechaProgI:  row[MC_OT.FECHA_PROG_I],
    fechaProgF:  row[MC_OT.FECHA_PROG_F],
    fechaRealI:  row[MC_OT.FECHA_REAL_I],
    fechaCierre: row[MC_OT.FECHA_CIERRE],
    costoRep:    row[MC_OT.COSTO_REP],
    costoMO:     row[MC_OT.COSTO_MO],
    costoTotal:  row[MC_OT.COSTO_TOTAL],
    observaciones:row[MC_OT.OBSERVACIONES],
    usuario:     row[MC_OT.USUARIO],
    timestamp:   row[MC_OT.TIMESTAMP]
  };
}

// ─────────────────────────────────────────────────────────────
//  crearOrdenTrabajo(params)
//  params: placa, tipo, descripcion, prioridad, tecnico,
//          fechaProgI, fechaProgF, idMant, usuario
// ─────────────────────────────────────────────────────────────
function crearOrdenTrabajo(params) {
  try {
    if (!params.placa)       throw new Error('Placa requerida');
    if (!params.descripcion) throw new Error('Descripción requerida');

    var sh  = _sheetOT();
    var id  = _nextIdOT();
    var now = new Date();
    var hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    sh.appendRow([
      id,
      hoy,
      params.idVehiculo || '',
      String(params.placa).toUpperCase(),
      String(params.tipo || 'CORRECTIVO').toUpperCase(),
      params.descripcion,
      'ABIERTA',
      String(params.prioridad || 'MEDIA').toUpperCase(),
      params.idMant || '',
      params.tecnico || '',
      params.fechaProgI || hoy,
      params.fechaProgF || '',
      '', '',  // fechas reales vacías
      0, 0, 0,
      params.observaciones || '',
      params.usuario || 'SISTEMA',
      now.toISOString()
    ]);

    return { ok: true, data: { id: id, mensaje: 'Orden de trabajo creada: ' + id } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  obtenerOrdenTrabajo(id)
// ─────────────────────────────────────────────────────────────
function obtenerOrdenTrabajo(id) {
  try {
    if (!id) throw new Error('ID requerido');
    var data = _sheetOT().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_OT.ID] == id) {
        var ot = _rowToOT(data[i]);
        // Enriquecer: traer intervenciones y repuestos
        ot.intervenciones = obtenerIntervencionesPorOT(id).data || [];
        ot.repuestos      = obtenerRepuestosPorOT(id).data    || [];
        ot.manoObra       = obtenerManoObraPorOT(id).data     || [];
        return { ok: true, data: ot };
      }
    }
    throw new Error('OT no encontrada: ' + id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  actualizarEstadoOT(id, estado, usuario, observaciones)
//  estados válidos: BORRADOR | ABIERTA | EN_PROCESO | CERRADA | CANCELADA
// ─────────────────────────────────────────────────────────────
function actualizarEstadoOT(id, estado, usuario, observaciones) {
  try {
    if (!id)     throw new Error('ID requerido');
    if (!estado) throw new Error('Estado requerido');

    var ESTADOS_VALIDOS = ['BORRADOR','ABIERTA','EN_PROCESO','CERRADA','CANCELADA'];
    estado = estado.toUpperCase();
    if (ESTADOS_VALIDOS.indexOf(estado) === -1) {
      throw new Error('Estado inválido. Use: ' + ESTADOS_VALIDOS.join(', '));
    }

    var sh   = _sheetOT();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_OT.ID] == id) {
        var fila = i + 1;
        sh.getRange(fila, MC_OT.ESTADO + 1).setValue(estado);
        if (observaciones) sh.getRange(fila, MC_OT.OBSERVACIONES + 1).setValue(observaciones);

        // Registrar fecha de inicio real si pasa a EN_PROCESO
        if (estado === 'EN_PROCESO' && !data[i][MC_OT.FECHA_REAL_I]) {
          var hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
          sh.getRange(fila, MC_OT.FECHA_REAL_I + 1).setValue(hoy);
        }

        sh.getRange(fila, MC_OT.TIMESTAMP + 1).setValue(new Date().toISOString());
        return { ok: true, data: { id: id, estado: estado } };
      }
    }
    throw new Error('OT no encontrada: ' + id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  cerrarOrdenTrabajo(params)
//  Cierra la OT, actualiza costos finales y vincula al mantenimiento
// ─────────────────────────────────────────────────────────────
function cerrarOrdenTrabajo(params) {
  try {
    if (!params.id) throw new Error('ID requerido');

    var sh   = _sheetOT();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_OT.ID] == params.id) {
        var fila    = i + 1;
        var hoy     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var costoR  = parseFloat(params.costoRep)   || 0;
        var costoMO = parseFloat(params.costoMO)    || 0;
        var total   = costoR + costoMO;

        sh.getRange(fila, MC_OT.ESTADO       + 1).setValue('CERRADA');
        sh.getRange(fila, MC_OT.FECHA_CIERRE + 1).setValue(params.fechaCierre || hoy);
        sh.getRange(fila, MC_OT.COSTO_REP    + 1).setValue(costoR);
        sh.getRange(fila, MC_OT.COSTO_MO     + 1).setValue(costoMO);
        sh.getRange(fila, MC_OT.COSTO_TOTAL  + 1).setValue(total);
        if (params.observaciones) sh.getRange(fila, MC_OT.OBSERVACIONES + 1).setValue(params.observaciones);
        sh.getRange(fila, MC_OT.TIMESTAMP    + 1).setValue(new Date().toISOString());

        // Actualizar mantenimiento vinculado
        var idMant = data[i][MC_OT.ID_MANT];
        if (idMant) {
          actualizarMantenimiento({
            id:         idMant,
            estado:     'COMPLETADO',
            fechaFin:   params.fechaCierre || hoy,
            costoRep:   costoR,
            costoMO:    costoMO,
            costoTotal: total
          });
        }

        return { ok: true, data: { id: params.id, costoTotal: total, mensaje: 'OT cerrada correctamente' } };
      }
    }
    throw new Error('OT no encontrada: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  listarOrdenesPendientes(filtros)
//  filtros: placa, tecnico, prioridad
// ─────────────────────────────────────────────────────────────
function listarOrdenesPendientes(filtros) {
  try {
    filtros = filtros || {};
    var sh   = _sheetOT();
    var data = sh.getDataRange().getValues();
    var resultado = [];
    var PENDIENTES = ['BORRADOR','ABIERTA','EN_PROCESO'];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[MC_OT.ID]) continue;
      if (PENDIENTES.indexOf(String(row[MC_OT.ESTADO])) === -1) continue;

      if (filtros.placa    && row[MC_OT.PLACA]    !== filtros.placa.toUpperCase()) continue;
      if (filtros.tecnico  && row[MC_OT.TECNICO]  !== filtros.tecnico.toUpperCase()) continue;
      if (filtros.prioridad && row[MC_OT.PRIORIDAD] !== filtros.prioridad.toUpperCase()) continue;

      resultado.push(_rowToOT(row));
    }

    // Ordenar: ALTA primero, luego MEDIA, luego BAJA, y por fecha
    var orden = { 'ALTA': 0, 'MEDIA': 1, 'BAJA': 2 };
    resultado.sort(function(a, b) {
      var pa = orden[a.prioridad] !== undefined ? orden[a.prioridad] : 9;
      var pb = orden[b.prioridad] !== undefined ? orden[b.prioridad] : 9;
      if (pa !== pb) return pa - pb;
      return a.fechaProgI > b.fechaProgI ? 1 : -1;
    });

    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
