// ============================================================
//  Baterias.gs  —  SIST-MTTO-SSISAC
//  Gestión del ciclo de vida de baterías
// ============================================================

var SHEET_BAT = 'BATERIAS';
var SHEET_BAT_HIST = 'BATERIAS_HISTORIAL';

var MC_BAT = {
  ID:            0,   // BAT-YYYY-NNNN
  SERIE:         1,
  MARCA:         2,
  CAPACIDAD_AH:  3,   // Amperios-hora
  VOLTAJE:       4,   // 12V | 24V
  TIPO:          5,   // PRINCIPAL | AUXILIAR | GRUPO_ELECTROGENO
  ESTADO:        6,   // NUEVO | EN_USO | BAJA | DESCARTADO
  ID_VEHICULO:   7,
  PLACA:         8,
  FECHA_FABRIC:  9,   // Fecha de fabricación
  FECHA_INST:    10,
  FECHA_VENC:    11,  // Vencimiento de garantía / vida útil estimada
  FECHA_RETIRO:  12,
  VOLTAJE_ACTUAL:13,  // Último voltaje medido
  DENSIDAD:      14,  // Última densidad medida
  CARGA_PCT:     15,  // Porcentaje de carga
  OBSERVACIONES: 16,
  USUARIO:       17,
  TIMESTAMP:     18
};

var CABECERA_BAT = [
  'ID_BATERIA','NUMERO_SERIE','MARCA','CAPACIDAD_AH','VOLTAJE_NOMINAL',
  'TIPO','ESTADO','ID_VEHICULO','PLACA',
  'FECHA_FABRICACION','FECHA_INSTALACION','FECHA_VENCIMIENTO','FECHA_RETIRO',
  'VOLTAJE_ACTUAL','DENSIDAD','CARGA_PORCENTAJE',
  'OBSERVACIONES','USUARIO','TIMESTAMP'
];

var MC_BAT_H = {
  ID:       0,
  FECHA:    1,
  ID_BAT:   2,
  SERIE:    3,
  PLACA:    4,
  ACCION:   5,   // INSTALACION | RETIRO | CARGA | PRUEBA | MANTENIMIENTO
  VOLTAJE:  6,
  DENSIDAD: 7,
  CARGA:    8,
  DETALLE:  9,
  USUARIO:  10,
  TIMESTAMP:11
};

var CABECERA_BAT_H = [
  'ID_HIST','FECHA','ID_BATERIA','NUMERO_SERIE','PLACA',
  'ACCION','VOLTAJE','DENSIDAD','CARGA_PCT','DETALLE','USUARIO','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetBat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_BAT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_BAT);
    sh.appendRow(CABECERA_BAT);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_BAT.length)
      .setBackground('#4A3A00').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _sheetBatHist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_BAT_HIST);
  if (!sh) {
    sh = ss.insertSheet(SHEET_BAT_HIST);
    sh.appendRow(CABECERA_BAT_H);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_BAT_H.length)
      .setBackground('#5A4A10').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdBat() {
  var sh   = _sheetBat();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'BAT-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('BAT-' + yr); });
  if (!ids.length) return 'BAT-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'BAT-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _nextIdBatH() {
  var sh   = _sheetBatHist();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'BATH-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('BATH-' + yr); });
  if (!ids.length) return 'BATH-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'BATH-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _rowToBat(row) {
  return {
    id:            row[MC_BAT.ID],
    serie:         row[MC_BAT.SERIE],
    marca:         row[MC_BAT.MARCA],
    capacidadAh:   row[MC_BAT.CAPACIDAD_AH],
    voltaje:       row[MC_BAT.VOLTAJE],
    tipo:          row[MC_BAT.TIPO],
    estado:        row[MC_BAT.ESTADO],
    idVehiculo:    row[MC_BAT.ID_VEHICULO],
    placa:         row[MC_BAT.PLACA],
    fechaFabric:   row[MC_BAT.FECHA_FABRIC],
    fechaInst:     row[MC_BAT.FECHA_INST],
    fechaVenc:     row[MC_BAT.FECHA_VENC],
    fechaRetiro:   row[MC_BAT.FECHA_RETIRO],
    voltajeActual: row[MC_BAT.VOLTAJE_ACTUAL],
    densidad:      row[MC_BAT.DENSIDAD],
    cargaPct:      row[MC_BAT.CARGA_PCT],
    observaciones: row[MC_BAT.OBSERVACIONES],
    usuario:       row[MC_BAT.USUARIO],
    timestamp:     row[MC_BAT.TIMESTAMP]
  };
}

// ─────────────────────────────────────────────────────────────
//  registrarBateria(params)
// ─────────────────────────────────────────────────────────────
function registrarBateria(params) {
  try {
    if (!params.serie) throw new Error('Número de serie requerido');
    if (!params.marca) throw new Error('Marca requerida');

    var sh  = _sheetBat();
    var id  = _nextIdBat();
    var now = new Date();
    var hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Calcular fecha de vencimiento (2 años por defecto)
    var fechaVenc = '';
    if (params.fechaFabric) {
      var base = new Date(params.fechaFabric);
      base.setFullYear(base.getFullYear() + (parseInt(params.vidaUtilAnios) || 2));
      fechaVenc = Utilities.formatDate(base, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    sh.appendRow([
      id,
      String(params.serie).toUpperCase(),
      params.marca,
      parseFloat(params.capacidadAh) || 150,
      parseFloat(params.voltaje)     || 12,
      String(params.tipo || 'PRINCIPAL').toUpperCase(),
      'NUEVO',
      '', '',
      params.fechaFabric || '',
      '', fechaVenc, '',
      parseFloat(params.voltajeActual) || 0,
      parseFloat(params.densidad)      || 0,
      parseFloat(params.cargaPct)      || 100,
      params.observaciones || '',
      params.usuario || 'SISTEMA',
      now.toISOString()
    ]);

    return { ok: true, data: { id: id, serie: params.serie, fechaVenc: fechaVenc,
                                mensaje: 'Batería registrada: ' + id } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  instalarBateria(params)
//  params: id, idVehiculo, placa, voltajeActual, cargaPct, usuario
// ─────────────────────────────────────────────────────────────
function instalarBateria(params) {
  try {
    if (!params.id)    throw new Error('ID de batería requerido');
    if (!params.placa) throw new Error('Placa requerida');

    var sh   = _sheetBat();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_BAT.ID] == params.id) {
        var fila = i + 1;
        var hoy  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

        sh.getRange(fila, MC_BAT.ESTADO       + 1).setValue('EN_USO');
        sh.getRange(fila, MC_BAT.ID_VEHICULO  + 1).setValue(params.idVehiculo || '');
        sh.getRange(fila, MC_BAT.PLACA        + 1).setValue(String(params.placa).toUpperCase());
        sh.getRange(fila, MC_BAT.FECHA_INST   + 1).setValue(params.fechaInst || hoy);
        sh.getRange(fila, MC_BAT.FECHA_RETIRO + 1).setValue('');
        if (params.voltajeActual) sh.getRange(fila, MC_BAT.VOLTAJE_ACTUAL + 1).setValue(params.voltajeActual);
        if (params.cargaPct)      sh.getRange(fila, MC_BAT.CARGA_PCT      + 1).setValue(params.cargaPct);
        sh.getRange(fila, MC_BAT.TIMESTAMP    + 1).setValue(new Date().toISOString());

        _agregarHistBat({
          idBat:   params.id,
          serie:   data[i][MC_BAT.SERIE],
          placa:   params.placa,
          accion:  'INSTALACION',
          voltaje: params.voltajeActual,
          carga:   params.cargaPct,
          detalle: 'Instalada en vehículo ' + params.placa,
          usuario: params.usuario || 'SISTEMA'
        });

        return { ok: true, data: { id: params.id, placa: params.placa, mensaje: 'Batería instalada' } };
      }
    }
    throw new Error('Batería no encontrada: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  retirarBateria(params)
//  params: id, motivo, nuevoEstado, usuario
// ─────────────────────────────────────────────────────────────
function retirarBateria(params) {
  try {
    if (!params.id) throw new Error('ID de batería requerido');
    var sh   = _sheetBat();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_BAT.ID] == params.id) {
        var fila = i + 1;
        var hoy  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

        sh.getRange(fila, MC_BAT.ESTADO       + 1).setValue(params.nuevoEstado || 'DESCARTADO');
        sh.getRange(fila, MC_BAT.FECHA_RETIRO + 1).setValue(params.fechaRetiro || hoy);
        sh.getRange(fila, MC_BAT.ID_VEHICULO  + 1).setValue('');
        sh.getRange(fila, MC_BAT.PLACA        + 1).setValue('');
        sh.getRange(fila, MC_BAT.TIMESTAMP    + 1).setValue(new Date().toISOString());

        _agregarHistBat({
          idBat:   params.id,
          serie:   data[i][MC_BAT.SERIE],
          placa:   data[i][MC_BAT.PLACA],
          accion:  'RETIRO',
          detalle: params.motivo || 'Retiro de vehículo',
          usuario: params.usuario || 'SISTEMA'
        });

        return { ok: true, data: { id: params.id, mensaje: 'Batería retirada' } };
      }
    }
    throw new Error('Batería no encontrada: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  obtenerHistorialBateria(id)
// ─────────────────────────────────────────────────────────────
function obtenerHistorialBateria(id) {
  try {
    if (!id) throw new Error('ID requerido');
    var data      = _sheetBatHist().getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_BAT_H.ID_BAT] == id) {
        resultado.push({
          id:       data[i][MC_BAT_H.ID],
          fecha:    data[i][MC_BAT_H.FECHA],
          placa:    data[i][MC_BAT_H.PLACA],
          accion:   data[i][MC_BAT_H.ACCION],
          voltaje:  data[i][MC_BAT_H.VOLTAJE],
          densidad: data[i][MC_BAT_H.DENSIDAD],
          carga:    data[i][MC_BAT_H.CARGA],
          detalle:  data[i][MC_BAT_H.DETALLE]
        });
      }
    }
    resultado.sort(function(a, b) { return b.fecha > a.fecha ? 1 : -1; });
    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function _agregarHistBat(p) {
  var sh  = _sheetBatHist();
  var id  = _nextIdBatH();
  var now = new Date();
  var hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow([
    id, hoy, p.idBat, p.serie, String(p.placa || '').toUpperCase(),
    p.accion || '', p.voltaje || '', p.densidad || '', p.carga || '',
    p.detalle || '', p.usuario || 'SISTEMA', now.toISOString()
  ]);
}
