// ============================================================
//  Neumaticos.gs  —  SIST-MTTO-SSISAC
//  Gestión del ciclo de vida de neumáticos
// ============================================================

var SHEET_NEUM = 'NEUMATICOS';
var SHEET_NEUM_HIST = 'NEUMATICOS_HISTORIAL';

var MC_NEUM = {
  ID:           0,   // NEUM-YYYY-NNNN
  SERIE:        1,   // Nº de serie del neumático
  MARCA:        2,
  MEDIDA:       3,   // Ej: 295/80R22.5
  TIPO:         4,   // TRACCION | DIRECCION | REMOLQUE
  ESTADO:       5,   // NUEVO | EN_USO | REENCAUCHADO | DESCARTADO
  ID_VEHICULO:  6,
  PLACA:        7,
  POSICION:     8,   // P1..P12 (posición en el vehículo)
  KM_INICIAL:   9,   // KM al instalar
  KM_ACTUAL:    10,
  KM_RECORRIDO: 11,
  KM_LIMITE:    12,  // KM máximo de vida útil
  FECHA_INST:   13,
  FECHA_RETIRO: 14,
  PROFUNDIDAD_MM:15, // Profundidad del labrado
  PRESION_PSI:  16,
  OBSERVACIONES:17,
  USUARIO:      18,
  TIMESTAMP:    19
};

var CABECERA_NEUM = [
  'ID_NEUMATICO','NUMERO_SERIE','MARCA','MEDIDA','TIPO','ESTADO',
  'ID_VEHICULO','PLACA','POSICION','KM_INSTALACION','KM_ACTUAL',
  'KM_RECORRIDO','KM_LIMITE_VIDA','FECHA_INSTALACION','FECHA_RETIRO',
  'PROFUNDIDAD_MM','PRESION_PSI','OBSERVACIONES','USUARIO','TIMESTAMP'
];

var MC_NEUM_H = {
  ID:        0,
  FECHA:     1,
  ID_NEUM:   2,
  SERIE:     3,
  PLACA:     4,
  POSICION:  5,
  ACCION:    6,   // INSTALACION | RETIRO | ROTACION | INSPECCION | REENCAUCHE
  KM:        7,
  PROFUNDIDAD:8,
  PRESION:   9,
  DETALLE:   10,
  USUARIO:   11,
  TIMESTAMP: 12
};

var CABECERA_NEUM_H = [
  'ID_HIST','FECHA','ID_NEUMATICO','NUMERO_SERIE','PLACA','POSICION',
  'ACCION','KM_ACTUAL','PROFUNDIDAD_MM','PRESION_PSI','DETALLE','USUARIO','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetNeum() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NEUM);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NEUM);
    sh.appendRow(CABECERA_NEUM);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_NEUM.length)
      .setBackground('#1A1A4A').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _sheetNeumHist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NEUM_HIST);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NEUM_HIST);
    sh.appendRow(CABECERA_NEUM_H);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_NEUM_H.length)
      .setBackground('#2A2A5A').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdNeum() {
  var sh   = _sheetNeum();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'NEUM-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('NEUM-' + yr); });
  if (!ids.length) return 'NEUM-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'NEUM-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _nextIdNeumH() {
  var sh   = _sheetNeumHist();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'NEUH-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('NEUH-' + yr); });
  if (!ids.length) return 'NEUH-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'NEUH-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _rowToNeum(row) {
  return {
    id:           row[MC_NEUM.ID],
    serie:        row[MC_NEUM.SERIE],
    marca:        row[MC_NEUM.MARCA],
    medida:       row[MC_NEUM.MEDIDA],
    tipo:         row[MC_NEUM.TIPO],
    estado:       row[MC_NEUM.ESTADO],
    idVehiculo:   row[MC_NEUM.ID_VEHICULO],
    placa:        row[MC_NEUM.PLACA],
    posicion:     row[MC_NEUM.POSICION],
    kmInicial:    row[MC_NEUM.KM_INICIAL],
    kmActual:     row[MC_NEUM.KM_ACTUAL],
    kmRecorrido:  row[MC_NEUM.KM_RECORRIDO],
    kmLimite:     row[MC_NEUM.KM_LIMITE],
    fechaInst:    row[MC_NEUM.FECHA_INST],
    fechaRetiro:  row[MC_NEUM.FECHA_RETIRO],
    profundidad:  row[MC_NEUM.PROFUNDIDAD_MM],
    presion:      row[MC_NEUM.PRESION_PSI],
    observaciones:row[MC_NEUM.OBSERVACIONES],
    usuario:      row[MC_NEUM.USUARIO],
    timestamp:    row[MC_NEUM.TIMESTAMP]
  };
}

// ─────────────────────────────────────────────────────────────
//  registrarNeumatico(params)
//  Registra un neumático nuevo en el sistema (sin instalar)
// ─────────────────────────────────────────────────────────────
function registrarNeumatico(params) {
  try {
    if (!params.serie) throw new Error('Número de serie requerido');
    if (!params.marca) throw new Error('Marca requerida');
    if (!params.medida) throw new Error('Medida requerida');

    var sh  = _sheetNeum();
    var id  = _nextIdNeum();
    var now = new Date();
    var hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    sh.appendRow([
      id,
      String(params.serie).toUpperCase(),
      params.marca,
      params.medida,
      String(params.tipo || 'TRACCION').toUpperCase(),
      'NUEVO',
      '', '', '',
      0, 0, 0,
      parseFloat(params.kmLimite) || 120000,
      '', '',
      parseFloat(params.profundidad) || 18,
      parseFloat(params.presion)     || 110,
      params.observaciones || '',
      params.usuario || 'SISTEMA',
      now.toISOString()
    ]);

    return { ok: true, data: { id: id, serie: params.serie, mensaje: 'Neumático registrado: ' + id } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  instalarNeumatico(params)
//  params: id (del neumático), idVehiculo, placa, posicion,
//          kmActual, profundidad, presion, usuario
// ─────────────────────────────────────────────────────────────
function instalarNeumatico(params) {
  try {
    if (!params.id)      throw new Error('ID de neumático requerido');
    if (!params.placa)   throw new Error('Placa requerida');
    if (!params.posicion) throw new Error('Posición requerida');

    var sh   = _sheetNeum();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_NEUM.ID] == params.id) {
        var fila = i + 1;
        var hoy  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var km   = parseFloat(params.kmActual) || 0;

        sh.getRange(fila, MC_NEUM.ESTADO      + 1).setValue('EN_USO');
        sh.getRange(fila, MC_NEUM.ID_VEHICULO + 1).setValue(params.idVehiculo || '');
        sh.getRange(fila, MC_NEUM.PLACA       + 1).setValue(String(params.placa).toUpperCase());
        sh.getRange(fila, MC_NEUM.POSICION    + 1).setValue(params.posicion);
        sh.getRange(fila, MC_NEUM.KM_INICIAL  + 1).setValue(km);
        sh.getRange(fila, MC_NEUM.KM_ACTUAL   + 1).setValue(km);
        sh.getRange(fila, MC_NEUM.KM_RECORRIDO + 1).setValue(0);
        sh.getRange(fila, MC_NEUM.FECHA_INST  + 1).setValue(params.fechaInst || hoy);
        sh.getRange(fila, MC_NEUM.FECHA_RETIRO + 1).setValue('');
        if (params.profundidad) sh.getRange(fila, MC_NEUM.PROFUNDIDAD_MM + 1).setValue(params.profundidad);
        if (params.presion)     sh.getRange(fila, MC_NEUM.PRESION_PSI    + 1).setValue(params.presion);
        sh.getRange(fila, MC_NEUM.TIMESTAMP   + 1).setValue(new Date().toISOString());

        // Historial
        _agregarHistNeum({
          idNeum:   params.id,
          serie:    data[i][MC_NEUM.SERIE],
          placa:    params.placa,
          posicion: params.posicion,
          accion:   'INSTALACION',
          km:       km,
          profundidad: params.profundidad,
          presion:     params.presion,
          detalle:  'Instalado en posición ' + params.posicion,
          usuario:  params.usuario || 'SISTEMA'
        });

        return { ok: true, data: { id: params.id, placa: params.placa,
                                   posicion: params.posicion, mensaje: 'Neumático instalado' } };
      }
    }
    throw new Error('Neumático no encontrado: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  retirarNeumatico(params)
//  params: id, kmRetiro, motivo, usuario
// ─────────────────────────────────────────────────────────────
function retirarNeumatico(params) {
  try {
    if (!params.id) throw new Error('ID de neumático requerido');
    var sh   = _sheetNeum();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_NEUM.ID] == params.id) {
        var fila      = i + 1;
        var hoy       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var kmRetiro  = parseFloat(params.kmRetiro) || parseFloat(data[i][MC_NEUM.KM_ACTUAL]) || 0;
        var kmIni     = parseFloat(data[i][MC_NEUM.KM_INICIAL]) || 0;
        var kmRec     = Math.max(0, kmRetiro - kmIni);

        sh.getRange(fila, MC_NEUM.ESTADO       + 1).setValue(params.nuevoEstado || 'DESCARTADO');
        sh.getRange(fila, MC_NEUM.KM_ACTUAL    + 1).setValue(kmRetiro);
        sh.getRange(fila, MC_NEUM.KM_RECORRIDO + 1).setValue(kmRec);
        sh.getRange(fila, MC_NEUM.FECHA_RETIRO + 1).setValue(params.fechaRetiro || hoy);
        sh.getRange(fila, MC_NEUM.ID_VEHICULO  + 1).setValue('');
        sh.getRange(fila, MC_NEUM.PLACA        + 1).setValue('');
        sh.getRange(fila, MC_NEUM.POSICION     + 1).setValue('');
        sh.getRange(fila, MC_NEUM.TIMESTAMP    + 1).setValue(new Date().toISOString());

        _agregarHistNeum({
          idNeum:   params.id,
          serie:    data[i][MC_NEUM.SERIE],
          placa:    data[i][MC_NEUM.PLACA],
          posicion: data[i][MC_NEUM.POSICION],
          accion:   'RETIRO',
          km:       kmRetiro,
          detalle:  params.motivo || 'Retiro de vehículo',
          usuario:  params.usuario || 'SISTEMA'
        });

        return { ok: true, data: { id: params.id, kmRecorrido: kmRec, mensaje: 'Neumático retirado' } };
      }
    }
    throw new Error('Neumático no encontrado: ' + params.id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularKmRecorridos(id)
// ─────────────────────────────────────────────────────────────
function calcularKmRecorridos(id) {
  try {
    var data = _sheetNeum().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_NEUM.ID] == id) {
        var neum       = _rowToNeum(data[i]);
        var kmIni      = parseFloat(neum.kmInicial) || 0;
        var kmAct      = parseFloat(neum.kmActual)  || 0;
        var kmRec      = Math.max(0, kmAct - kmIni);
        var kmLimite   = parseFloat(neum.kmLimite)  || 120000;
        var vidaUtil   = kmLimite > 0 ? Math.round(kmRec / kmLimite * 100) : 0;
        var kmRestante = Math.max(0, kmLimite - kmRec);
        return {
          ok: true,
          data: {
            id:          id,
            kmRecorrido: kmRec,
            kmRestante:  kmRestante,
            vidaUtil:    vidaUtil,  // porcentaje consumido
            alertaRetiro: kmRec >= kmLimite * 0.90 // alerta al 90%
          }
        };
      }
    }
    throw new Error('Neumático no encontrado: ' + id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  obtenerHistorialNeumatico(id)
// ─────────────────────────────────────────────────────────────
function obtenerHistorialNeumatico(id) {
  try {
    if (!id) throw new Error('ID requerido');
    var data      = _sheetNeumHist().getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_NEUM_H.ID_NEUM] == id) {
        resultado.push({
          id:        data[i][MC_NEUM_H.ID],
          fecha:     data[i][MC_NEUM_H.FECHA],
          placa:     data[i][MC_NEUM_H.PLACA],
          posicion:  data[i][MC_NEUM_H.POSICION],
          accion:    data[i][MC_NEUM_H.ACCION],
          km:        data[i][MC_NEUM_H.KM],
          profundidad:data[i][MC_NEUM_H.PROFUNDIDAD],
          presion:   data[i][MC_NEUM_H.PRESION],
          detalle:   data[i][MC_NEUM_H.DETALLE]
        });
      }
    }
    resultado.sort(function(a, b) { return b.fecha > a.fecha ? 1 : -1; });
    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function _agregarHistNeum(p) {
  var sh  = _sheetNeumHist();
  var id  = _nextIdNeumH();
  var now = new Date();
  var hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow([
    id, hoy, p.idNeum, p.serie, String(p.placa || '').toUpperCase(),
    p.posicion || '', p.accion || '', p.km || '', p.profundidad || '', p.presion || '',
    p.detalle  || '', p.usuario || 'SISTEMA', now.toISOString()
  ]);
}
