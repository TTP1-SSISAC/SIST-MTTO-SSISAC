// ============================================================
//  AlertasMantenimiento.gs  —  SIST-MTTO-SSISAC
//  Generación y gestión de alertas automáticas de mantenimiento
// ============================================================

var SHEET_ALERT = 'ALERTAS_MTTO';

var MC_ALERT = {
  ID:           0,   // ALERT-YYYY-NNNN
  FECHA_GEN:    1,
  TIPO:         2,   // PREVENTIVO_PROXIMO | PREVENTIVO_VENCIDO | KM_PROXIMO |
                     // NEUMATICO_DESGASTE | NEUMATICO_KM | BATERIA_VOLTAJE |
                     // BATERIA_VENCIMIENTO | OT_VENCIDA | INSPECCION_PENDIENTE
  PRIORIDAD:    3,   // ALTA | MEDIA | BAJA
  ID_VEHICULO:  4,
  PLACA:        5,
  DESCRIPCION:  6,
  DETALLE:      7,   // Detalles adicionales (JSON o texto)
  ESTADO:       8,   // ACTIVA | ATENDIDA | DESCARTADA
  FECHA_LIMITE: 9,   // Fecha hasta la que es relevante
  ID_REFERENCIA:10,  // ID del plan, OT, neumático, batería, etc.
  TIPO_REF:     11,  // PLAN | OT | NEUMATICO | BATERIA | PROGRAMACION
  ATENDIDA_POR: 12,
  FECHA_ATENCION:13,
  OBSERVACIONES:14,
  TIMESTAMP:    15
};

var CABECERA_ALERT = [
  'ID_ALERTA','FECHA_GENERACION','TIPO_ALERTA','PRIORIDAD',
  'ID_VEHICULO','PLACA','DESCRIPCION','DETALLE','ESTADO','FECHA_LIMITE',
  'ID_REFERENCIA','TIPO_REFERENCIA','ATENDIDA_POR','FECHA_ATENCION',
  'OBSERVACIONES','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_ALERT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ALERT);
    sh.appendRow(CABECERA_ALERT);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_ALERT.length)
      .setBackground('#7A0000').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdAlert() {
  var sh   = _sheetAlert();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'ALERT-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('ALERT-' + yr); });
  if (!ids.length) return 'ALERT-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'ALERT-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _crearAlerta(tipo, prioridad, idVehiculo, placa, descripcion, detalle, fechaLimite, idRef, tipoRef) {
  var sh  = _sheetAlert();
  var id  = _nextIdAlert();
  var now = new Date();
  sh.appendRow([
    id,
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    tipo,
    prioridad,
    idVehiculo  || '',
    String(placa || '').toUpperCase(),
    descripcion || '',
    detalle     || '',
    'ACTIVA',
    fechaLimite || '',
    idRef       || '',
    tipoRef     || '',
    '', '',
    '',
    now.toISOString()
  ]);
  return id;
}

function _alertaExiste(tipo, idRef) {
  var data = _sheetAlert().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][MC_ALERT.TIPO]         === tipo  &&
        data[i][MC_ALERT.ID_REFERENCIA] == idRef &&
        data[i][MC_ALERT.ESTADO]        === 'ACTIVA') {
      return true;
    }
  }
  return false;
}

// ─────────────────────────────────────────────────────────────
//  generarAlertas()
//  Función principal — revisa toda la flota y genera alertas
//  Diseñada para ejecutarse como trigger diario
// ─────────────────────────────────────────────────────────────
function generarAlertas() {
  try {
    var generadas  = 0;
    var hoy        = new Date();

    // 1. Alertas de planes preventivos próximos o vencidos
    try {
      var shPrev = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PREV);
      if (shPrev) {
        var dataPrev = shPrev.getDataRange().getValues();
        for (var i = 1; i < dataPrev.length; i++) {
          var row = dataPrev[i];
          if (!row[MC_PREV.ID] || row[MC_PREV.ESTADO] === 'INACTIVO') continue;

          var fProx = row[MC_PREV.FECHA_PROXIMA] ? new Date(row[MC_PREV.FECHA_PROXIMA]) : null;
          if (!fProx || isNaN(fProx.getTime())) continue;

          var diasRest = Math.ceil((fProx - hoy) / 86400000);
          var idPlan   = row[MC_PREV.ID];

          if (diasRest < 0 && !_alertaExiste('PREVENTIVO_VENCIDO', idPlan)) {
            _crearAlerta('PREVENTIVO_VENCIDO', 'ALTA',
              row[MC_PREV.ID_VEHICULO], row[MC_PREV.PLACA],
              'Mantenimiento preventivo VENCIDO: ' + row[MC_PREV.TIPO_SERVICIO],
              Math.abs(diasRest) + ' días vencido',
              Utilities.formatDate(fProx, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
              idPlan, 'PLAN');
            generadas++;
          } else if (diasRest >= 0 && diasRest <= 7 && !_alertaExiste('PREVENTIVO_PROXIMO', idPlan)) {
            _crearAlerta('PREVENTIVO_PROXIMO', 'MEDIA',
              row[MC_PREV.ID_VEHICULO], row[MC_PREV.PLACA],
              'Mantenimiento preventivo próximo: ' + row[MC_PREV.TIPO_SERVICIO],
              diasRest + ' días restantes',
              Utilities.formatDate(fProx, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
              idPlan, 'PLAN');
            generadas++;
          }
        }
      }
    } catch (ex) { Logger.log('Error alertas preventivo: ' + ex.message); }

    // 2. Alertas de neumáticos con alto desgaste
    try {
      var shNeum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NEUM);
      if (shNeum) {
        var dataNeum = shNeum.getDataRange().getValues();
        for (var j = 1; j < dataNeum.length; j++) {
          var rn = dataNeum[j];
          if (!rn[MC_NEUM.ID] || rn[MC_NEUM.ESTADO] !== 'EN_USO') continue;

          var kmRec  = parseFloat(rn[MC_NEUM.KM_RECORRIDO]) || 0;
          var kmLim  = parseFloat(rn[MC_NEUM.KM_LIMITE])    || 120000;
          var vidaUt = kmLim > 0 ? kmRec / kmLim : 0;

          if (vidaUt >= 0.90 && !_alertaExiste('NEUMATICO_KM', rn[MC_NEUM.ID])) {
            _crearAlerta('NEUMATICO_KM', vidaUt >= 1 ? 'ALTA' : 'MEDIA',
              rn[MC_NEUM.ID_VEHICULO], rn[MC_NEUM.PLACA],
              'Neumático al ' + Math.round(vidaUt * 100) + '% de vida útil',
              'Serie: ' + rn[MC_NEUM.SERIE] + ', Posición: ' + rn[MC_NEUM.POSICION] +
              ', KM recorrido: ' + kmRec,
              '', rn[MC_NEUM.ID], 'NEUMATICO');
            generadas++;
          }

          // Alerta por profundidad de labrado < 4mm
          var prof = parseFloat(rn[MC_NEUM.PROFUNDIDAD_MM]) || 0;
          if (prof > 0 && prof < 4 && !_alertaExiste('NEUMATICO_DESGASTE', rn[MC_NEUM.ID])) {
            _crearAlerta('NEUMATICO_DESGASTE', 'ALTA',
              rn[MC_NEUM.ID_VEHICULO], rn[MC_NEUM.PLACA],
              'Neumático con profundidad crítica: ' + prof + ' mm',
              'Serie: ' + rn[MC_NEUM.SERIE] + ', Posición: ' + rn[MC_NEUM.POSICION],
              '', rn[MC_NEUM.ID], 'NEUMATICO');
            generadas++;
          }
        }
      }
    } catch (ex) { Logger.log('Error alertas neumáticos: ' + ex.message); }

    // 3. Alertas de baterías vencidas o próximas a vencer
    try {
      var shBat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BAT);
      if (shBat) {
        var dataBat = shBat.getDataRange().getValues();
        for (var k = 1; k < dataBat.length; k++) {
          var rb = dataBat[k];
          if (!rb[MC_BAT.ID] || rb[MC_BAT.ESTADO] !== 'EN_USO') continue;

          var fVenc = rb[MC_BAT.FECHA_VENC] ? new Date(rb[MC_BAT.FECHA_VENC]) : null;
          if (fVenc && !isNaN(fVenc.getTime())) {
            var diasVenc = Math.ceil((fVenc - hoy) / 86400000);
            if (diasVenc < 0 && !_alertaExiste('BATERIA_VENCIMIENTO', rb[MC_BAT.ID])) {
              _crearAlerta('BATERIA_VENCIMIENTO', 'ALTA',
                rb[MC_BAT.ID_VEHICULO], rb[MC_BAT.PLACA],
                'Batería VENCIDA: ' + rb[MC_BAT.SERIE],
                Math.abs(diasVenc) + ' días vencida. Marca: ' + rb[MC_BAT.MARCA],
                '', rb[MC_BAT.ID], 'BATERIA');
              generadas++;
            } else if (diasVenc >= 0 && diasVenc <= 30 && !_alertaExiste('BATERIA_VENCIMIENTO', rb[MC_BAT.ID])) {
              _crearAlerta('BATERIA_VENCIMIENTO', 'MEDIA',
                rb[MC_BAT.ID_VEHICULO], rb[MC_BAT.PLACA],
                'Batería próxima a vencer: ' + rb[MC_BAT.SERIE],
                diasVenc + ' días restantes. Marca: ' + rb[MC_BAT.MARCA],
                Utilities.formatDate(fVenc, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
                rb[MC_BAT.ID], 'BATERIA');
              generadas++;
            }
          }
        }
      }
    } catch (ex) { Logger.log('Error alertas baterías: ' + ex.message); }

    return { ok: true, data: { generadas: generadas, fecha: Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
                                mensaje: generadas + ' alertas generadas' } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  obtenerAlertas(filtros)
//  filtros: placa, tipo, prioridad, estado
// ─────────────────────────────────────────────────────────────
function obtenerAlertas(filtros) {
  try {
    filtros       = filtros || {};
    var data      = _sheetAlert().getDataRange().getValues();
    var resultado = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[MC_ALERT.ID]) continue;

      if (filtros.placa     && row[MC_ALERT.PLACA]    !== String(filtros.placa).toUpperCase()) continue;
      if (filtros.tipo      && row[MC_ALERT.TIPO]     !== filtros.tipo.toUpperCase()) continue;
      if (filtros.prioridad && row[MC_ALERT.PRIORIDAD] !== filtros.prioridad.toUpperCase()) continue;
      if (filtros.estado    && row[MC_ALERT.ESTADO]   !== filtros.estado.toUpperCase()) continue;

      // Por defecto solo activas
      if (!filtros.estado && row[MC_ALERT.ESTADO] !== 'ACTIVA') continue;

      resultado.push({
        id:          row[MC_ALERT.ID],
        fechaGen:    row[MC_ALERT.FECHA_GEN],
        tipo:        row[MC_ALERT.TIPO],
        prioridad:   row[MC_ALERT.PRIORIDAD],
        placa:       row[MC_ALERT.PLACA],
        descripcion: row[MC_ALERT.DESCRIPCION],
        detalle:     row[MC_ALERT.DETALLE],
        estado:      row[MC_ALERT.ESTADO],
        fechaLimite: row[MC_ALERT.FECHA_LIMITE],
        idRef:       row[MC_ALERT.ID_REFERENCIA],
        tipoRef:     row[MC_ALERT.TIPO_REF]
      });
    }

    // Ordenar: ALTA primero, luego MEDIA, luego BAJA, y por fecha
    var ord = { 'ALTA': 0, 'MEDIA': 1, 'BAJA': 2 };
    resultado.sort(function(a, b) {
      var pa = ord[a.prioridad] !== undefined ? ord[a.prioridad] : 9;
      var pb = ord[b.prioridad] !== undefined ? ord[b.prioridad] : 9;
      if (pa !== pb) return pa - pb;
      return b.fechaGen > a.fechaGen ? 1 : -1;
    });

    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  marcarAlertaAtendida(id, usuario, observaciones)
// ─────────────────────────────────────────────────────────────
function marcarAlertaAtendida(id, usuario, observaciones) {
  try {
    if (!id) throw new Error('ID requerido');
    var sh   = _sheetAlert();
    var data = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_ALERT.ID] == id) {
        var fila = i + 1;
        var hoy  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        sh.getRange(fila, MC_ALERT.ESTADO          + 1).setValue('ATENDIDA');
        sh.getRange(fila, MC_ALERT.ATENDIDA_POR    + 1).setValue(usuario || 'SISTEMA');
        sh.getRange(fila, MC_ALERT.FECHA_ATENCION  + 1).setValue(hoy);
        if (observaciones) sh.getRange(fila, MC_ALERT.OBSERVACIONES + 1).setValue(observaciones);
        sh.getRange(fila, MC_ALERT.TIMESTAMP       + 1).setValue(new Date().toISOString());
        return { ok: true, data: { id: id, mensaje: 'Alerta marcada como atendida' } };
      }
    }
    throw new Error('Alerta no encontrada: ' + id);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
