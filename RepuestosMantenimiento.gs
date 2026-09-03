// ============================================================
//  RepuestosMantenimiento.gs  —  SIST-MTTO-SSISAC
//  Repuestos utilizados en Órdenes de Trabajo
//  Se conecta con SALIDAS → KARDEX del módulo de almacén
// ============================================================

var SHEET_REPMT = 'REPUESTOS_OT';

var MC_REPMT = {
  ID:          0,   // REPT-YYYY-NNNN
  FECHA_REG:   1,
  ID_OT:       2,
  ID_VEHICULO: 3,
  PLACA:       4,
  ID_REPUESTO: 5,   // Código del ítem en KARDEX
  DESCRIPCION: 6,
  CANTIDAD:    7,
  UNIDAD:      8,
  COSTO_UNIT:  9,
  COSTO_TOTAL: 10,
  ID_SALIDA:   11,  // Referencia a la salida de almacén
  TECNICO:     12,
  OBSERVACIONES:13,
  USUARIO:     14,
  TIMESTAMP:   15
};

var CABECERA_REPMT = [
  'ID_REPUESTO_OT','FECHA_REGISTRO','ID_OT','ID_VEHICULO','PLACA',
  'CODIGO_REPUESTO','DESCRIPCION_REPUESTO','CANTIDAD','UNIDAD',
  'COSTO_UNITARIO','COSTO_TOTAL','ID_SALIDA_ALMACEN',
  'TECNICO','OBSERVACIONES','USUARIO','TIMESTAMP'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetRepMt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_REPMT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_REPMT);
    sh.appendRow(CABECERA_REPMT);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_REPMT.length)
      .setBackground('#2D4A00').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdRepMt() {
  var sh   = _sheetRepMt();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'REPT-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('REPT-' + yr); });
  if (!ids.length) return 'REPT-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'REPT-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

function _rowToRepMt(row) {
  return {
    id:           row[MC_REPMT.ID],
    fechaReg:     row[MC_REPMT.FECHA_REG],
    idOT:         row[MC_REPMT.ID_OT],
    idVehiculo:   row[MC_REPMT.ID_VEHICULO],
    placa:        row[MC_REPMT.PLACA],
    idRepuesto:   row[MC_REPMT.ID_REPUESTO],
    descripcion:  row[MC_REPMT.DESCRIPCION],
    cantidad:     row[MC_REPMT.CANTIDAD],
    unidad:       row[MC_REPMT.UNIDAD],
    costoUnit:    row[MC_REPMT.COSTO_UNIT],
    costoTotal:   row[MC_REPMT.COSTO_TOTAL],
    idSalida:     row[MC_REPMT.ID_SALIDA],
    tecnico:      row[MC_REPMT.TECNICO],
    observaciones:row[MC_REPMT.OBSERVACIONES],
    usuario:      row[MC_REPMT.USUARIO],
    timestamp:    row[MC_REPMT.TIMESTAMP]
  };
}

// ─────────────────────────────────────────────────────────────
//  validarStock(idRepuesto, cantidadRequerida)
//  Verifica stock disponible en KARDEX antes de retirar
// ─────────────────────────────────────────────────────────────
function validarStock(idRepuesto, cantidadRequerida) {
  try {
    if (!idRepuesto) throw new Error('Código de repuesto requerido');
    cantidadRequerida = parseFloat(cantidadRequerida) || 1;

    // Buscar el ítem en KARDEX
    var ssKard = SpreadsheetApp.getActiveSpreadsheet();
    var shKard = ssKard.getSheetByName('KARDEX');
    if (!shKard) throw new Error('Hoja KARDEX no encontrada');

    var data = shKard.getDataRange().getValues();
    // KARDEX estructura esperada: col 0 = CODIGO, col 5 = STOCK_ACTUAL (ajustar según tu KARDEX real)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === String(idRepuesto).toUpperCase()) {
        var stock     = parseFloat(data[i][5]) || 0; // Ajustar índice de columna STOCK
        var descripcion = data[i][1] || '';
        var costoUnit   = parseFloat(data[i][4]) || 0; // Ajustar índice de columna COSTO
        return {
          ok:          true,
          data: {
            idRepuesto:  idRepuesto,
            descripcion: descripcion,
            stockActual: stock,
            suficiente:  stock >= cantidadRequerida,
            costoUnit:   costoUnit,
            deficit:     Math.max(0, cantidadRequerida - stock)
          }
        };
      }
    }
    throw new Error('Repuesto no encontrado en KARDEX: ' + idRepuesto);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  agregarRepuestoOT(params)
//  Registra el uso de un repuesto en una OT y genera la salida
//  params: idOT, idVehiculo, placa, idRepuesto, descripcion,
//          cantidad, unidad, costoUnit, tecnico, usuario
// ─────────────────────────────────────────────────────────────
function agregarRepuestoOT(params) {
  try {
    if (!params.idOT)       throw new Error('ID de OT requerido');
    if (!params.idRepuesto) throw new Error('Código de repuesto requerido');
    if (!params.cantidad)   throw new Error('Cantidad requerida');

    var cantidad  = parseFloat(params.cantidad)  || 0;
    var costoUnit = parseFloat(params.costoUnit) || 0;

    // Validar stock (si hay KARDEX disponible)
    var stockCheck = validarStock(params.idRepuesto, cantidad);
    if (stockCheck.ok && !stockCheck.data.suficiente) {
      throw new Error('Stock insuficiente. Disponible: ' + stockCheck.data.stockActual +
                      ', Requerido: ' + cantidad);
    }

    // Obtener costo unitario del KARDEX si no se proveyó
    if (!costoUnit && stockCheck.ok) {
      costoUnit = stockCheck.data.costoUnit;
    }

    var sh         = _sheetRepMt();
    var id         = _nextIdRepMt();
    var now        = new Date();
    var hoy        = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var costoTotal = Math.round(cantidad * costoUnit * 100) / 100;

    // Generar salida de almacén (si hay función registrarSalida disponible)
    var idSalida = '';
    try {
      var salidaRes = registrarSalidaAlmacen({
        idRepuesto:  params.idRepuesto,
        descripcion: params.descripcion || (stockCheck.ok ? stockCheck.data.descripcion : ''),
        cantidad:    cantidad,
        unidad:      params.unidad || 'UND',
        costoUnit:   costoUnit,
        motivo:      'OT: ' + params.idOT,
        idOT:        params.idOT,
        usuario:     params.usuario || 'SISTEMA'
      });
      if (salidaRes && salidaRes.ok) idSalida = salidaRes.data.id;
    } catch (ex) {
      // Si no existe la función de salida, continuar sin ella
      Logger.log('registrarSalidaAlmacen no disponible: ' + ex.message);
    }

    sh.appendRow([
      id,
      hoy,
      params.idOT,
      params.idVehiculo  || '',
      String(params.placa || '').toUpperCase(),
      String(params.idRepuesto).toUpperCase(),
      params.descripcion || (stockCheck.ok ? stockCheck.data.descripcion : ''),
      cantidad,
      params.unidad || 'UND',
      costoUnit,
      costoTotal,
      idSalida,
      params.tecnico  || '',
      params.observaciones || '',
      params.usuario || 'SISTEMA',
      now.toISOString()
    ]);

    return { ok: true, data: { id: id, costoTotal: costoTotal, idSalida: idSalida,
                                mensaje: 'Repuesto agregado a OT: ' + id } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  retirarRepuesto(params)
//  Alias de agregarRepuestoOT — semánticamente equivalente
// ─────────────────────────────────────────────────────────────
function retirarRepuesto(params) {
  return agregarRepuestoOT(params);
}

// ─────────────────────────────────────────────────────────────
//  obtenerRepuestosPorOT(idOT)
// ─────────────────────────────────────────────────────────────
function obtenerRepuestosPorOT(idOT) {
  try {
    if (!idOT) throw new Error('ID de OT requerido');
    var data      = _sheetRepMt().getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][MC_REPMT.ID_OT] == idOT) {
        resultado.push(_rowToRepMt(data[i]));
      }
    }
    return { ok: true, data: resultado, total: resultado.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularCostoRepuestos(idOT)
//  Suma todos los costos de repuestos de una OT
// ─────────────────────────────────────────────────────────────
function calcularCostoRepuestos(idOT) {
  try {
    var res = obtenerRepuestosPorOT(idOT);
    if (!res.ok) throw new Error(res.error);

    var total = res.data.reduce(function(s, r) {
      return s + (parseFloat(r.costoTotal) || 0);
    }, 0);

    return { ok: true, data: { idOT: idOT, costoRepuestos: Math.round(total * 100) / 100,
                                items: res.data.length } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
