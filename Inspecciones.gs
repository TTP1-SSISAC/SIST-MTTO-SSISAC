// ============================================================
//  Inspecciones.gs  —  SIST-MTTO-SSISAC
//  Inspecciones periódicas y checklists de vehículos
// ============================================================

var SHEET_INSP = 'INSPECCIONES';

var MC_INSP = {
  ID:           0,   // INSP-YYYY-NNNN
  FECHA_REG:    1,
  ID_VEHICULO:  2,
  PLACA:        3,
  TIPO_INSP:    4,   // PREOPERACIONAL | SEMANAL | MENSUAL | TECNICOMECANICA
  ESTADO:       5,   // EN_PROCESO | APROBADA | OBSERVADA | RECHAZADA
  TECNICO:      6,
  KM_ACTUAL:    7,
  // Items de checklist (columnas 8-25)
  MOTOR:        8,   // OK | OBSERVADO | FALLA
  FRENOS:       9,
  LLANTAS:      10,
  LUCES:        11,
  SUSPENSION:   12,
  DIRECCION:    13,
  TRANSMISION:  14,
  COMBUSTIBLE:  15,
  ACEITE:       16,
  REFRIGERANTE: 17,
  BATERIA:      18,
  EXTINTOR:     19,
  DOCUMENTOS:   20,
  CARROCERIA:   21,
  OBSERVACIONES:22,
  ID_OT:        23,  // OT generada si hay fallas
  USUARIO:      24,
  TIMESTAMP:    25
};

var CABECERA_INSP = [
  'ID_INSPECCION','FECHA_REGISTRO','ID_VEHICULO','PLACA','TIPO_INSPECCION',
  'ESTADO','TECNICO','KM_ACTUAL',
  'MOTOR','FRENOS','LLANTAS','LUCES','SUSPENSION','DIRECCION','TRANSMISION',
  'COMBUSTIBLE','ACEITE','REFRIGERANTE','BATERIA','EXTINTOR','DOCUMENTOS','CARROCERIA',
  'OBSERVACIONES','ID_OT_GENERADA','USUARIO','TIMESTAMP'
];

var ITEMS_CHECKLIST = [
  'motor','frenos','llantas','luces','suspension','direccion',
  'transmision','combustible','aceite','refrigerante','bateria',
  'extintor','documentos','carroceria'
];

// ─────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────
function _sheetInsp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_INSP);
  if (!sh) {
    sh = ss.insertSheet(SHEET_INSP);
    sh.appendRow(CABECERA_INSP);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECERA_INSP.length)
      .setBackground('#005B4A').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sh;
}

function _nextIdInsp() {
  var sh   = _sheetInsp();
  var yr   = new Date().getFullYear();
  var last = sh.getLastRow();
  if (last < 2) return 'INSP-' + yr + '-0001';
  var ids  = sh.getRange(2, 1, last - 1, 1).getValues()
               .map(function(r){ return r[0]; })
               .filter(function(v){ return String(v).startsWith('INSP-' + yr); });
  if (!ids.length) return 'INSP-' + yr + '-0001';
  var nums = ids.map(function(id){ return parseInt(id.split('-')[2]) || 0; });
  return 'INSP-' + yr + '-' + String(Math.max.apply(null, nums) + 1).padStart(4, '0');
}

// ─────────────────────────────────────────────────────────────
//  obtenerChecklist(tipoInspeccion)
//  Retorna el checklist estándar para el tipo de inspección
// ─────────────────────────────────────────────────────────────
function obtenerChecklist(tipoInspeccion) {
  var checklist = {
    PREOPERACIONAL: ['motor','frenos','llantas','luces','combustible','aceite',
                     'refrigerante','extintor','documentos'],
    SEMANAL:        ['motor','frenos','llantas','luces','suspension','combustible',
                     'aceite','refrigerante','bateria','extintor','documentos','carroceria'],
    MENSUAL:        ITEMS_CHECKLIST,
    TECNICOMECANICA:ITEMS_CHECKLIST
  };

  var tipo = String(tipoInspeccion || 'PREOPERACIONAL').toUpperCase();
  return {
    ok: true,
    data: {
      tipo:    tipo,
      items:   checklist[tipo] || checklist['PREOPERACIONAL'],
      estados: ['OK', 'OBSERVADO', 'FALLA']
    }
  };
}

// ─────────────────────────────────────────────────────────────
//  registrarInspeccion(params)
//  params: idVehiculo, placa, tipoInsp, tecnico, kmActual,
//          checklist: { motor, frenos, llantas, ... },
//          observaciones, usuario
// ─────────────────────────────────────────────────────────────
function registrarInspeccion(params) {
  try {
    if (!params.placa)   throw new Error('Placa requerida');
    if (!params.tipoInsp) throw new Error('Tipo de inspección requerido');

    var cl  = params.checklist || {};
    var sh  = _sheetInsp();
    var id  = _nextIdInsp();
    var now = new Date();
    var hoy = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Determinar estado general de la inspección
    var hasFalla    = false;
    var hasObservado = false;
    ITEMS_CHECKLIST.forEach(function(item) {
      var v = String(cl[item] || 'OK').toUpperCase();
      if (v === 'FALLA')     hasFalla     = true;
      if (v === 'OBSERVADO') hasObservado = true;
    });

    var estadoInsp = hasFalla ? 'RECHAZADA' : (hasObservado ? 'OBSERVADA' : 'APROBADA');

    sh.appendRow([
      id,
      hoy,
      params.idVehiculo || '',
      String(params.placa).toUpperCase(),
      String(params.tipoInsp).toUpperCase(),
      estadoInsp,
      params.tecnico  || '',
      params.kmActual || '',
      // Checklist items
      String(cl.motor        || 'OK').toUpperCase(),
      String(cl.frenos       || 'OK').toUpperCase(),
      String(cl.llantas      || 'OK').toUpperCase(),
      String(cl.luces        || 'OK').toUpperCase(),
      String(cl.suspension   || 'OK').toUpperCase(),
      String(cl.direccion    || 'OK').toUpperCase(),
      String(cl.transmision  || 'OK').toUpperCase(),
      String(cl.combustible  || 'OK').toUpperCase(),
      String(cl.aceite       || 'OK').toUpperCase(),
      String(cl.refrigerante || 'OK').toUpperCase(),
      String(cl.bateria      || 'OK').toUpperCase(),
      String(cl.extintor     || 'OK').toUpperCase(),
      String(cl.documentos   || 'OK').toUpperCase(),
      String(cl.carroceria   || 'OK').toUpperCase(),
      params.observaciones || '',
      '', // ID_OT_GENERADA — se llena si se generó una OT
      params.usuario || 'SISTEMA',
      now.toISOString()
    ]);

    // Si hay fallas, generar reporte/OT automáticamente
    var idOT = '';
    if (hasFalla && params.generarOT !== false) {
      var fallaItems = ITEMS_CHECKLIST.filter(function(item) {
        return String(cl[item] || '').toUpperCase() === 'FALLA';
      });
      var otRes = generarReporteFalla({
        id:        id,
        placa:     params.placa,
        idVehiculo:params.idVehiculo,
        fallas:    fallaItems,
        tecnico:   params.tecnico,
        usuario:   params.usuario,
        observaciones: params.observaciones
      });
      if (otRes.ok) {
        idOT = otRes.data.idOT;
        // Actualizar la inspección con el ID de OT
        var data = sh.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (data[i][MC_INSP.ID] === id) {
            sh.getRange(i + 1, MC_INSP.ID_OT + 1).setValue(idOT);
            break;
          }
        }
      }
    }

    return { ok: true, data: { id: id, estado: estadoInsp, idOT: idOT,
                                hasFalla: hasFalla, hasObservado: hasObservado,
                                mensaje: 'Inspección registrada: ' + id } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  generarReporteFalla(params)
//  Crea una OT correctiva a partir de fallas detectadas
// ─────────────────────────────────────────────────────────────
function generarReporteFalla(params) {
  try {
    if (!params.placa)  throw new Error('Placa requerida');
    if (!params.fallas || !params.fallas.length) throw new Error('Lista de fallas requerida');

    var descripcion = 'FALLAS EN INSPECCIÓN ' + params.id + ': ' +
                      params.fallas.map(function(f){ return f.toUpperCase(); }).join(', ');

    var otRes = crearOrdenTrabajo({
      idVehiculo:  params.idVehiculo,
      placa:       params.placa,
      tipo:        'CORRECTIVO',
      descripcion: descripcion,
      prioridad:   'ALTA',
      tecnico:     params.tecnico || '',
      usuario:     params.usuario || 'SISTEMA',
      observaciones: 'Generado automáticamente por inspección. ' + (params.observaciones || '')
    });

    if (!otRes.ok) throw new Error(otRes.error);
    return { ok: true, data: { idOT: otRes.data.id, descripcion: descripcion,
                                mensaje: 'OT correctiva generada: ' + otRes.data.id } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
