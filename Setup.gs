// ===========================================================
// Setup.gs  -  SIST-MTTO-SSISAC
// Inicialización de todas las hojas del módulo de mantenimiento
// Ejecutar una sola vez (o cuando se agregue un nuevo entorno)
// ===========================================================

/**
 * inicializarHojas()
 * Crea las 13 hojas del módulo de mantenimiento con sus
 * encabezados y formato de color. Si la hoja ya existe, la omite.
 * Ejecutar desde: Editor → seleccionar "inicializarHojas" → Ejecutar
 */
function inicializarHojas() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var creadas = [];
  var omitidas = [];

  var hojas = [
    // Nombre              Helper                    Color header
    { nombre: 'MANTENIMIENTOS',       fn: function(){ return _sheetMant();      }, color: '#1B4F72' },
    { nombre: 'ORDENES_TRABAJO',      fn: function(){ return _sheetOT();        }, color: '#145A32' },
    { nombre: 'PLANES_PREVENTIVOS',   fn: function(){ return _sheetPrev();      }, color: '#6E2F1A' },
    { nombre: 'PROGRAMACION_MTTO',    fn: function(){ return _sheetProg();      }, color: '#4A235A' },
    { nombre: 'INTERVENCIONES',       fn: function(){ return _sheetInt();       }, color: '#1A5276' },
    { nombre: 'REPUESTOS_OT',         fn: function(){ return _sheetRepMt();     }, color: '#7D6608' },
    { nombre: 'MANO_OBRA_OT',         fn: function(){ return _sheetMO();        }, color: '#1B2631' },
    { nombre: 'INSPECCIONES',         fn: function(){ return _sheetInsp();      }, color: '#005B4A' },
    { nombre: 'NEUMATICOS',           fn: function(){ return _sheetNeum();      }, color: '#1A1A4A' },
    { nombre: 'NEUMATICOS_HISTORIAL', fn: function(){ return _sheetNeumHist();  }, color: '#2A2A5A' },
    { nombre: 'BATERIAS',             fn: function(){ return _sheetBat();       }, color: '#4A3A00' },
    { nombre: 'BATERIAS_HISTORIAL',   fn: function(){ return _sheetBatHist();   }, color: '#5A4A10' },
    { nombre: 'ALERTAS_MTTO',         fn: function(){ return _sheetAlert();     }, color: '#7A0000' }
  ];

  hojas.forEach(function(h) {
    var existia = ss.getSheetByName(h.nombre) !== null;
    try {
      h.fn(); // llama al helper – crea la hoja si no existe
      if (existia) {
        omitidas.push(h.nombre);
      } else {
        creadas.push(h.nombre);
      }
    } catch(e) {
      omitidas.push(h.nombre + ' (ERROR: ' + e.message + ')');
    }
  });

  var msg = '✅ INICIALIZACIÓN COMPLETADA\n\n';
  if (creadas.length > 0) {
    msg += 'Hojas CREADAS (' + creadas.length + '):\n  • ' + creadas.join('\n  • ') + '\n\n';
  }
  if (omitidas.length > 0) {
    msg += 'Hojas OMITIDAS (ya existían): ' + omitidas.length + '\n  • ' + omitidas.join('\n  • ');
  }

  Logger.log(msg);
  console.log(msg);

  return { ok: true, data: { creadas: creadas, omitidas: omitidas } };
}

/**
 * resetearHojasMtto()
 * ⚠️  PELIGROSO — Elimina y recrea las 13 hojas de mantenimiento.
 * Confirmar manualmente antes de ejecutar.
 */
function resetearHojasMtto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nombres = [
    'MANTENIMIENTOS','ORDENES_TRABAJO','PLANES_PREVENTIVOS','PROGRAMACION_MTTO',
    'INTERVENCIONES','REPUESTOS_OT','MANO_OBRA_OT','INSPECCIONES',
    'NEUMATICOS','NEUMATICOS_HISTORIAL','BATERIAS','BATERIAS_HISTORIAL','ALERTAS_MTTO'
  ];

  Logger.log('⚠️  RESETEAR: eliminando ' + nombres.length + ' hojas...');

  nombres.forEach(function(nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (hoja) {
      ss.deleteSheet(hoja);
      Logger.log('  ❌ Eliminada: ' + nombre);
    }
  });

  Logger.log('  Recreando hojas...');
  inicializarHojas();
  Logger.log('✅ Reset completado.');
}
