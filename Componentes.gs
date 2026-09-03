// ===========================================================
// Componentes.gs  -  SIST-MTTO-SSISAC
// Catálogo de Componentes: crea la hoja COMPONENTES en el
// Spreadsheet y expone acciones REST para el Web App.
//
// ACCIONES (via doGet / doPost):
//   getComponentes        → devuelve todos los componentes
//   getInfoComponente     → dado {componente}, retorna {codigo, sistema}
//   crearHojaComponentes  → inicializa/rellena la hoja (setup)
// ===========================================================

var NOMBRE_HOJA_COMP = 'COMPONENTES';

// ─────────────────────────────────────────────────────────────
// DATOS MAESTROS  (121 componentes — fuente: BD_COMPONENTE-FALLA)
// ─────────────────────────────────────────────────────────────
var CATALOGO_COMPONENTES = [
  [1,  'MOTOR',                                                   'FRENO DE MOTOR'],
  [2,  'MOTOR',                                                   'NIVEL DE ACEITE DE MOTOR'],
  [3,  'MOTOR',                                                   'FILTRO DE AIRE'],
  [4,  'MOTOR',                                                   'PÉRDIDA DE POTENCIA'],
  [5,  'MOTOR',                                                   'FUGA DE FLUIDOS'],
  [6,  'MOTOR',                                                   'FAJAS, POLEAS, TEMPLADORES'],
  [7,  'MOTOR',                                                   'TURBO COMPRESOR'],
  [8,  'MOTOR',                                                   'SILENCIADOR, MATACHISPA'],
  [9,  'MOTOR',                                                   'MÚLTIPLE DE ESCAPE'],
  [10, 'MOTOR',                                                   'SOPORTE DE MOTOR'],
  [11, 'MOTOR',                                                   'SISTEMA DE COMBUSTIBLE'],
  [12, 'CAJA - CORONAS',                                         'EMBRAGUE, REGULACIÓN'],
  [13, 'CAJA - CORONAS',                                         'PALANCA DE CAMBIOS'],
  [14, 'CAJA - CORONAS',                                         'SELECTOR, SPLITER'],
  [15, 'CAJA - CORONAS',                                         'TEMPERATURA ELEVADA'],
  [16, 'CAJA - CORONAS',                                         'PTO (TOMA DE FUERZA)'],
  [17, 'CAJA - CORONAS',                                         'RUIDO EN LA CAJA DE CAMBIOS'],
  [18, 'CAJA - CORONAS',                                         'RUIDO EN LAS CORONAS'],
  [19, 'CAJA - CORONAS',                                         'RETENES, SEMIEJES DE RUEDA'],
  [20, 'CAJA - CORONAS',                                         'TEMPLADORES, SOPORTES'],
  [21, 'CAJA - CORONAS',                                         'REENVÍO'],
  [22, 'CAJA - CORONAS',                                         'CARDÁN Y CRUCETAS'],
  [23, 'DIRECCIÓN',                                              'ALINEAMIENTO Y BALANCEO'],
  [24, 'DIRECCIÓN',                                              'SERVO, SIST. HIDRÁULICO'],
  [25, 'DIRECCIÓN',                                              'CAJA DE DIRECCIÓN'],
  [26, 'DIRECCIÓN',                                              'BARRA Y TERMINALES'],
  [27, 'REFRIGERACIÓN',                                          'NIVEL DE REFRIGERANTE'],
  [28, 'REFRIGERACIÓN',                                          'FUGAS DE REFRIGERANTE'],
  [29, 'REFRIGERACIÓN',                                          'TANQUE DE EXPANSIÓN'],
  [30, 'REFRIGERACIÓN',                                          'TEMPERATURA ELEVADA'],
  [31, 'REFRIGERACIÓN',                                          'RADIADOR, INTERCOOLER'],
  [32, 'REFRIGERACIÓN',                                          'FAN CLUTCH, BOMBA DE AGUA'],
  [33, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'CAPOT, LITERA'],
  [34, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'LUNAS Y PARABRISAS'],
  [35, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'ASIENTOS'],
  [36, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'SUSPENSIÓN DE ASIENTOS'],
  [37, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'CINTURONES DE SEGURIDAD'],
  [38, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'TABLERO E INSTRUMENTOS'],
  [39, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'SUSPENSIÓN DE CABINA'],
  [40, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'PORTA LETRERO DE CONVOY'],
  [41, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'DEFLECTORES DE AIRE'],
  [42, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'LITERA Y AMORTIGUADORES'],
  [43, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'TANQUES DE COMBUSTIBLE'],
  [44, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'PUERTAS Y MANIJAS'],
  [45, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'TIMÓN Y ACCESORIOS'],
  [46, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'HERMETICIDAD DE CABINA'],
  [47, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'ESPEJOS LATERALES'],
  [48, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'SOPORTES DE CABINA'],
  [49, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'VÁLVULAS DE BRECKEO'],
  [50, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'CONTROL VELOC. CRUCERO'],
  [51, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'APOYABRAZOS'],
  [52, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'ACCESORIOS EN GENERAL'],
  [53, 'CABINA Y OTROS COMPONENTES DEL REMOLCADOR',             'AUTORADIO Y ANTENAS'],
  [54, 'SISTEMA ELÉCTRICO',                                     'LUCES EN GENERAL'],
  [55, 'SISTEMA ELÉCTRICO',                                     'FAROS DELANTEROS'],
  [56, 'SISTEMA ELÉCTRICO',                                     'NEBLINEROS, CIRCULINA'],
  [57, 'SISTEMA ELÉCTRICO',                                     'CLAXON, ALARMA DE RETROCESO'],
  [58, 'SISTEMA ELÉCTRICO',                                     'TRICO Y PLUMILLAS'],
  [59, 'SISTEMA ELÉCTRICO',                                     'BATERÍAS Y BORNES'],
  [60, 'SISTEMA ELÉCTRICO',                                     'TESTIGOS CHECK ENGINE'],
  [61, 'SISTEMA ELÉCTRICO',                                     'TESTIGO ABS'],
  [62, 'SISTEMA ELÉCTRICO',                                     'AIRE ACONDICIONADO'],
  [63, 'SISTEMA ELÉCTRICO',                                     'CALEFACCIÓN'],
  [64, 'SISTEMA ELÉCTRICO',                                     'CORTADOR DE CORRIENTE'],
  [65, 'FRENOS',                                                 'FRENOS, REGULACIÓN'],
  [66, 'FRENOS',                                                 'FAJAS DE FRENO'],
  [67, 'FRENOS',                                                 'COMPRESOR, LÍNEAS NEUMÁTICAS'],
  [68, 'FRENOS',                                                 'FUGAS DE AIRE'],
  [69, 'FRENOS',                                                 'SECADOR DE AIRE'],
  [70, 'FRENOS',                                                 'VÁLVULAS NEUMÁTICAS'],
  [71, 'LLANTAS',                                                'REPARACIÓN DE LLANTAS'],
  [72, 'LLANTAS',                                                'CAMBIO DE LLANTAS'],
  [73, 'LLANTAS',                                                'ROTACIÓN DE LLANTAS'],
  [74, 'LLANTAS',                                                'PRESIÓN DE AIRE'],
  [75, 'LLANTAS',                                                'SEGURO DE TUERCAS'],
  [76, 'LLANTAS',                                                'LLANTA DE REPUESTO'],
  [77, 'COMPONENTES CRÍTICOS',                                   'EQUIPO DE BOMBEO (TRASIEGO)'],
  [78, 'COMPONENTES CRÍTICOS',                                   'BOMBA, SIST. HIDRÁULICO'],
  [79, 'COMPONENTES CRÍTICOS',                                   'QUINTA RUEDA, KING PIN'],
  [80, 'COMPONENTES CRÍTICOS',                                   'ESCALERAS Y BARANDAS'],
  [81, 'COMPONENTES CRÍTICOS',                                   'TREN DE APOYO'],
  [82, 'COMPONENTES CRÍTICOS',                                   'FUNCIONAMIENTO DEL GPS'],
  [83, 'SUSPENSIÓN',                                             'AMORTIGUADORES'],
  [84, 'SUSPENSIÓN',                                             'BOLSAS DE AIRE'],
  [85, 'SUSPENSIÓN',                                             'REG. DE BOLSAS DE AIRE'],
  [86, 'SUSPENSIÓN',                                             'MUELLES Y GRILLETES'],
  [87, 'SUSPENSIÓN',                                             'ABRAZADERAS Y BUJES'],
  [88, 'SUSPENSIÓN',                                             'TEMPLADOR, BALANCINES'],
  [89, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'PARACHOQUES'],
  [90, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'TAPABARROS'],
  [91, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'ESCARPINES'],
  [92, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'CAJA DE HERRAMIENTAS'],
  [93, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'BOCAMAZAS Y RODAMIENTOS'],
  [94, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'LUBRICACIÓN, ENGRASE'],
  [95, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'AROS, ESPÁRRAGOS Y TUERCAS'],
  [96, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'CHASIS'],
  [97, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'PORTA CONOS'],
  [98, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'PORTA TACOS'],
  [99, 'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'PLACAS DE RODAJE'],
  [100,'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'BARRA DE ANTIEMPOTRAMIENTO'],
  [101,'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'PORTA LLANTAS Y TECLES'],
  [102,'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'PORTA EXTINTORES'],
  [103,'OTROS COMPONENTES DEL REMOLCADOR Y/O SEMIRREMOLQUE',    'SEÑALÉTICA Y PINTURA'],
  [104,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'CISTERNA DE COMBUSTIBLE'],
  [105,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'ACCESORIOS DE DESCARGA'],
  [106,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'CAJA DE VÁLVULAS'],
  [107,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'MANHOLE, VÁLVULAS, SCULLY'],
  [108,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'LÍNEA DE VIDA, ANTIDESLIZANTE'],
  [109,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'SIS. TOMA A TIERRA'],
  [110,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'VÁLVULAS API'],
  [111,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'VÁLVULA DE ACCIONAMIENTO INTERLOCK'],
  [112,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'PISTONES NEUMÁTICOS'],
  [113,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'VÁLVULA RECUPERADOR DE VAPORES'],
  [114,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'TAPA MANHOLE'],
  [115,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'VÁLVULA DE EMERGENCIA'],
  [116,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'PISO ANTIDESLIZANTE'],
  [117,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'LÍNEAS DE VIDA'],
  [118,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'VÁLVULAS Y MANGUERAS'],
  [119,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'MALACATES'],
  [120,'OTROS COMPONENTES DEL SEMIRREMOLQUE',                    'VIBRADORES'],
  [200,'GENERAL',                                                'OTROS']
];

// ─────────────────────────────────────────────────────────────
// SETUP — crear / poblar la hoja COMPONENTES
// Ejecutar UNA VEZ desde el editor de GAS:
//   Editor → seleccionar "crearHojaComponentes" → Ejecutar
// ─────────────────────────────────────────────────────────────
function crearHojaComponentes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(NOMBRE_HOJA_COMP);

  // Si ya existe, la limpia; si no, la crea
  if (hoja) {
    hoja.clearContents();
    Logger.log('⚠️  Hoja ' + NOMBRE_HOJA_COMP + ' existía — se limpió y se repobló.');
  } else {
    hoja = ss.insertSheet(NOMBRE_HOJA_COMP);
    Logger.log('✅  Hoja ' + NOMBRE_HOJA_COMP + ' creada.');
  }

  // ── Encabezado ──
  var encabezado = [['CÓDIGO', 'SISTEMA', 'COMPONENTE']];
  hoja.getRange(1, 1, 1, 3).setValues(encabezado);

  // Formato del encabezado
  var hdrRange = hoja.getRange(1, 1, 1, 3);
  hdrRange.setBackground('#1A3A5C')
          .setFontColor('#FFFFFF')
          .setFontWeight('bold')
          .setFontFamily('Arial')
          .setFontSize(10);

  // ── Datos ──
  hoja.getRange(2, 1, CATALOGO_COMPONENTES.length, 3)
      .setValues(CATALOGO_COMPONENTES);

  // Formato datos
  hoja.getRange(2, 1, CATALOGO_COMPONENTES.length, 3)
      .setFontFamily('Arial')
      .setFontSize(10);

  // Columna CÓDIGO: centrada y negrita
  hoja.getRange(2, 1, CATALOGO_COMPONENTES.length, 1)
      .setHorizontalAlignment('center')
      .setFontWeight('bold');

  // Auto-ajustar columnas
  hoja.autoResizeColumns(1, 3);

  // Congelar encabezado
  hoja.setFrozenRows(1);

  // Validación de datos en REPORTES_FALLA col COMPONENTE (col 9, si existe)
  _agregarValidacionComponente();

  var msg = '✅  COMPONENTES creada con ' + CATALOGO_COMPONENTES.length + ' filas.';
  Logger.log(msg);
  return { ok: true, data: { filas: CATALOGO_COMPONENTES.length } };
}

/**
 * Agrega validación desplegable en la columna COMPONENTE de REPORTES_FALLA.
 * Si la hoja no existe aún, lo omite silenciosamente.
 */
function _agregarValidacionComponente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rf = ss.getSheetByName('REPORTES_FALLA');
  if (!rf) return;

  var nombresComp = CATALOGO_COMPONENTES.map(function(r){ return r[2]; });
  var regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(nombresComp, true)
    .setAllowInvalid(false)
    .build();

  // Col 9 = COMPONENTE, fila 2 en adelante (hasta 500)
  rf.getRange(2, 9, 500).setDataValidation(regla);
  Logger.log('  ↳ Validación desplegable agregada en REPORTES_FALLA[col 9].');
}

// ─────────────────────────────────────────────────────────────
// ACCIONES REST — llamadas desde el Web App (doGet / doPost)
// ─────────────────────────────────────────────────────────────

/**
 * getComponentes()
 * Devuelve todos los componentes del catálogo interno.
 * Respuesta: { ok: true, data: [ {codigo, sistema, componente}, ... ] }
 */
function actionGetComponentes() {
  var lista = CATALOGO_COMPONENTES.map(function(r) {
    return { codigo: r[0], sistema: r[1], componente: r[2] };
  });
  return { ok: true, data: lista };
}

/**
 * getInfoComponente(params)
 * Dado params.componente (nombre exacto), retorna {codigo, sistema}.
 * Si no se encuentra, retorna ok:false con mensaje de error.
 *
 * Ejemplo de llamada desde el frontend:
 *   gasCall('getInfoComponente', { componente: 'TURBO COMPRESOR' })
 *   → { ok: true, data: { codigo: 7, sistema: 'MOTOR' } }
 */
function actionGetInfoComponente(params) {
  var nombre = (params && params.componente) ? params.componente.toString().trim().toUpperCase() : '';

  if (!nombre) {
    return { ok: false, error: 'Parámetro "componente" requerido.' };
  }

  for (var i = 0; i < CATALOGO_COMPONENTES.length; i++) {
    var fila = CATALOGO_COMPONENTES[i];
    if (fila[2].toUpperCase() === nombre) {
      return { ok: true, data: { codigo: fila[0], sistema: fila[1], componente: fila[2] } };
    }
  }

  return { ok: false, error: 'Componente "' + nombre + '" no encontrado en el catálogo.' };
}

/**
 * getSistemas()
 * Devuelve la lista única de sistemas (para poblar el primer select).
 * Respuesta: { ok: true, data: ['MOTOR', 'FRENOS', ...] }
 */
function actionGetSistemas() {
  var vistos = {};
  var sistemas = [];
  CATALOGO_COMPONENTES.forEach(function(r) {
    if (!vistos[r[1]]) { vistos[r[1]] = true; sistemas.push(r[1]); }
  });
  return { ok: true, data: sistemas };
}

/**
 * getComponentesPorSistema(params)
 * Dado params.sistema, retorna solo los componentes de ese sistema.
 * Respuesta: { ok: true, data: [{codigo, componente}, ...] }
 */
function actionGetComponentesPorSistema(params) {
  var sis = (params && params.sistema) ? params.sistema.toString().trim().toUpperCase() : '';
  if (!sis) return { ok: false, error: 'Parámetro "sistema" requerido.' };

  var result = [];
  CATALOGO_COMPONENTES.forEach(function(r) {
    if (r[1].toUpperCase() === sis) {
      result.push({ codigo: r[0], componente: r[2] });
    }
  });

  if (!result.length) return { ok: false, error: 'Sistema "' + sis + '" no encontrado.' };
  return { ok: true, data: result };
}

// ─────────────────────────────────────────────────────────────
// REGISTRO DE REPORTE DE FALLA
// Guarda una fila en REPORTES_FALLA con toda la info del componente
// autocompletada (codigo, sistema).
//
// params esperados:
//   { usuario, fecha, placa, kilometraje, personal, ndoc,
//     componente, detalle, estado }
//
// El script busca codigo y sistema automáticamente desde CATALOGO.
// ─────────────────────────────────────────────────────────────
function actionRegistrarFalla(params) {
  if (!params) return { ok: false, error: 'Sin parámetros.' };

  // Buscar info del componente
  var info = actionGetInfoComponente({ componente: params.componente });
  var codigo  = info.ok ? info.data.codigo  : '';
  var sistema = info.ok ? info.data.sistema : (params.sistema || '');

  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('REPORTES_FALLA');
  if (!hoja) return { ok: false, error: 'Hoja REPORTES_FALLA no encontrada.' };

  var usuario = Session.getActiveUser().getEmail() || 'WEB_APP';
  var fila = [
    usuario,
    params.fecha        || new Date(),
    params.placa        || '',
    params.kilometraje  || '',
    params.personal     || '',
    params.ndoc         || '',
    codigo,
    sistema,
    params.componente   || '',
    params.detalle      || '',
    params.estado       || 'PENDIENTE'
  ];

  hoja.appendRow(fila);

  Logger.log('✅  Falla registrada: ' + JSON.stringify(fila));
  return {
    ok: true,
    data: {
      codigo:     codigo,
      sistema:    sistema,
      componente: params.componente,
      fila:       hoja.getLastRow()
    }
  };
}

// ─────────────────────────────────────────────────────────────
// PUNTO DE ENTRADA DEL WEB APP
// Agrega estas acciones al doGet/doPost central (Code.gs o Main.gs).
// Si aún no tienes un despachador central, puedes usar este como base.
// ─────────────────────────────────────────────────────────────

/**
 * Despachador de componentes — llama desde tu doGet/doPost central así:
 *
 *   var resultado = despacharComponentes(action, params);
 *   if (resultado !== null) return resultado;  // acción manejada
 *
 * Retorna null si la acción no es de este módulo.
 */
function despacharComponentes(action, params) {
  switch (action) {
    case 'getComponentes':           return actionGetComponentes();
    case 'getSistemas':              return actionGetSistemas();
    case 'getInfoComponente':        return actionGetInfoComponente(params);
    case 'getComponentesPorSistema': return actionGetComponentesPorSistema(params);
    case 'registrarFalla':           return actionRegistrarFalla(params);
    default:                         return null;
  }
}

// ─────────────────────────────────────────────────────────────
// PRUEBA RÁPIDA — ejecutar desde el editor para verificar
// ─────────────────────────────────────────────────────────────
function testComponentes() {
  Logger.log('--- getInfoComponente(TURBO COMPRESOR) ---');
  Logger.log(JSON.stringify(actionGetInfoComponente({ componente: 'TURBO COMPRESOR' })));
  // Esperado: { ok:true, data:{ codigo:7, sistema:'MOTOR', componente:'TURBO COMPRESOR' } }

  Logger.log('--- getComponentesPorSistema(FRENOS) ---');
  Logger.log(JSON.stringify(actionGetComponentesPorSistema({ sistema: 'FRENOS' })));

  Logger.log('--- getSistemas ---');
  var sis = actionGetSistemas();
  Logger.log('Sistemas (' + sis.data.length + '): ' + sis.data.join(', '));

  Logger.log('--- getComponentes total ---');
  Logger.log('Total: ' + actionGetComponentes().data.length + ' componentes');
}
