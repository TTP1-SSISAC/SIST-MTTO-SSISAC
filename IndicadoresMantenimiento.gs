// ============================================================
//  IndicadoresMantenimiento.gs  —  SIST-MTTO-SSISAC
//  KPIs de mantenimiento: MTBF, MTTR, Disponibilidad,
//  Costo/KM, Costo por vehículo, % Preventivo, % Correctivo,
//  Número de fallas, Horas de parada
// ============================================================

// ─────────────────────────────────────────────────────────────
//  Helpers internos
// ─────────────────────────────────────────────────────────────
function _filtrarMantenimientos(placa, fechaDesde, fechaHasta, tipo) {
  var filtros = {};
  if (placa)      filtros.placa      = placa;
  if (tipo)       filtros.tipo       = tipo;
  if (fechaDesde) filtros.fechaDesde = fechaDesde;
  if (fechaHasta) filtros.fechaHasta = fechaHasta;
  var res = listarMantenimientos(filtros);
  return res.ok ? res.data : [];
}

function _soloCompletados(mantenimientos) {
  return mantenimientos.filter(function(m) {
    return m.estado === 'COMPLETADO';
  });
}

// ─────────────────────────────────────────────────────────────
//  calcularMTBF(placa, fechaDesde, fechaHasta)
//  Mean Time Between Failures — Tiempo promedio entre fallas
//  Fórmula: (Tiempo total operativo) / (Nº de fallas correctivas)
//  Retorna MTBF en horas
// ─────────────────────────────────────────────────────────────
function calcularMTBF(placa, fechaDesde, fechaHasta) {
  try {
    // Obtener todos los mantenimientos del período
    var todos      = _filtrarMantenimientos(placa, fechaDesde, fechaHasta);
    var completados = _soloCompletados(todos);

    // Fallas: solo correctivos completados
    var fallas = completados.filter(function(m) { return m.tipo === 'CORRECTIVO'; });
    var nFallas = fallas.length;

    if (nFallas === 0) {
      return { ok: true, data: { mtbf: null, nFallas: 0,
               mensaje: 'Sin fallas en el período — MTBF no calculable' } };
    }

    // Tiempo total del período en horas
    var fIni  = fechaDesde ? new Date(fechaDesde) : null;
    var fFin  = fechaHasta ? new Date(fechaHasta) : new Date();
    if (!fIni && fallas.length > 0) {
      // Usar fecha del primer mantenimiento como inicio
      var fechas = fallas.map(function(f) { return new Date(f.fechaInicio || f.fechaRegistro); });
      fIni       = new Date(Math.min.apply(null, fechas));
    }

    var tiempoTotalH    = fIni ? Math.max(0, (fFin - fIni) / 3600000) : 0;

    // Restar tiempo en parada (MTTR × N)
    var totalHsParada   = fallas.reduce(function(s, m) {
      return s + (parseFloat(m.duracionH) || 0);
    }, 0);

    var tiempoOperativoH = Math.max(0, tiempoTotalH - totalHsParada);
    var mtbf             = nFallas > 0 ? Math.round(tiempoOperativoH / nFallas * 10) / 10 : null;

    return {
      ok: true,
      data: {
        placa:             placa || 'FLOTA',
        fechaDesde:        fechaDesde,
        fechaHasta:        fechaHasta || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        mtbf:              mtbf,
        mtbfUnidad:        'horas',
        nFallas:           nFallas,
        tiempoTotalH:      Math.round(tiempoTotalH * 10) / 10,
        tiempoOperativoH:  Math.round(tiempoOperativoH * 10) / 10,
        totalHsParada:     Math.round(totalHsParada * 10) / 10
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularMTTR(placa, fechaDesde, fechaHasta)
//  Mean Time To Repair — Tiempo promedio de reparación
//  Fórmula: Σ(Horas de reparación) / Nº de fallas correctivas
//  Retorna MTTR en horas
// ─────────────────────────────────────────────────────────────
function calcularMTTR(placa, fechaDesde, fechaHasta) {
  try {
    var todos    = _filtrarMantenimientos(placa, fechaDesde, fechaHasta);
    var completados = _soloCompletados(todos);
    var fallas   = completados.filter(function(m) { return m.tipo === 'CORRECTIVO'; });
    var nFallas  = fallas.length;

    if (nFallas === 0) {
      return { ok: true, data: { mttr: null, nFallas: 0,
               mensaje: 'Sin fallas en el período — MTTR no calculable' } };
    }

    var totalHsReparacion = fallas.reduce(function(s, m) {
      return s + (parseFloat(m.duracionH) || 0);
    }, 0);

    var mttr = Math.round(totalHsReparacion / nFallas * 10) / 10;

    return {
      ok: true,
      data: {
        placa:             placa || 'FLOTA',
        fechaDesde:        fechaDesde,
        fechaHasta:        fechaHasta,
        mttr:              mttr,
        mttrUnidad:        'horas',
        nFallas:           nFallas,
        totalHsReparacion: Math.round(totalHsReparacion * 10) / 10
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularDisponibilidad(placa, fechaDesde, fechaHasta)
//  Disponibilidad = MTBF / (MTBF + MTTR) × 100
//  Retorna disponibilidad como porcentaje
// ─────────────────────────────────────────────────────────────
function calcularDisponibilidad(placa, fechaDesde, fechaHasta) {
  try {
    var resMTBF = calcularMTBF(placa, fechaDesde, fechaHasta);
    var resMTTR = calcularMTTR(placa, fechaDesde, fechaHasta);

    if (!resMTBF.ok) throw new Error(resMTBF.error);
    if (!resMTTR.ok) throw new Error(resMTTR.error);

    var mtbf = resMTBF.data.mtbf;
    var mttr = resMTTR.data.mttr;

    var disponibilidad = null;
    if (mtbf !== null && mttr !== null) {
      var denom = mtbf + mttr;
      disponibilidad = denom > 0 ? Math.round((mtbf / denom) * 10000) / 100 : 100;
    } else if (resMTBF.data.nFallas === 0) {
      // Sin fallas → disponibilidad 100%
      disponibilidad = 100;
    }

    // Disponibilidad alternativa: tiempo operativo / tiempo total
    var tiempoTotalH    = resMTBF.data.tiempoTotalH    || 0;
    var tiempoOperativoH = resMTBF.data.tiempoOperativoH || tiempoTotalH;
    var dispAlternativa = tiempoTotalH > 0
      ? Math.round((tiempoOperativoH / tiempoTotalH) * 10000) / 100
      : 100;

    return {
      ok: true,
      data: {
        placa:              placa || 'FLOTA',
        fechaDesde:         fechaDesde,
        fechaHasta:         fechaHasta,
        disponibilidad:     disponibilidad,    // Método MTBF/(MTBF+MTTR)
        dispAlternativa:    dispAlternativa,   // Método tiempo operativo/total
        disponibilidadUnidad: '%',
        mtbf:               mtbf,
        mttr:               mttr,
        nFallas:            resMTBF.data.nFallas,
        totalHsParada:      resMTBF.data.totalHsParada
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularCostoPorKm(placa, fechaDesde, fechaHasta)
//  Costo/KM = Costo total mantenimientos / KM recorridos en período
// ─────────────────────────────────────────────────────────────
function calcularCostoPorKm(placa, fechaDesde, fechaHasta) {
  try {
    if (!placa) throw new Error('Placa requerida para cálculo de costo/KM');

    var todos       = _filtrarMantenimientos(placa, fechaDesde, fechaHasta);
    var completados = _soloCompletados(todos);

    var costoTotal = completados.reduce(function(s, m) {
      return s + (parseFloat(m.costoTotal) || 0);
    }, 0);

    // Obtener KM del vehículo en el período
    // Se toma el KM máximo y mínimo registrado en los mantenimientos
    var kms = completados
      .map(function(m) { return parseFloat(m.km) || 0; })
      .filter(function(k) { return k > 0; });

    var kmRecorrido  = 0;
    var costoPorKm   = null;

    if (kms.length >= 2) {
      kmRecorrido = Math.max.apply(null, kms) - Math.min.apply(null, kms);
      if (kmRecorrido > 0) {
        costoPorKm = Math.round(costoTotal / kmRecorrido * 100) / 100;
      }
    }

    return {
      ok: true,
      data: {
        placa:       placa,
        fechaDesde:  fechaDesde,
        fechaHasta:  fechaHasta,
        costoTotal:  Math.round(costoTotal * 100) / 100,
        kmRecorrido: kmRecorrido,
        costoPorKm:  costoPorKm,
        unidad:      'S/./KM',
        nMantenimientos: completados.length
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularCostoPorVehiculo(fechaDesde, fechaHasta)
//  Costo total de mantenimiento agrupado por vehículo/placa
// ─────────────────────────────────────────────────────────────
function calcularCostoPorVehiculo(fechaDesde, fechaHasta) {
  try {
    var todos       = _filtrarMantenimientos(null, fechaDesde, fechaHasta);
    var completados = _soloCompletados(todos);

    var porVehiculo = {};
    completados.forEach(function(m) {
      var placa = String(m.placa || 'SIN_PLACA').toUpperCase();
      if (!porVehiculo[placa]) {
        porVehiculo[placa] = { placa: placa, costoTotal: 0, nMant: 0,
                               costoRep: 0, costoMO: 0 };
      }
      porVehiculo[placa].costoTotal += parseFloat(m.costoTotal) || 0;
      porVehiculo[placa].costoRep   += parseFloat(m.costoRep)   || 0;
      porVehiculo[placa].costoMO    += parseFloat(m.costoMO)    || 0;
      porVehiculo[placa].nMant      += 1;
    });

    var resultado = Object.keys(porVehiculo).map(function(k) {
      var v = porVehiculo[k];
      return {
        placa:     v.placa,
        costoTotal:Math.round(v.costoTotal * 100) / 100,
        costoRep:  Math.round(v.costoRep   * 100) / 100,
        costoMO:   Math.round(v.costoMO    * 100) / 100,
        nMant:     v.nMant
      };
    });

    resultado.sort(function(a, b) { return b.costoTotal - a.costoTotal; });

    var totalGeneral = resultado.reduce(function(s, v) { return s + v.costoTotal; }, 0);

    return {
      ok: true,
      data: resultado,
      total: resultado.length,
      resumen: {
        costoTotalFlota: Math.round(totalGeneral * 100) / 100,
        fechaDesde:      fechaDesde,
        fechaHasta:      fechaHasta
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularPorcentajePreventivo(placa, fechaDesde, fechaHasta)
//  % Preventivo = Nº preventivos / Nº total × 100
// ─────────────────────────────────────────────────────────────
function calcularPorcentajePreventivo(placa, fechaDesde, fechaHasta) {
  try {
    var todos       = _filtrarMantenimientos(placa, fechaDesde, fechaHasta);
    var completados = _soloCompletados(todos);

    var total       = completados.length;
    var preventivos = completados.filter(function(m) { return m.tipo === 'PREVENTIVO'; }).length;
    var predictivos = completados.filter(function(m) { return m.tipo === 'PREDICTIVO'; }).length;
    var correctivos = completados.filter(function(m) { return m.tipo === 'CORRECTIVO'; }).length;

    var pctPrev = total > 0 ? Math.round(preventivos / total * 1000) / 10 : 0;
    var pctPred = total > 0 ? Math.round(predictivos / total * 1000) / 10 : 0;
    var pctCorr = total > 0 ? Math.round(correctivos / total * 1000) / 10 : 0;

    return {
      ok: true,
      data: {
        placa:          placa || 'FLOTA',
        fechaDesde:     fechaDesde,
        fechaHasta:     fechaHasta,
        total:          total,
        preventivos:    preventivos,
        predictivos:    predictivos,
        correctivos:    correctivos,
        pctPreventivo:  pctPrev,
        pctPredictivo:  pctPred,
        pctCorrectivo:  pctCorr
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularPorcentajeCorrectivo(placa, fechaDesde, fechaHasta)
//  Alias directo — usa calcularPorcentajePreventivo internamente
// ─────────────────────────────────────────────────────────────
function calcularPorcentajeCorrectivo(placa, fechaDesde, fechaHasta) {
  return calcularPorcentajePreventivo(placa, fechaDesde, fechaHasta);
}

// ─────────────────────────────────────────────────────────────
//  calcularNumeroFallas(placa, fechaDesde, fechaHasta)
//  Conteo de fallas correctivas en el período
// ─────────────────────────────────────────────────────────────
function calcularNumeroFallas(placa, fechaDesde, fechaHasta) {
  try {
    var todos       = _filtrarMantenimientos(placa, fechaDesde, fechaHasta, 'CORRECTIVO');
    var completados = _soloCompletados(todos);

    // Agrupar por mes para tendencia
    var porMes = {};
    completados.forEach(function(m) {
      var fecha = String(m.fechaRegistro || '').substring(0, 7); // yyyy-MM
      if (!fecha) return;
      porMes[fecha] = (porMes[fecha] || 0) + 1;
    });

    var tendencia = Object.keys(porMes).sort().map(function(mes) {
      return { mes: mes, fallas: porMes[mes] };
    });

    return {
      ok: true,
      data: {
        placa:        placa || 'FLOTA',
        fechaDesde:   fechaDesde,
        fechaHasta:   fechaHasta,
        totalFallas:  completados.length,
        tendenciaMes: tendencia
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularHorasParada(placa, fechaDesde, fechaHasta)
//  Suma las horas de duración de todos los correctivos
// ─────────────────────────────────────────────────────────────
function calcularHorasParada(placa, fechaDesde, fechaHasta) {
  try {
    var todos       = _filtrarMantenimientos(placa, fechaDesde, fechaHasta, 'CORRECTIVO');
    var completados = _soloCompletados(todos);

    var totalHoras  = completados.reduce(function(s, m) {
      return s + (parseFloat(m.duracionH) || 0);
    }, 0);

    var porPlaca = {};
    completados.forEach(function(m) {
      var p = String(m.placa || 'SIN_PLACA').toUpperCase();
      if (!porPlaca[p]) porPlaca[p] = 0;
      porPlaca[p] += parseFloat(m.duracionH) || 0;
    });

    var detallePlaca = Object.keys(porPlaca).map(function(p) {
      return { placa: p, horasParada: Math.round(porPlaca[p] * 10) / 10 };
    }).sort(function(a, b) { return b.horasParada - a.horasParada; });

    return {
      ok: true,
      data: {
        placa:            placa || 'FLOTA',
        fechaDesde:       fechaDesde,
        fechaHasta:       fechaHasta,
        totalHorasParada: Math.round(totalHoras * 10) / 10,
        nEventos:         completados.length,
        promHorasPorEvento: completados.length > 0
          ? Math.round(totalHoras / completados.length * 10) / 10 : 0,
        detallePorVehiculo: detallePlaca
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  calcularDashboardKPIs(placa, fechaDesde, fechaHasta)
//  Consolida todos los KPIs en una sola llamada
// ─────────────────────────────────────────────────────────────
function calcularDashboardKPIs(placa, fechaDesde, fechaHasta) {
  try {
    var hoy     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var iniAnio = new Date().getFullYear() + '-01-01';

    fechaDesde = fechaDesde || iniAnio;
    fechaHasta = fechaHasta || hoy;

    var disp   = calcularDisponibilidad(placa, fechaDesde, fechaHasta);
    var pcts   = calcularPorcentajePreventivo(placa, fechaDesde, fechaHasta);
    var fallas = calcularNumeroFallas(placa, fechaDesde, fechaHasta);
    var hsPar  = calcularHorasParada(placa, fechaDesde, fechaHasta);
    var costV  = placa
      ? calcularCostoPorKm(placa, fechaDesde, fechaHasta)
      : calcularCostoPorVehiculo(fechaDesde, fechaHasta);

    return {
      ok: true,
      data: {
        placa:            placa || 'FLOTA',
        fechaDesde:       fechaDesde,
        fechaHasta:       fechaHasta,
        disponibilidad:   disp.ok    ? disp.data.disponibilidad   : null,
        mtbf:             disp.ok    ? disp.data.mtbf              : null,
        mttr:             disp.ok    ? disp.data.mttr              : null,
        pctPreventivo:    pcts.ok    ? pcts.data.pctPreventivo     : null,
        pctCorrectivo:    pcts.ok    ? pcts.data.pctCorrectivo     : null,
        totalFallas:      fallas.ok  ? fallas.data.totalFallas     : null,
        totalHorasParada: hsPar.ok   ? hsPar.data.totalHorasParada : null,
        costoInfo:        costV.ok   ? costV.data                  : null,
        tendenciaFallas:  fallas.ok  ? fallas.data.tendenciaMes    : []
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
