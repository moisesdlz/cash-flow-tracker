/**
 * CASH FLOW OC — Sincronización de base fuente hacia BASE_CASHFLOW
 * ------------------------------------------------------------------
 * Lee la hoja fuente (actualizada semanalmente), filtra las OCs válidas
 * y las sincroniza (upsert por N° de OC) contra la hoja BASE_CASHFLOW,
 * preservando los campos editables ya cargados por el usuario.
 */

/** Versión del código del servidor: el portal la compara con la suya y avisa si pegaste solo uno de los dos archivos. */
const VERSION_CODIGO = '2026-09-30.01';
function obtenerVersionCodigo() { return VERSION_CODIGO; }

const CONFIG = {
  HOJA_FUENTE: 'Base_OC_Rev',     // hoja procesada que alimenta el portal (antes era BASE_FUENTE)
  HOJA_BASE: 'BASE_CASHFLOW',      // hoja de trabajo que alimenta la app
  FILA_INICIO_DATOS: 2,            // fila donde empiezan los datos (1 = encabezados)
  DOMINIOS_VALIDOS: ['P&S', 'Tecnología', 'Seguridad'],
  TC_VALIDO: 'T&C',
  FECHA_CORTE: '26/08/2026',       // actualizar manualmente cada vez que se sube data nueva a BASE_FUENTE
  EMAIL_SOLICITUDES: 'CAMBIAR_POR_CORREO_REAL@bbva.com', // a dónde llega la solicitud de nuevo código

  // Columnas de la hoja fuente (índice numérico fijo, con la letra original en el comentario)
  COL: {
    FECHA_MOD: 2,     // B
    OC: 3,            // C
    PROVEEDOR: 4,     // D
    PROYECTO: 9,      // I
    OFICINA: 11,      // K
    PAGADO: 14,       // N
    SOLPED_IGV: 16,   // P
    PENDIENTE_OC: 17, // Q
    DOMINIO: 20,      // T
    DOMINIO_SP: 22,   // V
    SDATOOL: 23,      // W
    TRIMESTRE: 24,    // X
    AREA: 26,         // Z
    CONTRATO: 30,     // AD
    SOLICITANTE: 31,  // AE
    GB: 32,           // AF
  },
};

// Orden y encabezados de BASE_CASHFLOW: columnas fuente + bloque editable fijo (AG en adelante)
// Índices dentro de BASE_CASHFLOW (1-based) de los campos editables que la app puede escribir
const ENCABEZADOS_BASE = [
  'N° OC', 'Fecha modificación', 'Proveedor', 'Proyecto', 'Oficina',
  'Dominio', 'Área', 'SDATOOL', 'Trimestre', 'N° Contrato', 'Solicitante', 'GB',
  'Monto SOLPED + IGV', 'Monto pendiente / OC', 'Monto Pagado',
  'Proyecto específico', 'Intervención a realizar', 'Estado', 'Fecha de pago propuesta', 'Comentarios',
  'Fecha de pago validada',
  'Valorizado',
  'Fecha pago 1', 'Monto pago 1', 'Fecha pago 2', 'Monto pago 2', 'Fecha pago 3', 'Monto pago 3',
  'Fecha pago 4', 'Monto pago 4', 'Fecha pago 5', 'Monto pago 5',
  'Área corregida', 'Estado de verificación', 'Estado Manual', 'Dominio SP', 'Validación',
];
const COL_BASE = {
  OC: 1, ESTADO: 18, PROYECTO_ESPECIFICO: 16, INTERVENCION: 17, FECHA_PAGO: 19, COMENTARIOS: 20,
  FECHA_PAGO_VALIDADA: 21,
  VALORIZADO: 22,
  FECHA_PAGO_1: 23, MONTO_PAGO_1: 24, FECHA_PAGO_2: 25, MONTO_PAGO_2: 26,
  FECHA_PAGO_3: 27, MONTO_PAGO_3: 28, FECHA_PAGO_4: 29, MONTO_PAGO_4: 30,
  FECHA_PAGO_5: 31, MONTO_PAGO_5: 32,
  AREA_CORREGIDA: 33, ESTADO_VERIFICACION: 34, ESTADO_MANUAL: 35, DOMINIO_SP: 36, ESTADO_VALIDACION: 37,
};

/** Convierte letra de columna (A, B, ..., AF) a índice numérico */
function col(letra) {
  let n = 0;
  for (let i = 0; i < letra.length; i++) {
    n = n * 26 + (letra.charCodeAt(i) - 64);
  }
  return n;
}

/**
 * Punto de entrada: sincroniza la fuente hacia BASE_CASHFLOW.
 * Ejecutar manualmente o programar un trigger semanal (por fecha, después de subir la fuente).
 */
function sincronizarBaseCashFlow() {
  const prep = actualizarTodoAntesDeSincronizar(); // (filtrar_base.gs) refresca BASE_FILTRADA y espera el recálculo de Base_OC_Rev antes de leerla
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFuente = ss.getSheetByName(CONFIG.HOJA_FUENTE);
  let hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase) {
    hojaBase = ss.insertSheet(CONFIG.HOJA_BASE);
  }

  const datosFuente = hojaFuente
    .getRange(CONFIG.FILA_INICIO_DATOS, 1, hojaFuente.getLastRow() - 1, hojaFuente.getLastColumn())
    .getValues();

  // Índice de filas existentes en BASE_CASHFLOW por N° de OC, para preservar campos editables
  const filasExistentes = leerBaseExistente(hojaBase);

  const filasFinales = [ENCABEZADOS_BASE];

  datosFuente.forEach(fila => {
    const oc = fila[CONFIG.COL.OC - 1];
    const dominioSp = fila[CONFIG.COL.DOMINIO_SP - 1];
    const dominio = fila[CONFIG.COL.DOMINIO - 1];

    const cumpleFiltros =
      oc !== '' && oc !== null &&
      CONFIG.DOMINIOS_VALIDOS.includes(dominio);

    if (!cumpleFiltros) return;

    const solped = Number(fila[CONFIG.COL.SOLPED_IGV - 1]) || 0;
    const pendienteOC = Number(fila[CONFIG.COL.PENDIENTE_OC - 1]) || 0;
    const pagado = Number(fila[CONFIG.COL.PAGADO - 1]) || 0;
    const estado = calcularEstado(solped, pendienteOC, pagado);

    const existente = filasExistentes[oc] || {};

    filasFinales.push([
      oc,
      fila[CONFIG.COL.FECHA_MOD - 1],
      fila[CONFIG.COL.PROVEEDOR - 1],
      fila[CONFIG.COL.PROYECTO - 1],
      fila[CONFIG.COL.OFICINA - 1],
      dominio,
      fila[CONFIG.COL.AREA - 1],
      fila[CONFIG.COL.SDATOOL - 1],
      fila[CONFIG.COL.TRIMESTRE - 1],
      fila[CONFIG.COL.CONTRATO - 1],
      fila[CONFIG.COL.SOLICITANTE - 1],
      fila[CONFIG.COL.GB - 1],
      solped,
      pendienteOC,
      pagado,
      existente.proyectoEspecifico || '',   // se conserva si ya existía
      existente.intervencion || '',
      estado,
      existente.fechaPago || '',
      existente.comentarios || '',
      existente.fechaPagoValidada || '',
      existente.valorizado || false,
      existente.fechaPago1 || '', existente.montoPago1 || '',
      existente.fechaPago2 || '', existente.montoPago2 || '',
      existente.fechaPago3 || '', existente.montoPago3 || '',
      existente.fechaPago4 || '', existente.montoPago4 || '',
      existente.fechaPago5 || '', existente.montoPago5 || '',
      existente.areaCorregida || '',
      existente.estadoVerificacion || '',
      existente.estadoManual || '',
      dominioSp,
      existente.estadoValidacion || '',
    ]);
  });

  // Guardia: si se leyeron muchas menos filas de las esperadas, Base_OC_Rev estaba a medias: NO se toca BASE_CASHFLOW
  // (borrarla ahora haría perder las fechas y demás datos cargados de las OC que faltan).
  if (prep && prep.pasan && (filasFinales.length - 1) < prep.pasan * 0.9) {
    throw new Error('La sincronización solo leyó ' + (filasFinales.length - 1) + ' filas de ~' + prep.pasan + ' esperadas (Base_OC_Rev aún se estaba recalculando). No se modificó BASE_CASHFLOW. Vuelve a intentarlo en 1–2 minutos.');
  }

  hojaBase.clearContents();
  hojaBase.getRange(1, 1, filasFinales.length, ENCABEZADOS_BASE.length).setValues(filasFinales);
  PropertiesService.getScriptProperties().setProperty('ULTIMA_SYNC', new Date().toISOString());
  bumpCacheVer_();

  // Detectar OCs nuevas (que no estaban en BASE_CASHFLOW antes de esta sync)
  const ocsNuevas = filasFinales.slice(1).filter(f => !filasExistentes[f[0]]);

  if (ocsNuevas.length > 0) {
    // Registrar en la pestaña de historial
    const HOJA_HISTORIAL = 'HistorialSincronizacion';
    let hojaHist = ss.getSheetByName(HOJA_HISTORIAL);
    if (!hojaHist) {
      hojaHist = ss.insertSheet(HOJA_HISTORIAL);
      hojaHist.getRange(1,1,1,6).setValues([['Fecha Sync','N° OC','Proveedor','Proyecto','Dominio','Monto OC']]);
      hojaHist.setFrozenRows(1);
    }
    const fechaSync = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const filaHist = ocsNuevas.map(f => [fechaSync, f[0], f[2], f[3], f[5], f[13]]);
    hojaHist.getRange(hojaHist.getLastRow()+1, 1, filaHist.length, 6).setValues(filaHist);

    // Enviar correo a maestros
    const correosMaestros = obtenerCorreosMaestros_();
    if (correosMaestros.length) {
      const tabla = ocsNuevas.map((f,i) =>
        (i+1) + '. OC ' + f[0] + ' — ' + (f[2]||'') + ' | ' + (f[3]||'') + ' | S/ ' + (Number(f[13])||0).toLocaleString('es-PE')
      ).join('\n');
      MailApp.sendEmail({
        to: correosMaestros.join(','),
        subject: 'Control Presupuestal — Sincronización ' + fechaSync + ' (' + ocsNuevas.length + ' OC nuevas)',
        body: 'Se han agregado ' + ocsNuevas.length + ' orden(es) de compra nuevas en la sincronización del ' + fechaSync + ':\n\n' + tabla + '\n\nRevísalas en el portal.',
      });
    }
  }

  return { ok: true, total: filasFinales.length - 1, nuevas: ocsNuevas.length };
}

/* ====================================================================
   SINCRONIZACIÓN CON VISTA PREVIA (misma mecánica que "Importar fechas
   de Hans"): la fuente ahora se alimenta por IMPORTRANGE y cambia sola,
   así que antes de pisar BASE_CASHFLOW se calcula el resultado SIN
   escribir nada, se le muestra al Maestro qué OC son nuevas, cuáles
   cambiaron de monto/estado y cuáles desaparecerían, y solo si confirma
   se aplica de verdad. Al aplicar, se manda un correo resumen a los
   Líderes (y Maestros).
   ==================================================================== */

/**
 * Calcula (SIN escribir nada) el resultado de sincronizar ahora mismo,
 * comparado contra el BASE_CASHFLOW actual: arma las mismas filasFinales
 * que sincronizarBaseCashFlow(), y además detecta OC nuevas, OC con
 * cambio de monto o estado, y OC que desaparecerían (ya no están en la
 * fuente o salieron de los dominios válidos). Se usa dos veces con los
 * mismos datos: una para la vista previa, otra para ejecutar de verdad.
 */
function calcularSincronizacion_() {
  const prep = actualizarTodoAntesDeSincronizar();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFuente = ss.getSheetByName(CONFIG.HOJA_FUENTE);
  let hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase) hojaBase = ss.insertSheet(CONFIG.HOJA_BASE);

  const datosFuente = hojaFuente
    .getRange(CONFIG.FILA_INICIO_DATOS, 1, hojaFuente.getLastRow() - 1, hojaFuente.getLastColumn())
    .getValues();

  const filasExistentes = leerBaseExistente(hojaBase);   // campos editables a preservar, por OC
  const ultimaFilaBase = hojaBase.getLastRow();
  const datosBaseActual = ultimaFilaBase >= 2
    ? hojaBase.getRange(2, 1, ultimaFilaBase - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues()
    : [];
  const agrupadoViejoMap = {};
  agruparPorOC_(datosBaseActual).forEach(g => { agrupadoViejoMap[g.oc] = g; });   // estado ya considera Estado Manual

  const filasFinales = [ENCABEZADOS_BASE];
  const nuevoAgregado = {};   // OC -> montos sumados de esta pasada (una OC puede tener varias líneas en la fuente)

  datosFuente.forEach(fila => {
    const oc = fila[CONFIG.COL.OC - 1];
    const dominioSp = fila[CONFIG.COL.DOMINIO_SP - 1];
    const dominio = fila[CONFIG.COL.DOMINIO - 1];

    const cumpleFiltros =
      oc !== '' && oc !== null &&
      CONFIG.DOMINIOS_VALIDOS.includes(dominio);

    if (!cumpleFiltros) return;

    const solped = Number(fila[CONFIG.COL.SOLPED_IGV - 1]) || 0;
    const pendienteOC = Number(fila[CONFIG.COL.PENDIENTE_OC - 1]) || 0;
    const pagado = Number(fila[CONFIG.COL.PAGADO - 1]) || 0;
    const estado = calcularEstado(solped, pendienteOC, pagado);

    const existente = filasExistentes[oc] || {};

    filasFinales.push([
      oc,
      fila[CONFIG.COL.FECHA_MOD - 1],
      fila[CONFIG.COL.PROVEEDOR - 1],
      fila[CONFIG.COL.PROYECTO - 1],
      fila[CONFIG.COL.OFICINA - 1],
      dominio,
      fila[CONFIG.COL.AREA - 1],
      fila[CONFIG.COL.SDATOOL - 1],
      fila[CONFIG.COL.TRIMESTRE - 1],
      fila[CONFIG.COL.CONTRATO - 1],
      fila[CONFIG.COL.SOLICITANTE - 1],
      fila[CONFIG.COL.GB - 1],
      solped,
      pendienteOC,
      pagado,
      existente.proyectoEspecifico || '',
      existente.intervencion || '',
      estado,
      existente.fechaPago || '',
      existente.comentarios || '',
      existente.fechaPagoValidada || '',
      existente.valorizado || false,
      existente.fechaPago1 || '', existente.montoPago1 || '',
      existente.fechaPago2 || '', existente.montoPago2 || '',
      existente.fechaPago3 || '', existente.montoPago3 || '',
      existente.fechaPago4 || '', existente.montoPago4 || '',
      existente.fechaPago5 || '', existente.montoPago5 || '',
      existente.areaCorregida || '',
      existente.estadoVerificacion || '',
      existente.estadoManual || '',
      dominioSp,
      existente.estadoValidacion || '',
    ]);

    const acc = nuevoAgregado[oc] || (nuevoAgregado[oc] = {
      solped: 0, pendienteOC: 0, pagado: 0,
      proyecto: fila[CONFIG.COL.PROYECTO - 1], area: fila[CONFIG.COL.AREA - 1], dominio: dominio,
    });
    acc.solped += solped;
    acc.pendienteOC += pendienteOC;
    acc.pagado += pagado;
  });

  const nuevas = [], cambios = [];
  Object.keys(nuevoAgregado).forEach(oc => {
    const n = nuevoAgregado[oc];
    const estadoNuevo = calcularEstado(n.solped, n.pendienteOC, n.pagado);
    const viejo = agrupadoViejoMap[oc];
    const montoNuevo = n.pendienteOC || n.solped || 0;
    if (!viejo) {
      nuevas.push({ oc: oc, proyecto: n.proyecto, area: n.area, dominio: n.dominio, estado: estadoNuevo, monto: montoNuevo });
      return;
    }
    const montoViejo = viejo.pendienteOC || viejo.solped || 0;
    const cambioMonto = Math.abs(montoViejo - montoNuevo) > 0.5;
    const cambioEstado = viejo.estado !== estadoNuevo;
    if (cambioMonto || cambioEstado) {
      cambios.push({
        oc: oc, proyecto: n.proyecto, area: n.area,
        estadoAnterior: viejo.estado, estadoNuevo: estadoNuevo,
        montoAnterior: montoViejo, montoNuevo: montoNuevo,
        cambioEstado: cambioEstado, cambioMonto: cambioMonto,
      });
    }
  });

  const desaparecidas = Object.keys(agrupadoViejoMap)
    .filter(oc => !nuevoAgregado[oc])
    .map(oc => {
      const v = agrupadoViejoMap[oc];
      return { oc: oc, proyecto: v.proyecto, area: v.area, estado: v.estado, monto: v.pendienteOC || v.solped || 0 };
    });

  const totalOCViejo = Object.keys(agrupadoViejoMap).length;
  const totalOCNuevo = Object.keys(nuevoAgregado).length;

  return {
    prep: prep, filasFinales: filasFinales,
    nuevas: nuevas, cambios: cambios, desaparecidas: desaparecidas,
    sinCambio: totalOCNuevo - nuevas.length - cambios.length,
    totalOCViejo: totalOCViejo, totalOCNuevo: totalOCNuevo,
  };
}

/** Paso 1 (no escribe nada): calcula los cambios para mostrar el aviso "¿Proceder?" antes de tocar BASE_CASHFLOW. */
function previsualizarSincronizacionOC() {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede correr esta sincronización.' };
  const r = calcularSincronizacion_();
  const limitar = arr => arr.slice(0, 150);   // el navegador no necesita miles de filas para la vista previa
  return {
    ok: true,
    totalFilas: r.filasFinales.length - 1, totalOCViejo: r.totalOCViejo, totalOCNuevo: r.totalOCNuevo,
    cantNuevas: r.nuevas.length, cantCambios: r.cambios.length, cantDesaparecidas: r.desaparecidas.length, sinCambio: r.sinCambio,
    nuevas: limitar(r.nuevas), cambios: limitar(r.cambios), desaparecidas: limitar(r.desaparecidas),
  };
}

/**
 * Paso 2 (ya confirmado por el Maestro): vuelve a calcular (por si la fuente
 * cambió entre la vista previa y el clic en "Proceder"), aplica la misma
 * guarda de siempre, escribe BASE_CASHFLOW y manda el correo resumen a
 * Líderes y Maestros.
 */
function ejecutarSincronizacionConfirmada() {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede correr esta sincronización.' };
  const r = calcularSincronizacion_();

  if (r.prep && r.prep.pasan && r.filasFinales.length - 1 < r.prep.pasan * 0.9) {
    return { ok: false, error: 'La sincronización solo leyó ' + (r.filasFinales.length - 1) + ' filas de ~' + r.prep.pasan + ' esperadas (Base_OC_Rev aún se estaba recalculando). No se modificó BASE_CASHFLOW. Vuelve a intentarlo en 1–2 minutos.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
  hojaBase.clearContents();
  hojaBase.getRange(1, 1, r.filasFinales.length, ENCABEZADOS_BASE.length).setValues(r.filasFinales);
  PropertiesService.getScriptProperties().setProperty('ULTIMA_SYNC', new Date().toISOString());
  bumpCacheVer_();

  if (r.nuevas.length > 0) {
    const HOJA_HISTORIAL = 'HistorialSincronizacion';
    let hojaHist = ss.getSheetByName(HOJA_HISTORIAL);
    if (!hojaHist) {
      hojaHist = ss.insertSheet(HOJA_HISTORIAL);
      hojaHist.getRange(1, 1, 1, 6).setValues([['Fecha Sync', 'N° OC', 'Proyecto', 'Área', 'Dominio', 'Monto OC']]);
      hojaHist.setFrozenRows(1);
    }
    const fechaSync = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const filaHist = r.nuevas.map(n => [fechaSync, n.oc, n.proyecto || '', n.area || '', n.dominio || '', n.monto || 0]);
    hojaHist.getRange(hojaHist.getLastRow() + 1, 1, filaHist.length, 6).setValues(filaHist);
  }

  try { enviarResumenSincronizacionLideres_(r); } catch (eCorreo) { Logger.log('No se pudo enviar el correo de resumen de sincronización: ' + eCorreo); }
  try { registrarHistorialPagados_(r); } catch (eHistPag) { Logger.log('No se pudo registrar el historial de OC pagadas: ' + eHistPag); }

  return {
    ok: true, nuevas: r.nuevas.length, cambios: r.cambios.length, desaparecidas: r.desaparecidas.length,
    sinCambio: r.sinCambio, total: r.totalOCNuevo,
  };
}

/**
 * Registra en "HistorialEstadoPagado" el momento REAL en que una OC pasó a Estado = "Pagado"
 * (detectado por esta misma sincronización, comparando el estado antes/después). Esto reemplaza
 * agrupar por la columna "Fecha modificación" de la fuente, que se toca en bloque para miles de
 * filas sin que signifique que de verdad cambiaron — por eso el conteo por semana salía inflado
 * y sin sentido. Con este historial, "semana en que pasó a Pagado" es un dato real y verificable.
 */
function registrarHistorialPagados_(r) {
  const pagadasAhora = []
    .concat((r.nuevas || []).filter(n => n.estado === 'Pagado').map(n => ({ oc: n.oc, proyecto: n.proyecto, monto: n.monto })))
    .concat((r.cambios || []).filter(c => c.cambioEstado && c.estadoNuevo === 'Pagado').map(c => ({ oc: c.oc, proyecto: c.proyecto, monto: c.montoNuevo })));
  if (!pagadasAhora.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('HistorialEstadoPagado');
  if (!hoja) {
    hoja = ss.insertSheet('HistorialEstadoPagado');
    hoja.getRange(1, 1, 1, 4).setValues([['Fecha Sync', 'N° OC', 'Proyecto', 'Monto']]);
    hoja.setFrozenRows(1);
  }
  const fechaSync = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const filas = pagadasAhora.map(p => [fechaSync, p.oc, p.proyecto || '', p.monto || 0]);
  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 4).setValues(filas);
}

/**
 * Para el modal "Ver y validar todo lo Pagado": devuelve las OC Pagadas separadas en dos grupos:
 *  - conHistorial: se sabe EXACTAMENTE en qué sincronización pasaron a Pagado (fechaMod = esa fecha real).
 *  - sinHistorial: ya estaban Pagadas desde antes de que existiera este registro — se muestran aparte,
 *    sin inventarles una fecha, para no repetir el problema de antes.
 * Cacheado unos minutos, igual que el resto de reportes pesados.
 */
function obtenerHistorialPagosPorSemana() {
  return conCache_('histpagos', {}, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
    const infoPagadas = {};   // oc -> {proyecto, monto, estadoValidacion, fechaModAprox}
    if (hojaBase && hojaBase.getLastRow() >= 2) {
      const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
      agruparPorOC_(datos).filter(g => g.estado === 'Pagado').forEach(g => {
        infoPagadas[g.oc] = {
          proyecto: g.proyecto || '',
          monto: (Number(g.pendienteOC) || 0) + (Number(g.pagado) || 0),
          estadoValidacion: g.estadoValidacion || '',
          fechaModAprox: formatearFecha(g.fechaMod),
        };
      });
    }

    const hojaHist = ss.getSheetByName('HistorialEstadoPagado');
    const conHistorial = [];
    const vistos = {};
    if (hojaHist && hojaHist.getLastRow() >= 2) {
      hojaHist.getRange(2, 1, hojaHist.getLastRow() - 1, 4).getValues().forEach(f => {
        const oc = String(f[1]);
        if (!infoPagadas[oc] || vistos[oc]) return;   // ya no sigue Pagada hoy, o ya se registró (se queda con la primera vez que pasó a Pagado)
        vistos[oc] = true;
        conHistorial.push({
          oc: oc, proyecto: infoPagadas[oc].proyecto, monto: infoPagadas[oc].monto,
          estadoValidacion: infoPagadas[oc].estadoValidacion,
          fechaMod: Object.prototype.toString.call(f[0]) === '[object Date]' ? formatearFecha(f[0]) : String(f[0]).slice(0, 10),
        });
      });
    }
    // Backlog: ya estaban "Pagado" antes de que existiera este registro. No sabemos la fecha EXACTA
    // en que pasaron a Pagado, pero sí llevan la Fecha de modificación de la fuente — se manda igual,
    // marcada como aproximada, solo para poder partirlas por semana en vez de mostrarlas todas juntas.
    const sinHistorial = Object.keys(infoPagadas)
      .filter(oc => !vistos[oc])
      .map(oc => ({
        oc: oc, proyecto: infoPagadas[oc].proyecto, monto: infoPagadas[oc].monto,
        estadoValidacion: infoPagadas[oc].estadoValidacion, fechaMod: infoPagadas[oc].fechaModAprox || '',
      }));

    return { conHistorial: conHistorial, sinHistorial: sinHistorial };
  });
}

/** Correo a Líderes y Maestros con el detalle de la sincronización: OC nuevas, con cambios y desaparecidas. */
function enviarResumenSincronizacionLideres_(r) {
  const destinos = obtenerCorreosLideresYMaestros_();
  if (!destinos.length) return;
  const tz = Session.getScriptTimeZone();
  const fechaSync = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  const monto = v => 'S/ ' + (Number(v) || 0).toLocaleString('es-PE');
  const filaNueva = n => `<tr><td style="padding:5px 10px;">${n.oc}</td><td style="padding:5px 10px;">${escHtmlSrv_(n.proyecto)}</td><td style="padding:5px 10px;">${escHtmlSrv_(n.area)}</td><td style="padding:5px 10px;">${n.estado}</td><td style="padding:5px 10px;text-align:right;">${monto(n.monto)}</td></tr>`;
  const filaCambio = c => `<tr><td style="padding:5px 10px;">${c.oc}</td><td style="padding:5px 10px;">${escHtmlSrv_(c.proyecto)}</td><td style="padding:5px 10px;">${c.estadoAnterior} → ${c.estadoNuevo}</td><td style="padding:5px 10px;text-align:right;">${monto(c.montoAnterior)} → ${monto(c.montoNuevo)}</td></tr>`;
  const filaDesap = d => `<tr><td style="padding:5px 10px;">${d.oc}</td><td style="padding:5px 10px;">${escHtmlSrv_(d.proyecto)}</td><td style="padding:5px 10px;">${d.estado}</td><td style="padding:5px 10px;text-align:right;">${monto(d.monto)}</td></tr>`;
  const cuerpo = `
    <div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;">
      <div style="background:#0B1D33;color:#fff;padding:18px 22px;border-radius:8px 8px 0 0;">
        <div style="font-size:12px;opacity:.75;">Portal CAPEX P&S · Detalle OC</div>
        <h2 style="margin:6px 0 0;font-size:17px;">⟳ Sincronización aplicada — ${fechaSync}</h2>
      </div>
      <div style="background:#fff;padding:20px 22px;border:1px solid #e0e0e0;border-top:4px solid #2E6FBE;">
        ${r.nuevas.length ? `<h3 style="color:#4C9270;font-size:13px;">🆕 OC nuevas (${r.nuevas.length})</h3>
          <table style="border-collapse:collapse;font-size:13px;width:100%;margin-bottom:16px;"><tr style="color:#6B7A99;font-size:11px;text-align:left;"><td style="padding:5px 10px;">N° OC</td><td style="padding:5px 10px;">Proyecto</td><td style="padding:5px 10px;">Área</td><td style="padding:5px 10px;">Estado</td><td style="padding:5px 10px;text-align:right;">Monto</td></tr>${r.nuevas.map(filaNueva).join('')}</table>` : ''}
        ${r.cambios.length ? `<h3 style="color:#B98A3D;font-size:13px;">🔁 OC con cambio de monto o estado (${r.cambios.length})</h3>
          <table style="border-collapse:collapse;font-size:13px;width:100%;margin-bottom:16px;"><tr style="color:#6B7A99;font-size:11px;text-align:left;"><td style="padding:5px 10px;">N° OC</td><td style="padding:5px 10px;">Proyecto</td><td style="padding:5px 10px;">Estado</td><td style="padding:5px 10px;text-align:right;">Monto pendiente</td></tr>${r.cambios.map(filaCambio).join('')}</table>` : ''}
        ${r.desaparecidas.length ? `<h3 style="color:#C0392B;font-size:13px;">⚠️ OC que ya no aparecen en la fuente (${r.desaparecidas.length})</h3>
          <table style="border-collapse:collapse;font-size:13px;width:100%;"><tr style="color:#6B7A99;font-size:11px;text-align:left;"><td style="padding:5px 10px;">N° OC</td><td style="padding:5px 10px;">Proyecto</td><td style="padding:5px 10px;">Último estado</td><td style="padding:5px 10px;text-align:right;">Monto pendiente</td></tr>${r.desaparecidas.map(filaDesap).join('')}</table>
          <p style="font-size:12px;color:#6B7A99;margin-top:10px;">Revisa si corresponde (OC cerrada/anulada en la fuente) o si es un error de la fuente antes de asumir que ya no existen.</p>` : ''}
        ${!r.nuevas.length && !r.cambios.length && !r.desaparecidas.length ? '<p style="font-size:13px;color:#4C9270;">No hubo cambios esta vez.</p>' : ''}
      </div>
      <div style="background:#f0f4f8;padding:10px 22px;border-radius:0 0 8px 8px;font-size:11px;color:#6B7A99;">
        ${r.sinCambio} OC sin cambio de ${r.totalOCNuevo} OC totales en la fuente.
      </div>
    </div>`;
  MailApp.sendEmail({
    to: destinos.join(','),
    subject: `[Portal CAPEX] Sincronización OC — ${r.nuevas.length} nueva(s), ${r.cambios.length} con cambio, ${r.desaparecidas.length} ya no en la fuente`,
    htmlBody: cuerpo,
  });
}

/** Lee BASE_CASHFLOW actual y arma un mapa {OC: {campos editables}} para preservarlos en la sync */
function leerBaseExistente(hojaBase) {
  const mapa = {};
  const ultimaFila = hojaBase.getLastRow();
  if (ultimaFila < 2) return mapa;

  const datos = hojaBase.getRange(2, 1, ultimaFila - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  datos.forEach(fila => {
    const oc = fila[COL_BASE.OC - 1];
    if (!oc) return;
    mapa[oc] = {
      proyectoEspecifico: fila[COL_BASE.PROYECTO_ESPECIFICO - 1],
      intervencion: fila[COL_BASE.INTERVENCION - 1],
      fechaPago: fila[COL_BASE.FECHA_PAGO - 1],
      comentarios: fila[COL_BASE.COMENTARIOS - 1],
      fechaPagoValidada: fila[COL_BASE.FECHA_PAGO_VALIDADA - 1],
      valorizado: fila[COL_BASE.VALORIZADO - 1],
      fechaPago1: fila[COL_BASE.FECHA_PAGO_1 - 1], montoPago1: fila[COL_BASE.MONTO_PAGO_1 - 1],
      fechaPago2: fila[COL_BASE.FECHA_PAGO_2 - 1], montoPago2: fila[COL_BASE.MONTO_PAGO_2 - 1],
      fechaPago3: fila[COL_BASE.FECHA_PAGO_3 - 1], montoPago3: fila[COL_BASE.MONTO_PAGO_3 - 1],
      fechaPago4: fila[COL_BASE.FECHA_PAGO_4 - 1], montoPago4: fila[COL_BASE.MONTO_PAGO_4 - 1],
      fechaPago5: fila[COL_BASE.FECHA_PAGO_5 - 1], montoPago5: fila[COL_BASE.MONTO_PAGO_5 - 1],
      areaCorregida: fila[COL_BASE.AREA_CORREGIDA - 1],
      estadoManual: fila[COL_BASE.ESTADO_MANUAL - 1],
      estadoVerificacion: fila[COL_BASE.ESTADO_VERIFICACION - 1],
      estadoValidacion: fila[COL_BASE.ESTADO_VALIDACION - 1] || '',
    };
  });
  return mapa;
}

/**
 * Estado de la OC según qué columna tiene monto registrado.
 * Prioridad: Por comprometer > Por pagar / Pagado parcialmente > Pagado.
 * - Pagado solo cuando TODAS las posiciones de la OC ya están en la columna Pagado
 *   (nada queda en SOLPED ni en Pendiente de pago/OC).
 * - Pagado parcialmente cuando coexisten monto pendiente de OC y monto ya pagado.
 */
function calcularEstado(solped, pendienteOC, pagado) {
  if (solped > 0) return 'Por comprometer';
  if (pendienteOC > 0 && pagado > 0) return 'Pagado parcialmente';
  if (pendienteOC > 0) return 'Por pagar';
  if (pagado > 0) return 'Pagado';
  return 'Sin registrar';
}

/**
 * Revisa cuántas OCs en BASE_CASHFLOW tienen Fecha de pago propuesta / validada
 * cargada en este momento, y cuándo fue la última vez que corrió la sincronización.
 * Pensado para llamarse desde el botón de Diagnóstico en la app.
 */
function diagnosticoFechas() {
  const partes = [];
  const ultimaSync = PropertiesService.getScriptProperties().getProperty('ULTIMA_SYNC');
  partes.push('Última sincronización: ' + (ultimaSync ? new Date(ultimaSync).toLocaleString('es-PE') : 'no registrada (corre sincronizarBaseCashFlow al menos una vez más)'));

  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) {
    partes.push('BASE_CASHFLOW está vacía.');
    return partes.join('\n');
  }

  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  const conPropuesta = agrupadas.filter(g => !!g.fechaPago);
  const conValidada = agrupadas.filter(g => !!g.fechaPagoValidada);

  partes.push('Total de OCs (agrupadas): ' + agrupadas.length);
  partes.push('Con Fecha de pago propuesta cargada: ' + conPropuesta.length);
  partes.push('Con Fecha de pago validada cargada: ' + conValidada.length);

  if (conPropuesta.length > 0) {
    partes.push('\nEjemplos (hasta 5):');
    conPropuesta.slice(0, 5).forEach(g => {
      partes.push('  OC ' + g.oc + ' — propuesta: ' + formatearFecha(g.fechaPago) + ' | validada: ' + (formatearFecha(g.fechaPagoValidada) || '(sin validar)'));
    });
  }

  return partes.join('\n');
}

/**
 * Guarda la valorización de una OC: si está marcada como "Valorizado" y hasta
 * 5 cuotas (fecha + monto). cuotas es un array de hasta 5 objetos {fecha, monto}.
 * Escribe en todas las líneas que comparten esa OC, igual que actualizarCampoOC.
 */
function guardarValorizacion(numeroOC, valorizado, cuotas) {
  cuotas = cuotas || [];
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  const ultimaFila = hojaBase.getLastRow();
  const columnaOC = hojaBase.getRange(2, COL_BASE.OC, ultimaFila - 1, 1).getValues();

  const colsFecha = [COL_BASE.FECHA_PAGO_1, COL_BASE.FECHA_PAGO_2, COL_BASE.FECHA_PAGO_3, COL_BASE.FECHA_PAGO_4, COL_BASE.FECHA_PAGO_5];
  const colsMonto = [COL_BASE.MONTO_PAGO_1, COL_BASE.MONTO_PAGO_2, COL_BASE.MONTO_PAGO_3, COL_BASE.MONTO_PAGO_4, COL_BASE.MONTO_PAGO_5];
  const primeraCol = COL_BASE.VALORIZADO;
  const ultimaCol = COL_BASE.MONTO_PAGO_5;

  const filaValores = new Array(ultimaCol - primeraCol + 1).fill('');
  filaValores[COL_BASE.VALORIZADO - primeraCol] = !!valorizado;
  for (let i = 0; i < 5; i++) {
    filaValores[colsFecha[i] - primeraCol] = (cuotas[i] && cuotas[i].fecha) || '';
    filaValores[colsMonto[i] - primeraCol] = (cuotas[i] && cuotas[i].monto) ? Number(cuotas[i].monto) : '';
  }

  let filasActualizadas = 0;
  for (let i = 0; i < columnaOC.length; i++) {
    if (String(columnaOC[i][0]) !== String(numeroOC)) continue;
    const filaReal = i + 2;
    hojaBase.getRange(filaReal, primeraCol, 1, filaValores.length).setValues([filaValores]);
    filasActualizadas++;
  }

  if (filasActualizadas === 0) return { ok: false, error: 'OC no encontrada en BASE_CASHFLOW' };
  bumpCacheVer_();
  return { ok: true, filasActualizadas: filasActualizadas };
}

/**
 * DIAGNÓSTICO — ejecuta esta función manualmente desde el editor (▶) y
 * revisa Ver → Registros (o Ctrl+Enter) para ver exactamente qué detecta el script.
 */
function diagnosticoCashFlow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Spreadsheet activo: ' + (ss ? ss.getName() + ' (' + ss.getId() + ')' : 'NULL — el script no está encontrando el Sheet'));
  if (!ss) return;

  const nombres = ss.getSheets().map(s => "'" + s.getName() + "'");
  Logger.log('Pestañas encontradas: ' + nombres.join(', '));

  const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
  Logger.log('Buscando pestaña llamada exactamente: "' + CONFIG.HOJA_BASE + '"');
  Logger.log('¿La encontró?: ' + (hojaBase ? 'SÍ' : 'NO'));

  if (hojaBase) {
    Logger.log('Última fila con datos: ' + hojaBase.getLastRow());
    Logger.log('Última columna con datos: ' + hojaBase.getLastColumn());
    if (hojaBase.getLastRow() >= 2) {
      Logger.log('Primera fila de datos (fila 2): ' + JSON.stringify(hojaBase.getRange(2, 1, 1, hojaBase.getLastColumn()).getValues()));
    }
  }
}

/**
 * Igual que diagnosticoCashFlow, pero pensado para llamarse desde el HTML
 * vía google.script.run, para ver qué ve el contexto real de la app web.
 */
function diagnosticoWeb() {
  const partes = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  partes.push('Spreadsheet: ' + (ss ? ss.getName() + ' | ID: ' + ss.getId() : 'NULL'));
  if (!ss) return partes.join('\n');

  partes.push('Pestañas: ' + ss.getSheets().map(s => s.getName()).join(', '));
  const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
  partes.push('¿Encontró "' + CONFIG.HOJA_BASE + '"?: ' + (hojaBase ? 'SÍ' : 'NO'));
  if (hojaBase) {
    partes.push('Última fila: ' + hojaBase.getLastRow() + ' | Última columna: ' + hojaBase.getLastColumn());
  }
  const resultado = obtenerDatosCashFlow({}, 0, 5);
  partes.push('obtenerDatosCashFlow() → total filas que cumplen filtro: ' + resultado.total);
  if (resultado.filas.length > 0) {
    partes.push('Primera fila: ' + JSON.stringify(resultado.filas[0]));
  }
  return partes.join('\n');
}

/**
 * Data para el Gantt de pagos: agrupa el monto pendiente de OC por SDATOOL y por
 * semana (lunes a domingo), usando la Fecha de pago cargada en cada OC.
 * Solo entran OCs con Fecha de pago cargada y saldo pendiente > 0 (lo que falta pagar).
 */
function calcularGantt_(filtros) {
  filtros = filtros || {};
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return { semanas: [], filas: [] };

  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  const filtradas = agrupadas.filter(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return false;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return false;
    if (!coincideFiltro_(filtros.estado, g.estado)) return false;
    if (!coincideFiltro_(filtros.area, g.area)) return false;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return false;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return false;
    if (!coincideFiltro_(filtros.dominioSp, (g.dominioSp||'').trim() || '(Sin dato)')) return false;
    return true;
  });

  // Por cada OC arma sus "puntos" (fecha, monto): si está valorizada usa hasta
  // 5 cuotas propias; si no, cae al comportamiento anterior (una sola fecha propuesta).
  const vcMap = leerValidacionCuotas_(); // validación por cuota (hoja ValidacionCuotas)
  const puntos = [];
  filtradas.forEach(g => {
    const clave = (g.sdatool || 'Sin SDATOOL') + '|' + (g.proyecto || 'Sin proyecto');
    const validado = (g.estadoValidacion || '').toLowerCase() === 'verificado';

    if (g.valorizado) {
      // OC valorizada en cuotas: cada cuota siempre aparece en la proyección
      const cuotas = [
        [g.fechaPago1, g.montoPago1], [g.fechaPago2, g.montoPago2], [g.fechaPago3, g.montoPago3],
        [g.fechaPago4, g.montoPago4], [g.fechaPago5, g.montoPago5],
      ];
      cuotas.forEach(([fecha, monto], idx) => {
        if (fecha && Number(monto) > 0) {
          // cada cuota se valida por separado (pago parcial = solo la cuota verificada suma a la línea verde)
          const vc = vcMap[String(g.oc) + '|' + (idx + 1)] || {};
          const validadoCuota = String(vc.estado || '').toLowerCase() === 'verificado';
          puntos.push({ clave, fecha, monto: Number(monto), validado: validadoCuota, fechaValidada: vc.fechaValidada || null, dominio: g.dominio||'', dominioSp: g.dominioSp||'' });
        }
      });
    } else if (g.fechaPago) {
      // OC simple: siempre entra a la proyección con el monto total de la OC
      // (pendienteOC + pagado = monto OC total). Si pendienteOC ya es 0 porque
      // fue pagada, usamos pagado para mantener el monto en la línea azul.
      // La línea verde (validado) usa el mismo monto.
      const montoTotal = g.pendienteOC > 0 ? g.pendienteOC : g.pagado;
      if (montoTotal > 0) {
        puntos.push({ clave, fecha: g.fechaPago, monto: montoTotal, validado, fechaValidada: g.fechaPagoValidada || null, dominio: g.dominio||'', dominioSp: g.dominioSp||'' });
      }
    }
  });

  if (puntos.length === 0) return { semanas: [], filas: [] };

  // Rango del gráfico: usa filtros.ganttFechaMin/Max si vienen; si no, este año de negocio completo (según CONFIG.FECHA_CORTE, no el reloj real)
  const anioNegocio = obtenerFechaNegocioActual_().anio;
  const RANGO_MIN = new Date((filtros.ganttFechaMin || (anioNegocio + '-01-01')) + 'T00:00:00');
  const RANGO_MAX = new Date((filtros.ganttFechaMax || (anioNegocio + '-12-31')) + 'T23:59:59');

  const fechas = puntos.map(p => new Date(p.fecha).getTime()).filter(t => {
    const d = new Date(t);
    return d >= RANGO_MIN && d <= RANGO_MAX;
  });
  if (fechas.length === 0) return { semanas: [], filas: [] };

  const inicio = new Date(Math.min.apply(null, fechas));
  inicio.setDate(inicio.getDate() - ((inicio.getDay() + 6) % 7));
  inicio.setHours(0, 0, 0, 0);
  const fin = new Date(Math.min(Math.max.apply(null, fechas), RANGO_MAX.getTime()));

  const semanas = [];
  const cursor = new Date(inicio);
  while (cursor <= fin) {
    const viernes = new Date(cursor);
    viernes.setDate(viernes.getDate() + 4); // lunes + 4 días = viernes de esa semana
    semanas.push({
      mes: Utilities.formatDate(cursor, Session.getScriptTimeZone(), 'MMM yyyy'),
      etiqueta: Utilities.formatDate(viernes, Session.getScriptTimeZone(), 'dd/MM'),
      inicio: Utilities.formatDate(cursor, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    });
    cursor.setDate(cursor.getDate() + 7);
  }

  // Filtrar puntos fuera del rango permitido
  const puntosEnRango = puntos.filter(p => {
    const d = new Date(p.fecha);
    return d >= RANGO_MIN && d <= RANGO_MAX;
  });

  const filasMap = {};
  puntosEnRango.forEach(p => {
    const fecha = new Date(p.fecha);
    const diffSemanas = Math.floor((fecha - inicio) / (7 * 24 * 60 * 60 * 1000));
    if (!filasMap[p.clave]) {
      const [sdatool, proyecto] = p.clave.split('|');
      filasMap[p.clave] = { sdatool, proyecto, dominio: p.dominio||'', dominioSp: p.dominioSp||'', montos: {}, validadosPorSemana: {}, atrasadosPorSemana: {}, regularizadoEnSemana: {} };
    }
    filasMap[p.clave].montos[diffSemanas] = (filasMap[p.clave].montos[diffSemanas] || 0) + p.monto;
    if (p.validado) {
      filasMap[p.clave].validadosPorSemana[diffSemanas] = (filasMap[p.clave].validadosPorSemana[diffSemanas] || 0) + p.monto;
      // Se regularizó tarde si la fecha validada quedó después de la fecha propuesta de ESA semana
      if (p.fechaValidada && new Date(p.fechaValidada) > fecha) {
        filasMap[p.clave].atrasadosPorSemana[diffSemanas] = (filasMap[p.clave].atrasadosPorSemana[diffSemanas] || 0) + p.monto;
        // Además de anotarlo en su semana ORIGINAL, lo anotamos en la semana en que REALMENTE se pagó
        // (para la línea de seguimiento: cuánto se movió de verdad esa semana, venga de donde venga)
        const diffSemanasReal = Math.floor((new Date(p.fechaValidada) - inicio) / (7 * 24 * 60 * 60 * 1000));
        if (diffSemanasReal >= 0) {
          filasMap[p.clave].regularizadoEnSemana[diffSemanasReal] = (filasMap[p.clave].regularizadoEnSemana[diffSemanasReal] || 0) + p.monto;
        }
      }
    }
  });

  const filas = Object.keys(filasMap).sort().map(clave => {
    const f = filasMap[clave];
    return {
      sdatool: f.sdatool,
      proyecto: f.proyecto,
      dominio: f.dominio || '',
      dominioSp: f.dominioSp || '',
      montos: semanas.map((s, i) => f.montos[i] || 0),
      validadosPorSemana: semanas.map((s, i) => f.validadosPorSemana[i] || 0),
      atrasadosPorSemana: semanas.map((s, i) => f.atrasadosPorSemana[i] || 0),
      regularizadoEnSemana: semanas.map((s, i) => f.regularizadoEnSemana[i] || 0),
      validado: Object.keys(f.validadosPorSemana).length > 0,
    };
  });

  // Total por SDATOOL: suma de todas las filas (proyectos) que comparten ese SDATOOL
  const totalPorSdatool = {};
  filas.forEach(f => {
    const total = f.montos.reduce((a,b) => a+b, 0);
    totalPorSdatool[f.sdatool] = (totalPorSdatool[f.sdatool] || 0) + total;
  });
  filas.forEach(f => { f.totalSdatool = totalPorSdatool[f.sdatool]; });

  return { semanas: semanas, filas: filas };
}

/**
 * Compara Fecha de pago propuesta vs Fecha de pago validada por OC, para medir
 * desviación (en días) entre lo planeado y lo real. Solo entran OCs que ya
 * tienen Fecha de pago propuesta cargada.
 */
function calcularValidacionPagos_(filtros) {
  filtros = filtros || {};
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return { filas: [], resumen: {} };

  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos).filter(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return false;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return false;
    if (!coincideFiltro_(filtros.estado, g.estado)) return false;
    if (!coincideFiltro_(filtros.area, g.area)) return false;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return false;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return false;
    return true;
  });

  // Un "pago programado" = una cuota (OC valorizada) o la fecha propuesta de una OC simple / pago parcial.
  let items = construirPagosProgramados_(agrupadas, leerValidacionCuotas_());
  const semanas = semanasDePagos_(items);      // opciones del filtro: todas las semanas que tienen pagos programados
  if (filtros.semana) {
    const luneSem = new Date(filtros.semana);
    const finSem = new Date(luneSem.getTime() + 7 * 86400000 - 1);      // hasta el domingo inclusive
    items = items.filter(it => { const fp = new Date(it.fechaPago); return fp >= luneSem && fp <= finSem; });
  }
  // De mayor a menor monto
  items.sort((a, b) => (b.monto - a.monto) || String(a.oc).localeCompare(String(b.oc)) || (a.cuota - b.cuota));

  const R = { total: items.length, cuotas: 0, validadas: 0, verificadas: 0, revisadoSinPago: 0, pagoParcial: 0,
              pendientesValidar: 0, aTiempo: 0, adelantado: 0, atrasado: 0, desviacionPromedio: 0,
              montoTotal: 0, montoRevisadoSinPago: 0, montoPagoParcial: 0 };
  let sumaDesviacion = 0;

  const filas = items.map(it => {
    let desviacionDias = null, desviacionSemanas = null;
    const est = String(it.estado || '').toLowerCase();
    R.montoTotal += it.monto;
    if (it.esCuota) R.cuotas++;
    if (it.fechaValidada) {
      const propuesta = new Date(it.fechaPago);
      const validada = new Date(it.fechaValidada);
      desviacionDias = Math.round((validada - propuesta) / (24 * 60 * 60 * 1000));
      desviacionSemanas = Math.floor(Math.abs(desviacionDias) / 7) * (desviacionDias < 0 ? -1 : 1);
      R.validadas++;
      sumaDesviacion += desviacionDias;
      if (desviacionDias === 0) R.aTiempo++;
      else if (desviacionDias < 0) R.adelantado++;
      else R.atrasado++;
    }
    if (est === 'verificado') R.verificadas++;
    else if (est === 'revisado - sin pago') { R.revisadoSinPago++; R.montoRevisadoSinPago += it.monto; }
    else if (est === 'pago parcial') { R.pagoParcial++; R.montoPagoParcial += it.monto; }
    if (!it.fechaValidada && est !== 'verificado' && est !== 'revisado - sin pago' && est !== 'pago parcial') R.pendientesValidar++;

    return {
      oc: it.oc, cuota: it.cuota, esCuota: it.esCuota, tipo: it.tipo,
      proyecto: it.g.proyecto, sdatool: it.g.sdatool, area: it.g.area || '',
      monto: it.monto, pagado: it.pagado, totalOC: it.totalOC,
      fechaPago: formatearFecha(it.fechaPago),
      fechaPagoValidada: formatearFecha(it.fechaValidada),
      desviacionDias: desviacionDias, desviacionSemanas: desviacionSemanas,
      estadoVerificacion: it.g.estadoVerificacion || '',
      estadoValidacion: it.estado || '',
    };
  });
  R.desviacionPromedio = R.validadas > 0 ? (sumaDesviacion / R.validadas) : 0;
  return { filas: filas, resumen: R, semanas: semanas };
}

/**
 * Monto SOLPED / OC pendiente / Pagado, sumado por Dominio (todas las OCs,
 * no solo las que tienen Fecha de pago). Para los números grandes del banner
 * de Detalle de OC.
 */
function obtenerResumenPorDominio(filtros) {
  filtros = filtros || {};
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return {};
  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  const porDominio = {};
  agrupadas.forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.estado, g.estado)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;

    const clave = g.dominio || 'Sin dominio';
    if (!porDominio[clave]) porDominio[clave] = { solped: 0, pendienteOC: 0, pagado: 0 };
    porDominio[clave].solped += g.solped;
    porDominio[clave].pendienteOC += g.pendienteOC;
    porDominio[clave].pagado += g.pagado;
  });
  return porDominio;
}

/**
 * Monto SOLPED / OC pendiente / Pagado, sumado por Trimestre.
 */
function obtenerResumenPorTrimestre(filtros) {
  filtros = filtros || {};
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return {};
  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  const porTrimestre = {};
  agrupadas.forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.estado, g.estado)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;

    const clave = g.trimestre || 'Sin trimestre';
    if (!porTrimestre[clave]) porTrimestre[clave] = { solped: 0, pendienteOC: 0, pagado: 0 };
    porTrimestre[clave].solped += g.solped;
    porTrimestre[clave].pendienteOC += g.pendienteOC;
    porTrimestre[clave].pagado += g.pagado;
  });
  return porTrimestre;
}
/**
 * Resume Monto SOLPED / OC pendiente / Pagado por Área, respetando los filtros
 * compartidos. Para el banner de la pestaña Detalle de OC.
 */
function obtenerResumenPorArea(filtros) {
  filtros = filtros || {};
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return { porArea: {}, totalGeneral: {solped:0,pendienteOC:0,pagado:0} };

  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  const porArea = {};
  const totalGeneral = { solped: 0, pendienteOC: 0, pagado: 0 };

  agrupadas.forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.estado, g.estado)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;

    const clave = g.area || 'Sin área';
    if (!porArea[clave]) porArea[clave] = { solped: 0, pendienteOC: 0, pagado: 0 };
    porArea[clave].solped += g.solped;
    porArea[clave].pendienteOC += g.pendienteOC;
    porArea[clave].pagado += g.pagado;

    totalGeneral.solped += g.solped;
    totalGeneral.pendienteOC += g.pendienteOC;
    totalGeneral.pagado += g.pagado;
  });

  return { porArea: porArea, totalGeneral: totalGeneral };
}
/**
 * Resume el monto proyectado de pago (pendiente de OC con Fecha de pago propuesta
 * cargada) agrupado por Dominio y, dentro de cada uno, por Área. Para el banner
 * desplegable de la pestaña Cash Flow.
 */
function obtenerResumenPorDominioArea(filtros) {
  filtros = filtros || {};
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return { porDominio: {}, totalGeneral: 0 };

  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  const porDominio = {};
  let totalGeneral = 0;

  agrupadas.forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.estado, g.estado)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (!(g.fechaPago && g.pendienteOC > 0)) return;

    const dom = g.dominio || 'Sin dominio';
    const area = g.area || 'Sin área';
    if (!porDominio[dom]) porDominio[dom] = { total: 0, areas: {} };
    porDominio[dom].total += g.pendienteOC;
    porDominio[dom].areas[area] = (porDominio[dom].areas[area] || 0) + g.pendienteOC;
    totalGeneral += g.pendienteOC;
  });

  return { porDominio: porDominio, totalGeneral: totalGeneral };
}
/**
 * Desglose adicional del banner de Detalle de OC: para P&S, por Área; para
 * Tecnología y Seguridad, por SDATOOL (no tienen la misma estructura de Área).
 */
function obtenerDesgloseDominios(filtros) {
  filtros = filtros || {};
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return {};
  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  const resultado = { 'P&S': {}, 'Tecnología': {}, 'Seguridad': {} };
  agrupadas.forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.estado, g.estado)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (!resultado[g.dominio]) return;

    const clave = g.dominio === 'P&S' ? (g.area || 'Sin área') : (g.sdatool || 'Sin SDATOOL');
    if (!resultado[g.dominio][clave]) resultado[g.dominio][clave] = { solped: 0, oc: 0, pagado: 0 };
    resultado[g.dominio][clave].solped += g.solped;
    resultado[g.dominio][clave].oc += g.pendienteOC;
    resultado[g.dominio][clave].pagado += g.pagado;
  });
  return resultado;
}

/**
 * true si `valor` cumple el filtro `seleccion`. `seleccion` es un array de valores
 * marcados por el usuario (multi-select) — si viene vacío o no viene, no filtra nada.
 */
function coincideFiltro_(seleccion, valor) {
  if (!seleccion || !seleccion.length) return true;
  return seleccion.indexOf(valor) !== -1;
}

/**
 * Agrupa las filas de BASE_CASHFLOW por N° de OC, sumando los montos de todas
 * sus líneas/posiciones. Los campos descriptivos toman el primer valor no vacío
 * encontrado; el Estado se recalcula sobre los montos ya sumados.
 */
function agruparPorOC_(datos, incluirEliminadas) {
  const mapa = {};
  const orden = [];
  datos.forEach(fila => {
    const oc = fila[0];
    if (!oc) return;
    if (!mapa[oc]) {
      mapa[oc] = {
        oc: oc, fechaMod: fila[1], proveedor: fila[2], proyecto: fila[3], oficina: fila[4],
        dominio: fila[5], area: fila[6], sdatool: fila[7], trimestre: fila[8],
        contrato: fila[9], solicitante: fila[10], gb: fila[11],
        solped: 0, pendienteOC: 0, pagado: 0,
        proyectoEspecifico: fila[15], intervencion: fila[16], fechaPago: fila[18], comentarios: fila[19],
        fechaPagoValidada: fila[20],
        valorizado: fila[21],
        fechaPago1: fila[22], montoPago1: fila[23], fechaPago2: fila[24], montoPago2: fila[25],
        fechaPago3: fila[26], montoPago3: fila[27], fechaPago4: fila[28], montoPago4: fila[29],
        fechaPago5: fila[30], montoPago5: fila[31],
        areaCorregida: fila[32], estadoVerificacion: fila[33],
        dominioSp: fila[35], estadoValidacion: fila[36] || '',
        lineas: 0,
      };
      orden.push(oc);
    }
    const g = mapa[oc];
    g.solped += Number(fila[12]) || 0;
    g.pendienteOC += Number(fila[13]) || 0;
    g.pagado += Number(fila[14]) || 0;
    g.lineas += 1;
    if (!g.proyectoEspecifico && fila[15]) g.proyectoEspecifico = fila[15];
    if (!g.intervencion && fila[16]) g.intervencion = fila[16];
    if (!g.fechaPago && fila[18]) g.fechaPago = fila[18];
    if (!g.comentarios && fila[19]) g.comentarios = fila[19];
    if (!g.fechaPagoValidada && fila[20]) g.fechaPagoValidada = fila[20];
    if (!g.valorizado && fila[21]) g.valorizado = fila[21];
    if (!g.fechaPago1 && fila[22]) { g.fechaPago1 = fila[22]; g.montoPago1 = fila[23]; }
    if (!g.fechaPago2 && fila[24]) { g.fechaPago2 = fila[24]; g.montoPago2 = fila[25]; }
    if (!g.fechaPago3 && fila[26]) { g.fechaPago3 = fila[26]; g.montoPago3 = fila[27]; }
    if (!g.fechaPago4 && fila[28]) { g.fechaPago4 = fila[28]; g.montoPago4 = fila[29]; }
    if (!g.fechaPago5 && fila[30]) { g.fechaPago5 = fila[30]; g.montoPago5 = fila[31]; }
    if (!g.areaCorregida && fila[32]) g.areaCorregida = fila[32];
    if (!g.estadoVerificacion && fila[33]) g.estadoVerificacion = fila[33];
    if (!g.estadoManual && fila[34]) g.estadoManual = fila[34];
    if (!g.dominioSp && fila[35]) g.dominioSp = fila[35];
    if (!g.estadoValidacion && fila[36]) g.estadoValidacion = fila[36];
  });
  return orden.map(oc => {
    const g = mapa[oc];
    g.estado = g.estadoManual || calcularEstado(g.solped, g.pendienteOC, g.pagado);
    // Si el área fue corregida manualmente, esa versión manda en toda la app (filtros, resúmenes, etc.)
    if (g.areaCorregida) g.area = g.areaCorregida;
    return g;
  }).filter(g => incluirEliminadas || g.estado !== 'Eliminado');   // una OC eliminada (aprobada por Maestro) no entra en ningún cálculo
}

/**
 * Resumen agregado para el Dashboard — calculado en el servidor sobre OCs YA
 * AGRUPADAS (no líneas sueltas), así el navegador nunca recibe las 15,000+ filas
 * completas solo para mostrar totales y KPIs.
 * Los montos de Comprometido / Por pagar / Pagado son la suma directa de sus
 * columnas (SOLPED, Pendiente/OC, Pagado) sobre todas las OCs que pasan el filtro —
 * así el total siempre cuadra con la suma de las 3 columnas fuente, sin duplicar nada.
 * "Pagado parcialmente" se reporta aparte, solo como conteo informativo (esas OCs
 * ya están sumadas dentro de Por pagar y Pagado, cada una por su columna real).
 */
function obtenerResumenDashboard(filtros, datosPre) {
  filtros = filtros || {};
  let datos = datosPre;
  if (!datos) {
    const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
    if (!hojaBase || hojaBase.getLastRow() < 2) {
      return { totales: {}, conteos: {}, porSdatool: {}, proyectos: [], sdatools: [], areas: [], trimestres: [], dominios: [], estadosVerificacion: [], dominiosSp: [] };
    }
    datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  }
  const agrupadas = agruparPorOC_(datos);

  const ESTADO_KEY = {
    'Por comprometer': 'comprometer', 'Por pagar': 'pagar',
    'Pagado parcialmente': 'parcial', 'Pagado': 'pagado',
  };
  const totales = { comprometer: 0, pagar: 0, pagado: 0 };
  const conteos = { comprometer: 0, pagar: 0, parcial: 0, pagado: 0 };
  const porSdatool = {};
  const porProyecto = {};
  const proyectosSet = {}, sdatoolsSet = {}, areasSet = {}, trimestresSet = {}, dominiosSet = {}, estadosVerifSet = {}, dominiosSpSet = {};

  agrupadas.forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.estado, g.estado)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (!coincideFiltro_(filtros.estadoVerificacion, g.estadoVerificacion || '(Sin estado)')) return;
    if (!coincideFiltro_(filtros.dominioSp, (g.dominioSp||'').trim() || '(Sin dato)')) return;

    totales.comprometer += g.solped;
    totales.pagar += g.pendienteOC;
    totales.pagado += g.pagado;
    const key = ESTADO_KEY[g.estado] || 'comprometer';
    conteos[key] = (conteos[key] || 0) + 1;

    if (g.proyecto) proyectosSet[g.proyecto] = true;
    if (g.sdatool) sdatoolsSet[g.sdatool] = true;
    if (g.area) areasSet[g.area] = true;
    if (g.trimestre) trimestresSet[g.trimestre] = true;
    if (g.dominio) dominiosSet[g.dominio] = true;
    estadosVerifSet[g.estadoVerificacion || '(Sin estado)'] = true;
    dominiosSpSet[(g.dominioSp||'').trim() || '(Sin dato)'] = true;

    const claveSd = g.sdatool || 'Sin SDATOOL';
    if (!porSdatool[claveSd]) porSdatool[claveSd] = { comprometer: 0, pagar: 0, pagado: 0, count: 0 };
    porSdatool[claveSd].comprometer += g.solped;
    porSdatool[claveSd].pagar += g.pendienteOC;
    porSdatool[claveSd].pagado += g.pagado;
    porSdatool[claveSd].count += 1;

    const claveProy = g.proyecto || 'Sin proyecto';
    if (!porProyecto[claveProy]) porProyecto[claveProy] = { comprometer: 0, pagar: 0, pagado: 0, count: 0 };
    porProyecto[claveProy].comprometer += g.solped;
    porProyecto[claveProy].pagar += g.pendienteOC;
    porProyecto[claveProy].pagado += g.pagado;
    porProyecto[claveProy].count += 1;
  });

  return {
    totales, conteos, porSdatool, porProyecto,
    fechaCorte: CONFIG.FECHA_CORTE,
    proyectos: Object.keys(proyectosSet).sort(),
    sdatools: Object.keys(sdatoolsSet).sort(),
    areas: Object.keys(areasSet).sort(),
    trimestres: Object.keys(trimestresSet).sort(),
    dominios: Object.keys(dominiosSet).sort(),
    estadosVerificacion: Object.keys(estadosVerifSet).sort(),
    dominiosSp: Object.keys(dominiosSpSet).sort(),
  };
}

/**
 * Devuelve todas las filas de BASE_CASHFLOW como objetos, para que el HTML
 * las pinte con google.script.run. Se llama desde el frontend al cargar la página.
 */
/**
 * Devuelve una PÁGINA de OCs YA AGRUPADAS (una fila por N° de OC, montos sumados
 * entre todas sus posiciones), aplicando filtros en el servidor.
 * filtros: { proyecto, estado, sdatool, area, trimestre, dominio, busqueda } — todos opcionales.
 * pagina: número de página (empieza en 0). tamanoPagina: OCs por página.
 */
function obtenerDatosCashFlow(filtros, pagina, tamanoPagina) {
  filtros = filtros || {};
  pagina = pagina || 0;
  tamanoPagina = tamanoPagina || 100;

  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return { filas: [], total: 0 };

  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  const filtradas = agrupadas.filter(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return false;
    if (!coincideFiltro_(filtros.estado, g.estado)) return false;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return false;
    if (!coincideFiltro_(filtros.area, g.area)) return false;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return false;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return false;
    if (!coincideFiltro_(filtros.estadoVerificacion, g.estadoVerificacion || '(Sin estado)')) return false;
    if (!coincideFiltro_(filtros.dominioSp, (g.dominioSp||'').trim() || '(Sin dato)')) return false;
    if (filtros.sinFechaPago) { if (g.fechaPago) return false; }
    else if (filtros.semanaFechaPago && g.fechaPago) {
      const lunes = new Date(filtros.semanaFechaPago);
      const vier = new Date(lunes); vier.setDate(vier.getDate() + 6);
      const fp = new Date(g.fechaPago);
      if (fp < lunes || fp > vier) return false;
    } else if (filtros.semanaFechaPago && !g.fechaPago) return false;
    if (filtros.busqueda) {
      // Sin distinguir mayúsculas/minúsculas, y busca en OC, Proveedor, Proyecto y SDATOOL
      const q = String(filtros.busqueda).trim().toUpperCase();
      const enOC = String(g.oc || '').toUpperCase().indexOf(q) !== -1;
      const enProveedor = String(g.proveedor || '').toUpperCase().indexOf(q) !== -1;
      const enProyecto = String(g.proyecto || '').toUpperCase().indexOf(q) !== -1;
      const enSdatool = String(g.sdatool || '').toUpperCase().indexOf(q) !== -1;
      if (!enOC && !enProveedor && !enProyecto && !enSdatool) return false;
    }
    return true;
  });

  const total = filtradas.length;
  const paginaDatos = filtradas.slice(pagina * tamanoPagina, (pagina + 1) * tamanoPagina);

  const vcMap = leerValidacionCuotas_();
  const pendElim = leerSolicitudesEliminacionPendientes_();
  const filas = paginaDatos.map(g => ({
    oc: g.oc,
    fechaMod: formatearFecha(g.fechaMod),
    proveedor: g.proveedor,
    proyecto: g.proyecto,
    oficina: g.oficina,
    dominio: g.dominio,
    area: g.area,
    sdatool: g.sdatool,
    trimestre: g.trimestre,
    contrato: g.contrato,
    solicitante: g.solicitante,
    gb: g.gb,
    solped: g.solped,
    pendienteOC: g.pendienteOC,
    pagado: g.pagado,
    proyectoEspecifico: g.proyectoEspecifico,
    intervencion: g.intervencion,
    estado: g.estado,
    fechaPago: formatearFecha(g.fechaPago),
    comentarios: g.comentarios,
    fechaPagoValidada: formatearFecha(g.fechaPagoValidada),
    valorizado: !!g.valorizado,
    estadoVerificacion: g.estadoVerificacion || '',
    estadoValidacion: g.estadoValidacion || '',
    dominioSp: g.dominioSp || '',
    cuotas: [1,2,3,4,5].map(n => ({
      fecha: formatearFecha(g['fechaPago'+n]),
      monto: Number(g['montoPago'+n]) || 0,
      estado: (vcMap[String(g.oc) + '|' + n] || {}).estado || '',
      fechaValidada: formatearFecha((vcMap[String(g.oc) + '|' + n] || {}).fechaValidada),
    })),
    lineas: g.lineas,
    eliminacionPendiente: !!pendElim[String(g.oc)],
    eliminacionInfo: pendElim[String(g.oc)] ? ('Solicitada por ' + pendElim[String(g.oc)].solicitante + ' — ' + pendElim[String(g.oc)].motivo) : '',
  }));

  return { filas: filas, total: total };
}

/**
 * Versión liviana de OC con Estado = "Pagado", pensada para el botón "Ver y validar
 * todo lo Pagado" de Validación de Pagos. A diferencia de obtenerDatosCashFlow, NO
 * arma cuotas, info de eliminación ni el resto de columnas — solo los 5 campos que
 * ese modal necesita, para que la respuesta sea chica incluso con miles de OC.
 * Se cachea unos minutos (se invalida solo, igual que el resto, al sincronizar o editar).
 */
function obtenerOCPagadasLigero() {
  return conCache_('ocpagadas', {}, () => {
    const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
    if (!hojaBase || hojaBase.getLastRow() < 2) return { filas: [] };
    const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
    const agrupadas = agruparPorOC_(datos);
    const filas = agrupadas
      .filter(g => g.estado === 'Pagado')
      .map(g => ({
        oc: g.oc,
        proyecto: g.proyecto || '',
        monto: (Number(g.pendienteOC) || 0) + (Number(g.pagado) || 0),
        estadoValidacion: g.estadoValidacion || '',
        fechaMod: formatearFecha(g.fechaMod),
      }));
    return { filas: filas };
  });
}

function formatearFecha(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return valor;
}

/**
 * Sirve el HTML como app web independiente.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('CashFlowOC').setTitle('Dashboard Gestión Presupuestal');
}
/**
 * Registra una visita a la app (correo + fecha/hora) en la pestaña Accesos,
 * para poder medir después uso/frecuencia por usuario. Se llama una vez por carga.
 */
function registrarAcceso() {
  const correo = obtenerCorreoVisitante_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Accesos');
  if (!hoja) {
    hoja = ss.insertSheet('Accesos');
    hoja.getRange(1, 1, 1, 2).setValues([['Correo', 'Fecha y hora']]);
  }
  hoja.appendRow([correo, new Date()]);
  return { ok: true };
}

/**
 * true si el usuario actual tiene Rol de Editor o Programador en la pestaña
 * Usuarios (columna D) — controla el acceso a botones de diagnóstico técnico.
 */
function esUsuarioAdmin() {
  const correo = obtenerCorreoVisitante_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Usuarios');
  if (!hoja || hoja.getLastRow() < 2) return false;
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues();
  const fila = datos.find(f => String(f[0]).toLowerCase() === String(correo).toLowerCase());
  if (!fila) return false;
  const rol = String(fila[3] || '').toLowerCase();
  return rol.indexOf('editor') !== -1 || rol.indexOf('programador') !== -1 || rol.indexOf('admin') !== -1;
}

/**
 * Correo real de quien está usando la app AHORA MISMO (no el dueño del script).
 * Requiere que la implementación esté configurada como "Ejecutar como: Usuario
 * que accede a la app web" y "Quién tiene acceso" restringido a tu organización —
 * si no, Apps Script no puede saber quién es el visitante y esto devuelve ''.
 */
function obtenerCorreoVisitante_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

/**
 * Verifica si el correo de sesión está en la pestaña Usuarios. Si NO lo está,
 * avisa por correo a todos los que tengan Rol Editor/Programador (a lo más una
 * vez cada 24h por persona, registrado en la pestaña SolicitudesAcceso).
 */
function verificarAcceso() {
  const correo = obtenerCorreoVisitante_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaUsuarios = ss.getSheetByName('Usuarios');
  let autorizado = false;
  let nombre = correo;

  if (correo && hojaUsuarios && hojaUsuarios.getLastRow() >= 2) {
    const datos = hojaUsuarios.getRange(2, 1, hojaUsuarios.getLastRow() - 1, 2).getValues();
    const fila = datos.find(f => String(f[0]).toLowerCase() === String(correo).toLowerCase());
    if (fila) { autorizado = true; nombre = fila[1] || correo; }
  }

  if (!autorizado) {
    notificarSolicitudAcceso_(correo || '(correo no identificado)');
  }

  return { autorizado: autorizado, correo: correo, nombre: nombre };
}

function notificarSolicitudAcceso_(correo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaLog = ss.getSheetByName('SolicitudesAcceso');
  if (!hojaLog) {
    hojaLog = ss.insertSheet('SolicitudesAcceso');
    hojaLog.getRange(1, 1, 1, 2).setValues([['Correo', 'Última solicitud']]);
  }

  // Evita reenviar el mismo correo más de una vez cada 24h
  const datos = hojaLog.getLastRow() >= 2 ? hojaLog.getRange(2, 1, hojaLog.getLastRow() - 1, 2).getValues() : [];
  const idx = datos.findIndex(f => String(f[0]).toLowerCase() === String(correo).toLowerCase());
  const ahora = new Date();
  if (idx !== -1) {
    const ultima = new Date(datos[idx][1]);
    if (ahora - ultima < 24 * 60 * 60 * 1000) return; // ya se avisó hoy
    hojaLog.getRange(idx + 2, 2).setValue(ahora);
  } else {
    hojaLog.appendRow([correo, ahora]);
  }

  const hojaUsuarios = ss.getSheetByName('Usuarios');
  if (!hojaUsuarios || hojaUsuarios.getLastRow() < 2) return;
  const usuarios = hojaUsuarios.getRange(2, 1, hojaUsuarios.getLastRow() - 1, 4).getValues();
  const admins = usuarios.filter(f => {
    const rol = String(f[3] || '').toLowerCase();
    return rol.indexOf('editor') !== -1 || rol.indexOf('programador') !== -1 || rol.indexOf('admin') !== -1;
  }).map(f => f[0]).filter(Boolean);

  if (admins.length === 0) return;
  MailApp.sendEmail(
    admins.join(','),
    'Solicitud de acceso — Cash Flow OC',
    'El usuario ' + correo + ' intentó ingresar a Cash Flow OC y no está en la lista de Usuarios autorizados.\n\n' +
    'Si corresponde darle acceso, agrégalo en la pestaña Usuarios (Correo, Nombre, Área, Rol).'
  );
}

/**
 * Lee la pestaña FasesAvance y agrupa las etapas por Proceso, para los botones
 * de "Resumen de avance" en Vista General. Columnas esperadas: Proceso, Etapa,
 * Porcentaje, Detalle (monto o texto libre, opcional).
 */
function obtenerFasesAvance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('FasesAvance');
  if (!hoja) {
    hoja = ss.insertSheet('FasesAvance');
    hoja.getRange(1, 1, 1, 4).setValues([['Proceso', 'Etapa', 'Porcentaje', 'Detalle']]);
    return {};
  }
  if (hoja.getLastRow() < 2) return {};
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues();
  const porProceso = {};
  datos.forEach(fila => {
    const proceso = String(fila[0] || '').trim();
    if (!proceso) return;
    if (!porProceso[proceso]) porProceso[proceso] = [];
    porProceso[proceso].push({
      etapa: fila[1] || '',
      porcentaje: Number(fila[2]) || 0,
      detalle: fila[3] || '',
    });
  });
  return porProceso;
}

/**
 * Guarda una nota/actividad para una fecha del calendario de Vista General.
 */
function guardarNotaCalendario(fecha, nota) {
  if (!fecha || !nota) return { ok: false, error: 'Falta fecha o nota' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('CalendarioNotas');
  if (!hoja) {
    hoja = ss.insertSheet('CalendarioNotas');
    hoja.getRange(1, 1, 1, 5).setValues([['Fecha', 'Nota', 'Usuario', 'Registrado', 'Estado']]);
  }
  const usuario = obtenerCorreoVisitante_();
  hoja.appendRow([fecha, nota, usuario, new Date(), 'Por confirmar']);
  return { ok: true };
}

/**
 * Devuelve todas las notas del calendario, agrupadas por fecha (yyyy-MM-dd).
 */
function obtenerNotasCalendario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('CalendarioNotas');
  if (!hoja || hoja.getLastRow() < 2) return {};
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues();
  const porFecha = {};
  datos.forEach((fila, i) => {
    const fecha = formatearFecha(fila[0]) || String(fila[0]);
    if (!fecha) return;
    if (!porFecha[fecha]) porFecha[fecha] = [];
    porFecha[fecha].push({ fila: i + 2, nota: fila[1], usuario: fila[2], estado: fila[4] || 'Por confirmar' });
  });
  return porFecha;
}

/** El admin confirma una actividad del calendario. */
function confirmarNotaCalendario(filaNumero) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('CalendarioNotas');
  if (!hoja) return { ok: false, error: 'No existe CalendarioNotas' };
  hoja.getRange(filaNumero, 5).setValue('Confirmado');
  return { ok: true };
}

/**
 * Cuenta los usuarios autorizados registrados en la pestaña Usuarios (columna Correo).
 * Crea la pestaña con encabezados si todavía no existe.
 */
function obtenerContadorUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Usuarios');
  if (!hoja) {
    hoja = ss.insertSheet('Usuarios');
    hoja.getRange(1, 1, 1, 4).setValues([['Correo', 'Nombre', 'Área', 'Rol']]);
    return 0;
  }
  return Math.max(0, hoja.getLastRow() - 1);
}

/**
 * Identifica a quien está usando la app ahora mismo (por su correo de sesión) y
 * busca su Nombre/Área en la pestaña Usuarios. Si no lo encuentra, usa el correo.
 */
function obtenerUsuarioActual() {
  const correo = obtenerCorreoVisitante_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Usuarios');
  if (hoja && hoja.getLastRow() >= 2) {
    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues();
    const fila = datos.find(f => String(f[0]).trim().toLowerCase() === String(correo).trim().toLowerCase());
    if (fila) {
      return { correo: correo, nombre: fila[1] || correo, area: fila[2] || '', rol: fila[3] || '', subRol: fila[4] || '' };
    }
  }
  return { correo: correo, nombre: correo || 'Usuario', area: '', rol: '', subRol: '' };
}

function esLider_(subRol) {
  return String(subRol || '').toLowerCase().includes('l\u00edder') || String(subRol || '').toLowerCase().includes('lider');
}
function esMaestro_(subRol) {
  return String(subRol || '').toLowerCase().includes('maestro');
}

/** Sub Rol del usuario que visita, para decidir accesos en el portal. */
function obtenerSubRolVisitante() {
  const u = obtenerUsuarioActual();
  return { subRol: u.subRol, esLider: esLider_(u.subRol), esMaestro: esMaestro_(u.subRol), nombre: u.nombre, correo: u.correo };
}

/**
 * Lee la pestaña "Data" y arma los 5 indicadores CAPEX (Autorizado, Solped+IGV,
 * OC, Ejecutado, Disponible) sumados por dominio, filtrando por T&C y por
 * Trimestre si se indica. Igual que BASE_FUENTE: solo entra lo que sea T&C
 * y esté en un dominio válido (P&S, Tecnología, Seguridad).
 */
const COL_DATA = { AUTORIZADO: 5, SOLPED: 7, OC: 8, EJECUTADO: 9, DISPONIBLE: 10, TRIMESTRE: 14, DOMINIO: 16, TC: 18 };

function obtenerCapexPorDominio(filtros) {
  filtros = filtros || {};
  const hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  if (!hojaData || hojaData.getLastRow() < 2) return {};

  const datos = hojaData.getRange(2, 1, hojaData.getLastRow() - 1, hojaData.getLastColumn()).getValues();

  const resultado = {};
  CONFIG.DOMINIOS_VALIDOS.forEach(d => {
    resultado[d] = { autorizado: 0, solped: 0, oc: 0, ejecutado: 0, disponible: 0 };
  });

  datos.forEach(fila => {
    const tc = fila[COL_DATA.TC - 1];
    const dominio = fila[COL_DATA.DOMINIO - 1];
    const trimestre = fila[COL_DATA.TRIMESTRE - 1];
    if (tc !== CONFIG.TC_VALIDO) return;
    if (CONFIG.DOMINIOS_VALIDOS.indexOf(dominio) === -1) return;
    if (!coincideFiltro_(filtros.trimestre, trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, dominio)) return;

    resultado[dominio].autorizado += Number(fila[COL_DATA.AUTORIZADO - 1]) || 0;
    resultado[dominio].solped += Number(fila[COL_DATA.SOLPED - 1]) || 0;
    resultado[dominio].oc += Number(fila[COL_DATA.OC - 1]) || 0;
    resultado[dominio].ejecutado += Number(fila[COL_DATA.EJECUTADO - 1]) || 0;
    resultado[dominio].disponible += Number(fila[COL_DATA.DISPONIBLE - 1]) || 0;
  });

  const total = { autorizado: 0, solped: 0, oc: 0, ejecutado: 0, disponible: 0 };
  Object.values(resultado).forEach(d => {
    total.autorizado += d.autorizado;
    total.solped += d.solped;
    total.oc += d.oc;
    total.ejecutado += d.ejecutado;
    total.disponible += d.disponible;
  });
  resultado['Total general'] = total;

  return resultado;
}

/**
 * Registra una solicitud de código de proyecto (CAPEX u OPEX): la agrega como
 * fila nueva en la pestaña Solicitudes y avisa por correo. Cuando exista la
 * pestaña Destinatarios, aquí se debe reemplazar CONFIG.EMAIL_SOLICITUDES por
 * una búsqueda del correo correspondiente según Dominio/Tipo.
 */
function enviarSolicitudCodigo(datos) {
  datos = datos || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Solicitudes');
  if (!hoja) {
    hoja = ss.insertSheet('Solicitudes');
    hoja.getRange(1, 1, 1, ENCABEZADOS_SOLICITUDES.length).setValues([ENCABEZADOS_SOLICITUDES]);
  }

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const valoresPorNombre = {
    'Fecha': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    'Tipo': datos.tipo || '',
    'Nombre del proyecto': datos.nombreProyecto || '',
    'Descripción': datos.descripcion || '',
    'Fecha de SOLPED': datos.fechaSolped || '',
    'Fecha de pliego técnico': datos.fechaPliego || '',
    'Importe global': Number(datos.importeGlobal) || '',
    'Fecha de inicio del proyecto': datos.fechaInicio || '',
    'Fecha de fin del proyecto': datos.fechaFin || '',
    'Estado': 'Solicitud Enviada',
    'Código': '',
  };
  const filaNueva = encabezados.map(h => valoresPorNombre.hasOwnProperty(h) ? valoresPorNombre[h] : '');
  hoja.appendRow(filaNueva);
  const filaInsertada = hoja.getLastRow();

  // Columnas fijas agregadas manualmente en la hoja (fuera del esquema por
  // nombre de encabezado de arriba): G = nombre del solicitante; M = correo
  // del solicitante; N en adelante = moneda, tipo de cambio, importe
  // digitado, si incluye IGV y el monto final ya convertido a soles y con IGV aplicado.
  const correoUsuario_ = resolverCorreoSolicitante_(datos.nombreSolicitante);
  hoja.getRange(filaInsertada, 7).setValue(datos.nombreSolicitante || ''); // G: nombre del solicitante
  hoja.getRange(filaInsertada, 13).setValue(correoUsuario_ || ''); // M: correo
  hoja.getRange(filaInsertada, 14, 1, 5).setValues([[
    datos.moneda || '',
    Number(datos.tipoCambio) || '',
    Number(datos.importeGlobal) || '',
    datos.incluyeIGV === 'No' ? 'No' : (datos.incluyeIGV === 'Si' ? 'Sí' : ''),
    Number(datos.montoFinal) || '',
  ]]); // N: Moneda, O: Tipo de cambio, P: Importe digitado, Q: ¿Incluye IGV?, R: Monto final (S/.)

  const asunto = 'Solicitud de código ' + (datos.tipo || '') + ' — ' + (datos.nombreProyecto || 'Sin nombre');
  const cuerpo =
    'Se registró una nueva solicitud de código de proyecto en Cash Flow OC.\n\n' +
    'Solicitante: ' + (datos.nombreSolicitante || '') + '\n' +
    'Tipo: ' + (datos.tipo || '') + '\n' +
    'Nombre del proyecto: ' + (datos.nombreProyecto || '') + '\n' +
    'Descripción: ' + (datos.descripcion || '') + '\n' +
    'Fecha de SOLPED (referencial): ' + (datos.fechaSolped || '') + '\n' +
    'Fecha de pliego técnico (referencial): ' + (datos.fechaPliego || '') + '\n' +
    'Moneda: ' + (datos.moneda || '') + '\n' +
    (datos.moneda && datos.moneda !== 'Soles' ? 'Tipo de cambio: ' + (datos.tipoCambio || '') + '\n' : '') +
    'Importe digitado: ' + (datos.importeGlobal || '') + ' ' + (datos.moneda || '') + '\n' +
    '¿Incluye IGV?: ' + (datos.incluyeIGV === 'No' ? 'No' : 'Sí') + '\n' +
    'Monto final (S/., con IGV): ' + (datos.montoFinal ? Number(datos.montoFinal).toFixed(2) : '') + '\n' +
    'Fecha de inicio del proyecto (referencial): ' + (datos.fechaInicio || '') + '\n' +
    'Fecha de fin del proyecto (referencial): ' + (datos.fechaFin || '');

  const admins = obtenerCorreosAdmin_();
  const destinatarios = admins.length > 0 ? admins.join(',') : CONFIG.EMAIL_SOLICITUDES;
  MailApp.sendEmail({ to: destinatarios, subject: asunto, body: cuerpo });

  // Copia para quien envía la solicitud: va como correo APARTE (no en CC), así siempre le llega
  // aunque también esté entre los destinatarios. Si falla, no se pierde la solicitud ya registrada.
  let copiaEnviadaA = '';
  if (correoUsuario_) {
    try {
      MailApp.sendEmail({
        to: correoUsuario_,
        subject: 'Copia de tu solicitud — ' + asunto,
        body: 'Esta es una copia de la solicitud que acabas de enviar desde Cash Flow OC.\n\n' + cuerpo,
      });
      copiaEnviadaA = correoUsuario_;
    } catch (eCopia) {
      Logger.log('No se pudo enviar la copia al solicitante: ' + eCopia);
    }
  }
  return { ok: true, copiaEnviadaA: copiaEnviadaA };
}

/**
 * Correo de quien envía la solicitud. Primero el de la sesión; si Apps Script no lo entrega
 * (según cómo esté publicada la app), se busca por el nombre en la pestaña Usuarios.
 */
function resolverCorreoSolicitante_(nombre) {
  const directo = obtenerCorreoVisitante_();
  if (directo) return directo;
  const n = String(nombre || '').trim().toLowerCase();
  if (!n) return '';
  const mapa = obtenerMapaCorreoNombre_();
  const coincidencias = Object.keys(mapa).filter(c => String(mapa[c] || '').toLowerCase() === n);
  return coincidencias.length === 1 ? coincidencias[0] : '';
}

/** Devuelve los correos de los usuarios con Rol Editor o Programador (los admins). */
function obtenerCorreosAdmin_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Usuarios');
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues();
  return datos
    .filter(f => {
      const rol = String(f[3] || '').toLowerCase();
      return rol.indexOf('editor') !== -1 || rol.indexOf('programador') !== -1 || rol.indexOf('admin') !== -1;
    })
    .map(f => f[0])
    .filter(Boolean);
}

const ENCABEZADOS_SOLICITUDES = [
  'Fecha', 'Tipo', 'Nombre del proyecto', 'Descripción',
  'Fecha de SOLPED', 'Fecha de pliego técnico', 'Importe global',
  'Fecha de inicio del proyecto', 'Fecha de fin del proyecto', 'Estado', 'Código',
];

/**
 * Devuelve todas las solicitudes registradas en la pestaña Solicitudes, con el
 * número de fila real (para poder actualizarlas después). Lee por NOMBRE de
 * encabezado (no por posición fija), así sigue funcionando aunque la hoja
 * tenga columnas en otro orden o de una versión anterior del formulario.
 */
function obtenerSolicitudesArray_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Solicitudes');
  if (!hoja || hoja.getLastRow() < 2) return [];

  const encabezadosCrudos = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const encabezados = encabezadosCrudos.map(h => String(h).trim().toLowerCase());
  const numCols = hoja.getLastColumn();
  const idx = (nombre, posicionRespaldo) => {
    const i = encabezados.indexOf(String(nombre).trim().toLowerCase());
    if (i !== -1) return i;
    // la posición de respaldo solo sirve si existe de verdad en la hoja
    return (posicionRespaldo !== undefined && posicionRespaldo < numCols) ? posicionRespaldo : -1;
  };

  // Si el nombre exacto no calza (mayúsculas, espacios, hoja vieja), cae a la
  // posición del esquema actual (ENCABEZADOS_SOLICITUDES) como respaldo.
  const iFecha = idx('Fecha', 0), iTipo = idx('Tipo', 1), iProyecto = idx('Nombre del proyecto', 2),
        iEstado = idx('Estado', 9), iCodigo = idx('Código', 10);

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
  const mapaUsuarios = obtenerMapaCorreoNombre_();
  const lista = datos
    .map((fila, i) => {
      const correoSolicitante = numCols >= 13 ? String(fila[12] || '').trim() : '';
      const nombreCol = numCols >= 7 ? String(fila[6] || '').trim() : '';
      // Si no se guardó el nombre a mano (solicitudes viejas, o el campo vino vacío), se resuelve
      // solo: primero buscando el correo en la pestaña Usuarios, y si no está ahí, armando el
      // nombre a partir del correo (nombre.apellido@bbva.com -> "Nombre Apellido").
      const nombreSolicitante = nombreCol || mapaUsuarios[correoSolicitante.toLowerCase()] || nombreDesdeCorreo_(correoSolicitante);
      return {
        fila: i + 2,
        fecha: iFecha !== -1 ? formatearFecha(fila[iFecha]) : '',
        tipo: iTipo !== -1 ? String(fila[iTipo] || '') : '',
        nombreProyecto: iProyecto !== -1 ? String(fila[iProyecto] || '') : '',
        estado: (iEstado !== -1 && fila[iEstado]) ? String(fila[iEstado]) : 'Solicitud Enviada',
        codigo: iCodigo !== -1 ? String(fila[iCodigo] || '') : '',
        nombreSolicitante: nombreSolicitante,
        correoSolicitante: correoSolicitante,
      };
    })
    .filter(s => s.nombreProyecto);
  lista.reverse();   // las últimas solicitudes primero
  return lista;
}

/** correo (en minúsculas) -> nombre, desde la pestaña Usuarios. */
function obtenerMapaCorreoNombre_() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const mapa = {};
  if (!hoja || hoja.getLastRow() < 2) return mapa;
  hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues().forEach(f => {
    const correo = String(f[0] || '').trim().toLowerCase();
    if (correo) mapa[correo] = String(f[1] || '').trim();
  });
  return mapa;
}

/** "nombre.apellido@bbva.com" -> "Nombre Apellido". Respaldo cuando no hay nombre guardado ni en Usuarios. */
function nombreDesdeCorreo_(correo) {
  if (!correo) return '';
  const local = String(correo).split('@')[0];
  if (!local) return '';
  return local.split(/[._\-]+/).filter(Boolean)
    .map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase())
    .join(' ');
}

/**
 * Avisa por correo al solicitante que su código de proyecto ya está asignado.
 * Se puede presionar cuando se quiera (no depende del Estado) — solo pide que la
 * columna Código ya tenga algo escrito y que haya un correo de solicitante registrado.
 */
function notificarCodigoAsignado(filaNumero) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Solicitudes');
  if (!hoja) return { ok: false, error: 'No existe la pestaña Solicitudes.' };

  const encabezadosCrudos = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const encabezados = encabezadosCrudos.map(h => String(h).trim().toLowerCase());
  const idx = nombre => encabezados.indexOf(String(nombre).trim().toLowerCase());
  const iProyecto = idx('Nombre del proyecto'), iCodigo = idx('Código');

  const fila = hoja.getRange(filaNumero, 1, 1, hoja.getLastColumn()).getValues()[0];
  const nombreProyecto = iProyecto !== -1 ? String(fila[iProyecto] || '') : '';
  const codigo = iCodigo !== -1 ? String(fila[iCodigo] || '') : '';
  const correo = hoja.getLastColumn() >= 13 ? String(fila[12] || '') : '';   // columna M
  const nombreSolicitante = hoja.getLastColumn() >= 7 ? String(fila[6] || '') : ''; // columna G

  if (!correo) return { ok: false, error: 'Esta solicitud no tiene el correo del solicitante registrado (columna M).' };
  if (!codigo) return { ok: false, error: 'Todavía no has puesto el Código en esta solicitud.' };

  MailApp.sendEmail({
    to: correo,
    subject: 'Tu código de proyecto ya está asignado — ' + (nombreProyecto || 'tu solicitud'),
    body: 'Hola' + (nombreSolicitante ? ' ' + nombreSolicitante : '') + ',\n\n' +
      'Tu solicitud de código para el proyecto "' + nombreProyecto + '" ya tiene código asignado:\n\n' +
      'Código: ' + codigo + '\n\n' +
      'Saludos.',
  });
  return { ok: true, correo: correo };
}

/**
 * Versión para el cliente: entrega las solicitudes como texto JSON en vez de
 * un array de objetos, para blindar la serialización google.script.run contra
 * cualquier problema con fechas u otros tipos de dato al cruzar el límite
 * cliente-servidor. El cliente hace JSON.parse() del resultado.
 */
function obtenerSolicitudes() {
  return JSON.stringify(obtenerSolicitudesArray_());
}

/**
 * Cuenta las solicitudes por estado, para el banner pequeño de la pestaña
 * Solicitud de Código.
 */
function obtenerResumenSolicitudes() {
  const lista = obtenerSolicitudesArray_();
  const resumen = { total: lista.length, enviada: 0, aprobada: 0, denegada: 0 };
  lista.forEach(s => {
    if (s.estado === 'Solicitud Aprobada') resumen.aprobada++;
    else if (s.estado === 'Solicitud Denegada') resumen.denegada++;
    else resumen.enviada++;
  });
  return resumen;
}
/**
 * DIAGNÓSTICO: muestra los encabezados reales de Solicitudes y cuántas filas
 * de datos tiene, para detectar por qué obtenerSolicitudes() no encuentra nada.
 */
function diagnosticoSolicitudes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Solicitudes');
  if (!hoja) return 'No existe la pestaña "Solicitudes".';
  const partes = [];
  partes.push('Última fila: ' + hoja.getLastRow() + ' | Última columna: ' + hoja.getLastColumn());
  if (hoja.getLastRow() >= 1) {
    partes.push('Encabezados (fila 1): ' + JSON.stringify(hoja.getRange(1,1,1,hoja.getLastColumn()).getValues()[0]));
  }
  if (hoja.getLastRow() >= 2) {
    partes.push('Fila 2 completa: ' + JSON.stringify(hoja.getRange(2,1,1,hoja.getLastColumn()).getValues()[0]));
  }
  partes.push('obtenerSolicitudes() devuelve: ' + obtenerSolicitudesArray_().length + ' solicitudes');
  return partes.join('\n');
}

/**
 * Actualiza el Estado o el Código de una solicitud puntual, por número de fila.
 * También lee por nombre de encabezado para ubicar la columna correcta.
 */
function actualizarSolicitud(filaNumero, campo, valor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Solicitudes');
  if (!hoja) return { ok: false, error: 'No existe la pestaña Solicitudes' };
  const nombreCol = campo === 'estado' ? 'Estado' : campo === 'codigo' ? 'Código' : null;
  if (!nombreCol) return { ok: false, error: 'Campo no reconocido: ' + campo };

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  let col = encabezados.indexOf(nombreCol) + 1;
  if (!col) {
    // La columna no existe todavía en esta hoja (versión anterior del formulario) — se crea al vuelo.
    col = hoja.getLastColumn() + 1;
    hoja.getRange(1, col).setValue(nombreCol);
  }
  hoja.getRange(filaNumero, col).setValue(valor);
  return { ok: true };
}

/**
 * Busca en BB_Cod_Pro coincidencias del nombre de proyecto (columna C) y
 * devuelve el código (columna A) de cada coincidencia. Hasta 20 resultados.
 */
function buscarCodigoProyecto(texto) {
  texto = String(texto || '').trim().toUpperCase();
  if (!texto) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('BB_Cod_Pro');
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getValues();
  const resultados = [];
  for (let i = 0; i < datos.length && resultados.length < 20; i++) {
    const nombre = String(datos[i][2] || '');
    if (nombre.toUpperCase().indexOf(texto) !== -1) {
      resultados.push({ codigo: datos[i][0], nombre: nombre });
    }
  }
  return resultados;
}

/**
 * Escribe un campo editable para una OC específica desde la app (dashboard / base de datos).
 * campo: 'proyectoEspecifico' | 'intervencion' | 'fechaPago' | 'comentarios'
 * Los comentarios se acumulan (no se sobrescriben) con fecha.
 */
function actualizarCampoOC(numeroOC, campo, valor) {
  // Si hay alerta de bloqueo activa, solo Maestro puede editar
  const alertas = obtenerAlertas();
  if (alertas.bloquearEdicion) {
    const u = obtenerSubRolVisitante();
    if (!u.esMaestro) return { ok: false, error: '🔒 El portal está en modo solo lectura. Contacta a un perfil Maestro.' };
  }

  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  const ultimaFila = hojaBase.getLastRow();
  const columnaOC = hojaBase.getRange(2, COL_BASE.OC, ultimaFila - 1, 1).getValues();

  // estadoVerificacion ahora es editable por todos los perfiles
  if (campo === 'estadoManual' && !esUsuarioAdmin()) {
    return { ok: false, error: 'Solo un usuario con rol admin puede cambiar el Estado.' };
  }
  if (campo === 'areaCorregida' && !esUsuarioAdmin()) {
    return { ok: false, error: 'Solo un usuario con rol admin puede cambiar el Área.' };
  }
  if (campo === 'estadoValidacion') {
    const u = obtenerSubRolVisitante();
    if (!u.esMaestro) return { ok: false, error: 'Solo un perfil Maestro puede cambiar la Validación.' };
  }

  // Cambiar una Fecha de pago que YA estaba pactada es sensible (mueve el Cash Flow).
  // Si quien edita no es admin, el cambio queda pendiente de aprobación y la fecha
  // real (la que usa el Cash Flow) no se toca hasta que un admin lo confirme.
  if (campo === 'fechaPago' && !esUsuarioAdmin()) {
    const colFecha = hojaBase.getRange(2, COL_BASE.FECHA_PAGO, ultimaFila - 1, 1).getValues();
    let fechaActual = '';
    for (let i = 0; i < columnaOC.length; i++) {
      if (String(columnaOC[i][0]) === String(numeroOC)) { fechaActual = colFecha[i][0]; break; }
    }
    const fechaActualTxt = formatearFecha(fechaActual);
    if (fechaActualTxt && fechaActualTxt !== valor) {
      registrarCambioFechaPendiente_(numeroOC, fechaActualTxt, valor);
      return { ok: true, pendiente: true, mensaje: 'La fecha ya estaba pactada — el cambio quedó pendiente de aprobación por un admin. El Cash Flow sigue usando la fecha anterior hasta que se apruebe.' };
    }
  }

  const mapaColumnas = {
    proyectoEspecifico: COL_BASE.PROYECTO_ESPECIFICO,
    intervencion: COL_BASE.INTERVENCION,
    fechaPago: COL_BASE.FECHA_PAGO,
    comentarios: COL_BASE.COMENTARIOS,
    fechaPagoValidada: COL_BASE.FECHA_PAGO_VALIDADA,
    areaCorregida: COL_BASE.AREA_CORREGIDA,
    estadoVerificacion: COL_BASE.ESTADO_VERIFICACION,
    estadoManual: COL_BASE.ESTADO_MANUAL,
    estadoValidacion: COL_BASE.ESTADO_VALIDACION,
  };
  const colDestino = mapaColumnas[campo];
  if (!colDestino) return { ok: false, error: 'Campo no reconocido: ' + campo };

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let filasActualizadas = 0;

  for (let i = 0; i < columnaOC.length; i++) {
    if (String(columnaOC[i][0]) !== String(numeroOC)) continue;
    const filaReal = i + 2;
    if (campo === 'comentarios') {
      const celda = hojaBase.getRange(filaReal, colDestino);
      const previo = celda.getValue();
      celda.setValue(previo ? `${previo}\n[${fecha}] ${valor}` : `[${fecha}] ${valor}`);
    } else {
      hojaBase.getRange(filaReal, colDestino).setValue(valor);
    }
    filasActualizadas++;
  }

  if (filasActualizadas === 0) return { ok: false, error: 'OC no encontrada en BASE_CASHFLOW' };
  bumpCacheVer_();
  return { ok: true, filasActualizadas: filasActualizadas };
}

/** Registra un cambio de Fecha de pago que quedó pendiente de aprobación. */
function registrarCambioFechaPendiente_(numeroOC, fechaAnterior, fechaNueva) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('HistorialFechas');
  if (!hoja) {
    hoja = ss.insertSheet('HistorialFechas');
    hoja.getRange(1, 1, 1, 6).setValues([['OC', 'Fecha anterior', 'Fecha nueva', 'Usuario', 'Fecha del cambio', 'Estado']]);
  }
  const usuario = obtenerCorreoVisitante_();
  hoja.appendRow([numeroOC, fechaAnterior, fechaNueva, usuario, new Date(), 'Pendiente']);
}

/**
 * Lista los cambios de fecha pendientes de aprobar, con el monto de la OC
 * (pendiente de OC) para poder mostrar el impacto de aprobarlo o no.
 */
function leerCambiosFechaPendientesRaw_() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HistorialFechas');
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 6).getValues()
    .map((fila, i) => ({
      fila: i + 2,
      oc: fila[0],
      fechaAnterior: formatearFecha(fila[1]) || fila[1],
      fechaNueva: formatearFecha(fila[2]) || fila[2],
      usuario: fila[3],
      fechaCambio: formatearFecha(fila[4]),
      estado: fila[5] || 'Pendiente',
    }))
    .filter(c => c.oc !== '' && c.estado === 'Pendiente');
}
function mapearCambiosFecha_(raw, info) {
  return raw.map(c => Object.assign(c, { monto: (info[String(c.oc)] || {}).saldo || 0, proyecto: (info[String(c.oc)] || {}).proyecto || '' }));
}
function obtenerCambiosFechaPendientes() {
  const raw = leerCambiosFechaPendientesRaw_();
  if (!raw.length) return [];
  const set = {};
  raw.forEach(c => { set[String(c.oc)] = true; });
  return mapearCambiosFecha_(raw, infoLigeraOCs_(set));      // solo 4 columnas de la base, no toda la hoja
}

/** El admin aprueba un cambio de fecha: recién ahí se actualiza la fecha real en BASE_CASHFLOW. */
function aprobarCambioFecha(filaHistorial) {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede aprobar cambios de fecha.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('HistorialFechas');
  if (!hoja) return { ok: false, error: 'No existe HistorialFechas' };
  const fila = hoja.getRange(filaHistorial, 1, 1, 6).getValues()[0];
  const oc = fila[0], fechaNueva = fila[2];

  const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
  const ultimaFila = hojaBase.getLastRow();
  const columnaOC = hojaBase.getRange(2, COL_BASE.OC, ultimaFila - 1, 1).getValues();
  let actualizadas = 0;
  for (let i = 0; i < columnaOC.length; i++) {
    if (String(columnaOC[i][0]) !== String(oc)) continue;
    hojaBase.getRange(i + 2, COL_BASE.FECHA_PAGO).setValue(fechaNueva);
    actualizadas++;
  }
  hoja.getRange(filaHistorial, 6).setValue('Aprobado');
  bumpCacheVer_();
  return { ok: true, filasActualizadas: actualizadas };
}

/** El admin rechaza un cambio de fecha: la fecha original queda intacta. */
function rechazarCambioFecha(filaHistorial) {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede rechazar cambios de fecha.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('HistorialFechas');
  if (!hoja) return { ok: false, error: 'No existe HistorialFechas' };
  hoja.getRange(filaHistorial, 6).setValue('Rechazado');
  return { ok: true };
}

/**
 * Crea (si no existe) la pestaña TechosYFondeos con la estructura acordada y
 * la deja precargada con los valores que veníamos usando como referencia
 * manual, para que desde ahí se puedan editar/actualizar sin tocar código.
 * Columnas: Tipo (Techo/Fondeo/Barrido) | Dominio | Proyecto | Monto | Fecha | Comentario
 */
/**
 * Estructura de "TechosYFondeos" (confirmada con el usuario):
 *   A = Sección: '1', '2' ó '3' para las líneas de cada sección; '2-Velocimetro' y '3-Velocimetro'
 *       para las filas que alimentan el cálculo del velocímetro de esas secciones.
 *   B = Título (agrupa las líneas; en las filas "-Velocimetro" es el nombre exacto usado en la fórmula).
 *   C = Dominio (P&S / Tecnología / Seguridad) — filtro.
 *   D = Plazo (Plurianual / Anual) — filtro.
 *   E = Vista/Tipo (por ejemplo SOLPED+IGV / OC) — filtro.
 *   F = Trimestre — filtro.
 *   G = Monto, en millones (se multiplica x1,000,000 al leerlo).
 */
function crearTechosYFondeosSiFalta_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('TechosYFondeos');
  if (hoja) return hoja;

  hoja = ss.insertSheet('TechosYFondeos');
  hoja.getRange(1, 1, 1, 7).setValues([['Sección', 'Título', 'Dominio', 'Plazo', 'Vista', 'Trimestre', 'Monto (MM)']]);

  const filasSemilla = [
    // Sección 1 — líneas por dominio (alimentan el anillo P&S / Tecnología / Seguridad). Columna G en SOLES.
    ['1', 'Capex Anual Autorizado', 'P&S', '', '', '', 67062141],
    ['1', 'Capex Anual Autorizado', 'Tecnología', '', '', '', 7383430],
    ['1', 'Capex Anual Autorizado', 'Seguridad', '', '', '', 4484990],
    ['1', 'Capex Fondeado', 'P&S', '', '', '', 64051991],
    ['1', 'Capex Fondeado', 'Tecnología', '', '', '', 6848435],
    ['1', 'Capex Fondeado', 'Seguridad', '', '', '', 4276702],
    // Sección 2 — líneas (ejemplo)
    ['2', 'Monto SOLPED+IGV', 'P&S', '', 'SOLPED+IGV', '', 20000000],
    ['2', 'Monto OC pendiente de pago', 'P&S', '', 'OC', '', 15000000],
    ['2', 'Monto OC Pagadas', 'P&S', '', 'OC', '', 10000000],
    // Sección 2 — velocímetro: % = Capex Total OC / Capex Comprometido Total
    ['2-Velocimetro', 'Capex Total OC', 'P&S', '', '', '', 25000000],
    ['2-Velocimetro', 'Capex Comprometido Total', 'P&S', '', '', '', 67060000],
    // Sección 3 — líneas (ejemplo)
    ['3', 'Capex Pagado', 'P&S', '', '', '', 10000000],
    ['3', 'Capex Total', 'P&S', '', '', '', 67060000],
    // Sección 3 — velocímetro: % = Capex Pagado / Capex Total
    ['3-Velocimetro', 'Capex Pagado', 'P&S', '', '', '', 10000000],
    ['3-Velocimetro', 'Capex Total', 'P&S', '', '', '', 67060000],
  ];
  hoja.getRange(2, 1, filasSemilla.length, 7).setValues(filasSemilla);
  return hoja;
}

/**
 * Lee TechosYFondeos y devuelve los montos de Techo (Autorizado) y Fondeo por
 * Dominio, sumando por si hay varias filas del mismo Tipo+Dominio (por ejemplo,
 * si ya empezaste a itemizar por Proyecto). Si la pestaña no existe, la crea
 * con los valores semilla la primera vez que se llama.
 */
function obtenerDistribucionCapex() {
  const hoja = crearTechosYFondeosSiFalta_();
  const ultimaFila = hoja.getLastRow();
  const autorizadoPorDominio = { 'P&S': 0, 'Tecnología': 0, 'Seguridad': 0 };
  const fondeadoPorDominio = { 'P&S': 0, 'Tecnología': 0, 'Seguridad': 0 };
  const barridoPorDominio = { 'P&S': 0, 'Tecnología': 0, 'Seguridad': 0 };

  if (ultimaFila >= 2) {
    // A=Sección(0), B=Título(1), C=Dominio(2), D=(sin uso), E=Monto en MM(4)
    const datos = hoja.getRange(2, 1, ultimaFila - 1, 5).getValues();
    datos.forEach(fila => {
      const seccionTexto = String(fila[0] || '').trim();
      const matchNumero = seccionTexto.match(/\d+/);
      const seccion = matchNumero ? matchNumero[0] : seccionTexto;
      if (seccion !== '1') return; // Sección 1 = Techo/Fondeo/Solicitado por dominio
      const titulo = String(fila[1] || '').trim().toLowerCase();
      const dominio = String(fila[2] || '').trim();
      const montoMM = Number(fila[4]) || 0;
      const monto = montoMM * 1000000; // el Sheet lo expresa en millones

      if (!autorizadoPorDominio.hasOwnProperty(dominio)) return;
      if (titulo.includes('autorizad')) autorizadoPorDominio[dominio] += monto;
      else if (titulo.includes('fondead')) fondeadoPorDominio[dominio] += monto;
      else if (titulo.includes('barrid')) barridoPorDominio[dominio] += monto;
    });
  }

  return { autorizadoPorDominio, fondeadoPorDominio, barridoPorDominio };
}

/**
 * Cuenta las OCs "Por pagar" (con saldo pendiente de OC > 0), agrupadas por
 * Trimestre y también por Trimestre+Proyecto, para la sección de cuadritos
 * de Vista General.
 */
function obtenerOCPorPagarPorTrimestreYProyecto(filtros) {
  filtros = filtros || {};
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return { porTrimestre: {}, porProyecto: [] };

  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  const porTrimestre = {};
  const claveProyTrim = {}; // { 'Proyecto|Q1': count }

  agrupadas.forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (g.pendienteOC <= 0) return; // solo "Por pagar"
    if (['1Q','2Q','3Q'].indexOf(g.trimestre) === -1) return;

    porTrimestre[g.trimestre] = (porTrimestre[g.trimestre] || 0) + 1;
    const clave = (g.proyecto || 'Sin proyecto') + '|' + g.trimestre;
    claveProyTrim[clave] = (claveProyTrim[clave] || 0) + 1;
  });

  const porProyectoSet = {};
  Object.keys(claveProyTrim).forEach(clave => {
    const proyecto = clave.split('|')[0];
    porProyectoSet[proyecto] = true;
  });

  const porProyecto = Object.keys(porProyectoSet).sort().map(proyecto => ({
    proyecto: proyecto,
    q1: claveProyTrim[proyecto + '|1Q'] || 0,
    q2: claveProyTrim[proyecto + '|2Q'] || 0,
    q3: claveProyTrim[proyecto + '|3Q'] || 0,
  }));

  return { porTrimestre, porProyecto };
}

/**
 * Resumen para "Órdenes de Pago Pendientes": cantidad de OCs y su monto,
 * separado en SOLPED+IGV (todavía no es OC), Órdenes (ya es OC, pendiente de
 * pago) y Total. También el avance de registro de Fecha de pago (solo de las
 * OCs que ya tienen saldo de OC pendiente).
 */
function obtenerResumenOrdenesPendientes(filtros, datosPre) {
  filtros = filtros || {};
  let datos = datosPre;
  if (!datos) {
    const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
    if (!hojaBase || hojaBase.getLastRow() < 2) {
      return { countSolped:0, countOC:0, countTotal:0, montoSolped:0, montoOC:0, montoTotal:0, ocConFecha:0, ocSinFecha:0 };
    }
    datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  }
  const agrupadas = agruparPorOC_(datos);

  let countSolped=0, countOC=0, montoSolped=0, montoOC=0, ocConFecha=0, ocSinFecha=0;

  agrupadas.forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.dominioSp, (g.dominioSp||'').trim() || '(Sin dato)')) return;
    if (g.solped > 0) { countSolped++; montoSolped += g.solped; }
    if (g.pendienteOC > 0) {
      countOC++; montoOC += g.pendienteOC;
      if (g.fechaPago) ocConFecha++; else ocSinFecha++;
    }
  });

  return {
    countSolped, countOC, countTotal: countSolped + countOC,
    montoSolped, montoOC, montoTotal: montoSolped + montoOC,
    ocConFecha, ocSinFecha,
  };
}

/** Lista de usuarios autorizados (correo + nombre), para elegir destinatarios de un correo. */
function obtenerListaUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Usuarios');
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
  return datos
    .filter(f => f[0])
    .map(f => ({ correo: f[0], nombre: f[1] || f[0] }));
}

/** Envía un correo con asunto y mensaje libres a los destinatarios elegidos. */
function enviarCorreoActividades(destinatarios, asunto, mensaje) {
  if (!destinatarios || !destinatarios.length) return { ok: false, error: 'No se eligió ningún destinatario.' };
  const remitente = obtenerCorreoVisitante_();
  MailApp.sendEmail({
    to: destinatarios.join(','),
    subject: asunto || 'Actividad — Control Presupuestal',
    body: mensaje || '',
    cc: remitente || undefined,
  });
  return { ok: true };
}

/**
 * Exporta el Cash Flow (Calendario de pagos por SDATOOL y Proyecto) a un
 * archivo .xlsx real: arma una hoja de cálculo temporal con los datos ya
 * filtrados, la exporta como Excel, y la borra. Devuelve el archivo en base64
 * para que el navegador lo descargue.
 */
function exportarCashFlowExcel(filtros) {
  const datosGantt = obtenerGantt(filtros || {});
  const ssTemp = SpreadsheetApp.create('CashFlow_export_temp_' + new Date().getTime());
  const hoja = ssTemp.getSheets()[0];

  const encabezados = ['SDATOOL', 'Proyecto', ...datosGantt.semanas.map(s => s.etiqueta), 'Total'];
  hoja.appendRow(encabezados);
  datosGantt.filas.forEach(f => {
    const total = f.montos.reduce((a, b) => a + b, 0);
    hoja.appendRow([f.sdatool, f.proyecto, ...f.montos, total]);
  });
  hoja.getRange(1, 1, 1, encabezados.length).setFontWeight('bold');
  SpreadsheetApp.flush();

  let base64 = '';
  try {
    const url = 'https://docs.google.com/spreadsheets/d/' + ssTemp.getId() + '/export?format=xlsx';
    const token = ScriptApp.getOAuthToken();
    const respuesta = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    base64 = Utilities.base64Encode(respuesta.getContent());
  } finally {
    DriveApp.getFileById(ssTemp.getId()).setTrashed(true);
  }

  const nombreArchivo = 'CashFlow_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.xlsx';
  return { ok: true, base64: base64, filename: nombreArchivo };
}

/* ======================================================================
   DATA SEMANAL — hoja de datos "duros" que se pega a mano cada semana.
   Todo esto trabaja buscando columnas POR NOMBRE de encabezado (no por
   letra), para no depender de que la posición exacta sea siempre la misma.
   ====================================================================== */

const CONFIG_SEMANAL = {
  HOJA_DATOS: 'Rep.Capex',            // hoja cruda que pegas cada semana
  // Nombres posibles de la hoja con la lista maestra de plurianuales (nombres de proyecto y/o Elementos PEP).
  // Se leen y combinan TODAS las que existan; si no existe ninguna, se crea la primera con la semilla.
  HOJAS_PLURIANUALES: ['ProyPluri (DATA)', 'ProyectosPlurianuales'],
  HOJA_RESUMEN: 'Rep.Capex (Rev)',    // hoja procesada que alimenta Vista General
};

const PROYECTOS_PLURIANUALES_SEMILLA = [
  'Renovación de Ascensores',
  'Renov. de equipos CPD',
  'Renovación de luces de fachada',
  'Banquero remoto',
  'Implementación pantalla gran Hall',
  'Optimización de Áreas piso 1',
  'Redimensionamiento Of. Iquitos',
  'Reforzamiento Club Chaclacayo',
  // Elementos PEP plurianuales (la lista puede tener nombres de proyecto O códigos PEP; se compara contra ambas columnas)
  'GB.00175390-004', 'GB.00179819-002', 'GB.00179819-004', 'GB.00180312-003', 'GB.00180312-004',
  'GB.00182595-002', 'GB.00182622-002', 'GB.00185236-002', 'GB.00185236-004', 'GB.00185402-002',
  'GB.00185402-005', 'GB.00185402-006',
];

/**
 * Lee la lista maestra de plurianuales (nombres de proyecto y/o Elementos PEP), combinando
 * TODAS las hojas de CONFIG_SEMANAL.HOJAS_PLURIANUALES que existan (para no perder nada si hay
 * más de una, p.ej. "ProyPluri (DATA)" y la antigua "ProyectosPlurianuales"). Si ninguna existe,
 * crea la primera de la lista con la semilla de ejemplo.
 */
function leerPlurianualesMaestros_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombres = CONFIG_SEMANAL.HOJAS_PLURIANUALES;
  const hojasExistentes = nombres.map(n => ss.getSheetByName(n)).filter(h => h);

  if (!hojasExistentes.length) {
    const hoja = ss.insertSheet(nombres[0]);
    hoja.getRange(1, 1, 1, 1).setValues([['Proyecto o Elemento PEP (Plurianual)']]);
    hoja.getRange(2, 1, PROYECTOS_PLURIANUALES_SEMILLA.length, 1)
      .setValues(PROYECTOS_PLURIANUALES_SEMILLA.map(p => [p]));
    hojasExistentes.push(hoja);
  }

  const plurianuales = {};
  hojasExistentes.forEach(hoja => {
    const ultima = hoja.getLastRow();
    if (ultima < 2) return;
    hoja.getRange(2, 1, ultima - 1, 1).getValues()
      .forEach(f => { if (f[0]) plurianuales[String(f[0]).trim().toLowerCase()] = true; });
  });
  return plurianuales;
}

/** Busca el índice (0-based) de la primera columna cuyo encabezado contiene el texto dado. */
function buscarColumnaPorNombre_(encabezados, textoBuscado) {
  const idx = encabezados.findIndex(h => String(h).toLowerCase().trim().includes(textoBuscado.toLowerCase()));
  return idx; // -1 si no se encontró
}

/**
 * Procesa la hoja DataSemanal recién pegada:
 * 1) Inserta "Solped* (incluye IGV)" al lado de Solped = Solped * 1.18
 * 2) Recalcula Imp. Disponible = Autorizado - Solped* - Comp.(Pedido) - Imp.Realizado
 * 3) Llena "Barrido" = Imp.Disponible si Trimestre es 1Q/2Q y ese disponible > 1000
 * 4) Llena "Tipo de Proyecto" (Plurianual/Anual) cruzando con la lista maestra
 * 5) Reconstruye la pestaña ResumenProyectos (una fila por Proyecto+Trimestre)
 * Se puede correr manualmente desde el menú "Control Presupuestal" del Sheet,
 * o llamando a esta función desde el editor de Apps Script.
 */
function procesarDataSemanal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG_SEMANAL.HOJA_DATOS);
  if (!hoja) {
    throw new Error('No encontre la pestana "' + CONFIG_SEMANAL.HOJA_DATOS + '". Creala y pega ahi tu data cruda.');
  }
  const ultimaFila = hoja.getLastRow();
  const ultimaCol = hoja.getLastColumn();
  if (ultimaFila < 2) throw new Error('La pestana "' + CONFIG_SEMANAL.HOJA_DATOS + '" no tiene datos todavia.');

  // Columnas por letra (1-based), segun el mapa confirmado:
  // E=Autorizado(5), F=Solped(6), H=Comp.Pedido(8 tras insertar G), I=Imp.Realizado(9), J=Disponible(10)
  const COL_AUTORIZADO = 5;
  const COL_SOLPED = 6;

  const encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
  const datosCrudos = hoja.getRange(2, 1, ultimaFila - 1, ultimaCol).getValues();

  // Ubicar por nombre las columnas que necesitamos para clasificar (no dependen de letra fija)
  const iProyecto = buscarColumnaPorNombre_(encabezados, 'proyecto');
  const iTrimestre = buscarColumnaPorNombre_(encabezados, 'trimestre');
  const iPEP = buscarColumnaPorNombre_(encabezados, 'pep');       // "Elemento PEP"
  const iDenominacion = buscarColumnaPorNombre_(encabezados, 'denominac');
  const iDominio = buscarColumnaPorNombre_(encabezados, 'dominio');

  // Lista maestra de plurianuales (combina todas las hojas de CONFIG_SEMANAL.HOJAS_PLURIANUALES)
  const plurianuales = leerPlurianualesMaestros_();

  // Construir los encabezados de la hoja procesada:
  // se inserta "Solped* (incluye IGV)" despues de F, y se agregan 2 columnas al final
  const encabezadosRev = encabezados.slice(0, COL_SOLPED)
    .concat(['Solped* (incluye IGV)'])
    .concat(encabezados.slice(COL_SOLPED))
    .concat(['Tipo de Proyecto', 'Alerta']);

  const filasRev = datosCrudos.map(filaCruda => {
    const autorizado = Number(filaCruda[COL_AUTORIZADO - 1]) || 0;
    const solped = Number(filaCruda[COL_SOLPED - 1]) || 0;
    const solpedIGV = Math.round(solped * 1.18 * 100) / 100;

    // Reconstruir la fila con la columna nueva insertada en la posicion G
    const fila = filaCruda.slice(0, COL_SOLPED)
      .concat([solpedIGV])
      .concat(filaCruda.slice(COL_SOLPED));

    // Ahora en la fila nueva: G=solpedIGV(7), H=Comp.Pedido(8), I=Imp.Realizado(9), J=Disponible(10)
    const compPedido = Number(fila[7]) || 0;
    const impRealizado = Number(fila[8]) || 0;
    const disponible = autorizado - solpedIGV - compPedido - impRealizado;
    fila[9] = disponible;

    // Tipo de proyecto: plurianual si el NOMBRE o el ELEMENTO PEP está en la lista maestra; si no, anual
    const proyecto = iProyecto !== -1 ? String(filaCruda[iProyecto] || '').trim() : '';
    const pep = iPEP !== -1 ? String(filaCruda[iPEP] || '').trim() : '';
    const tipoProyecto = (plurianuales[proyecto.toLowerCase()] || (pep && plurianuales[pep.toLowerCase()])) ? 'Plurianual' : 'Anual';

    // Alerta: 3Q y 4Q sin alerta; 1Q y 2Q con "Barrido" si el disponible es mayor a cero
    const trimestre = iTrimestre !== -1 ? String(filaCruda[iTrimestre] || '').trim() : '';
    let alerta = '';
    if ((trimestre === '1Q' || trimestre === '2Q') && disponible > 0) alerta = 'Barrido';

    return fila.concat([tipoProyecto, alerta]);
  });

  // Escribir la hoja procesada desde cero
  let hojaRev = ss.getSheetByName(CONFIG_SEMANAL.HOJA_RESUMEN);
  if (!hojaRev) hojaRev = ss.insertSheet(CONFIG_SEMANAL.HOJA_RESUMEN);
  hojaRev.clear();
  hojaRev.getRange(1, 1, 1, encabezadosRev.length).setValues([encabezadosRev]).setFontWeight('bold');
  if (filasRev.length) {
    hojaRev.getRange(2, 1, filasRev.length, encabezadosRev.length).setValues(filasRev);
  }
  hojaRev.setFrozenRows(1);

  try { generarAlertaBarrido(); } catch (eBarrido) { Logger.log('No se pudo actualizar "Alerta de Barrido": ' + eBarrido); }

  return { ok: true, filas: filasRev.length };
}

/**
 * Reconstruye la pestaña "Alerta de Barrido": todas las filas de "Rep.Capex (Rev)" con Alerta = "Barrido",
 * agrupadas por Proyecto (con subtotal por proyecto), mostrando Denominación, Elemento PEP, el importe
 * disponible a barrer, Trimestre, Dominio y Tipo de Proyecto.
 */
/** Calcula las líneas con Alerta = "Barrido" desde "Rep.Capex (Rev)", ordenadas por proyecto. Función pura (no escribe nada). */
function calcularAlertaBarrido_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRev = ss.getSheetByName(CONFIG_SEMANAL.HOJA_RESUMEN);
  if (!hojaRev || hojaRev.getLastRow() < 2) throw new Error('No hay datos en "' + CONFIG_SEMANAL.HOJA_RESUMEN + '". Corre primero "Procesar data semanal".');

  const nCol = hojaRev.getLastColumn();
  const encabezados = hojaRev.getRange(1, 1, 1, nCol).getValues()[0];
  const datos = hojaRev.getRange(2, 1, hojaRev.getLastRow() - 1, nCol).getValues();
  const idx = nombre => buscarColumnaPorNombre_(encabezados, nombre);
  const iProyecto = idx('proyecto'), iDenom = idx('denominac'), iPEP = idx('pep'),
        iDisponible = idx('disponible'), iTrimestre = idx('trimestre'), iDominio = idx('dominio'),
        iTipoProyecto = idx('tipo de proyecto'), iAlerta = idx('alerta'), iDominioSp = idx('dominio sp');
  if (iAlerta === -1) throw new Error('No encuentro la columna "Alerta" en "' + CONFIG_SEMANAL.HOJA_RESUMEN + '".');
  const DOMINIOS_VALIDOS = ['p&s', 'tecnologia', 'tecnología', 'seguridad'];

  const filas = datos
    .filter(f => String(f[iAlerta] || '').trim() === 'Barrido')
    .filter(f => iDominioSp === -1 || String(f[iDominioSp] || '').trim().toUpperCase() === 'T&C')
    .filter(f => iDominio === -1 || DOMINIOS_VALIDOS.indexOf(String(f[iDominio] || '').trim().toLowerCase()) !== -1)
    .map(f => ({
      proyecto: iProyecto !== -1 ? String(f[iProyecto] || '') : '',
      denominacion: iDenom !== -1 ? f[iDenom] : '',
      pep: iPEP !== -1 ? f[iPEP] : '',
      disponible: iDisponible !== -1 ? (Number(f[iDisponible]) || 0) : 0,
      trimestre: iTrimestre !== -1 ? f[iTrimestre] : '',
      dominio: iDominio !== -1 ? f[iDominio] : '',
      tipoProyecto: iTipoProyecto !== -1 ? f[iTipoProyecto] : '',
    }));
  filas.sort((a, b) => a.proyecto.localeCompare(b.proyecto) || String(a.trimestre).localeCompare(String(b.trimestre)));
  return filas;
}

/** Para el portal: las líneas de Barrido, agrupadas por proyecto con su subtotal, y el total general. */
function obtenerAlertaBarrido() {
  try {
    const filas = calcularAlertaBarrido_();
    const grupos = [];
    let actual = null, totalGeneral = 0;
    filas.forEach(f => {
      if (!actual || actual.proyecto !== f.proyecto) { actual = { proyecto: f.proyecto, subtotal: 0, lineas: [] }; grupos.push(actual); }
      actual.subtotal += f.disponible; totalGeneral += f.disponible;
      actual.lineas.push(f);
    });
    return { ok: true, grupos: grupos, totalGeneral: totalGeneral, totalLineas: filas.length };
  } catch (e) {
    return { ok: false, error: e.message, grupos: [], totalGeneral: 0, totalLineas: 0 };
  }
}

function generarAlertaBarrido() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filas = calcularAlertaBarrido_();
  let hoja = ss.getSheetByName('Alerta de Barrido');
  if (!hoja) hoja = ss.insertSheet('Alerta de Barrido'); else hoja.clear();
  const encab = ['Proyecto', 'Denominación', 'Elemento PEP', 'Importe disponible a barrer', 'Trimestre', 'Dominio', 'Tipo de Proyecto'];
  const salida = [encab];
  const filasNegrita = [1];
  let proyectoActual = null, subtotal = 0, totalGeneral = 0;
  filas.forEach(f => {
    if (proyectoActual !== null && f.proyecto !== proyectoActual) {
      salida.push(['Subtotal ' + proyectoActual, '', '', subtotal, '', '', '']);
      filasNegrita.push(salida.length);
      subtotal = 0;
    }
    proyectoActual = f.proyecto;
    subtotal += f.disponible; totalGeneral += f.disponible;
    salida.push([f.proyecto, f.denominacion, f.pep, f.disponible, f.trimestre, f.dominio, f.tipoProyecto]);
  });
  if (proyectoActual !== null) { salida.push(['Subtotal ' + proyectoActual, '', '', subtotal, '', '', '']); filasNegrita.push(salida.length); }
  salida.push(['TOTAL GENERAL', '', '', totalGeneral, '', '', '']);
  filasNegrita.push(salida.length);

  if (hoja.getMaxColumns() < encab.length) hoja.insertColumnsAfter(hoja.getMaxColumns(), encab.length - hoja.getMaxColumns());
  if (hoja.getMaxRows() < salida.length) hoja.insertRowsAfter(Math.max(hoja.getMaxRows(),1), salida.length - hoja.getMaxRows());
  hoja.getRange(1, 1, salida.length, encab.length).setValues(salida);
  hoja.getRange(1, 4, salida.length, 1).setNumberFormat('#,##0.00');
  filasNegrita.forEach(r => hoja.getRange(r, 1, 1, encab.length).setFontWeight('bold'));
  hoja.setFrozenRows(1);
  try { hoja.autoResizeColumns(1, encab.length); } catch (e) {}
  return { ok: true, filas: filas.length };
}
function generarAlertaBarridoConAviso() {
  try {
    const r = generarAlertaBarrido();
    SpreadsheetApp.getUi().alert('Listo — ' + r.filas + ' líneas con alerta de barrido, agrupadas por proyecto.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Control Presupuestal')
    .addItem('Procesar data semanal', 'procesarDataSemanalConAviso')
    .addItem('Actualizar Alerta de Barrido', 'generarAlertaBarridoConAviso')
    .addItem('Diagnóstico Rep. Capex', 'diagnosticoRepCapexConAviso')
    .addItem('Diagnóstico TechosYFondeos', 'diagnosticoTechosYFondeosConAviso')
    .addItem('Diagnóstico Fondeo', 'diagnosticoFondeoConAviso')
    .addItem('Diagnóstico columna Validación', 'diagnosticoColumnaValidacion')
    .addSeparator()
    .addItem('Instalar correo semanal (lunes 9AM)', 'instalarTriggerCorreoSemanal')
    .addToUi();
}

function procesarDataSemanalConAviso() {
  try {
    const r = procesarDataSemanal();
    SpreadsheetApp.getUi().alert('Listo — se procesaron ' + r.filas + ' filas en "' + CONFIG_SEMANAL.HOJA_RESUMEN + '".');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message + '\n\nDetalle técnico:\n' + (e.stack || '(sin más detalle)'));
  }
}

/**
 * DIAGNÓSTICO paso a paso de "Rep. Capex" — corre esto (desde el editor de
 * Apps Script, botón ▶ con esta función seleccionada, o agrégala al menú)
 * para ver EXACTAMENTE en qué paso falla, en vez de un error genérico.
 */
function devolverYRegistrar_(partes) {
  const resultado = partes.join('\n');
  Logger.log(resultado);
  return resultado;
}

function diagnosticoRepCapex() {
  const partes = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    partes.push('✓ Spreadsheet activo: ' + ss.getName());

    const hoja = ss.getSheetByName(CONFIG_SEMANAL.HOJA_DATOS);
    if (!hoja) {
      partes.push('✗ NO existe la pestaña "' + CONFIG_SEMANAL.HOJA_DATOS + '". Pestañas que sí existen: ' + ss.getSheets().map(h=>h.getName()).join(', '));
      return devolverYRegistrar_(partes);
    }
    partes.push('✓ Existe la pestaña "' + CONFIG_SEMANAL.HOJA_DATOS + '"');

    const ultimaFila = hoja.getLastRow();
    const ultimaCol = hoja.getLastColumn();
    partes.push('Última fila: ' + ultimaFila + ' | Última columna: ' + ultimaCol);
    if (ultimaFila < 2) { partes.push('✗ No hay filas de datos (solo encabezado o vacío)'); return devolverYRegistrar_(partes); }

    const encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
    partes.push('Encabezados (fila 1): ' + JSON.stringify(encabezados));

    const fila2 = hoja.getRange(2, 1, 1, ultimaCol).getValues()[0];
    partes.push('Fila 2 completa: ' + JSON.stringify(fila2));

    partes.push('Columna E (Autorizado, fila 2): ' + fila2[4]);
    partes.push('Columna F (Solped, fila 2): ' + fila2[5]);
    partes.push('Columna H (Comp.Pedido tras insertar G, fila 2 ANTES de insertar): ' + fila2[6]);

    const iProyecto = buscarColumnaPorNombre_(encabezados, 'proyecto');
    const iTrimestre = buscarColumnaPorNombre_(encabezados, 'trimestre');
    partes.push('Columna "proyecto" encontrada por nombre en índice: ' + iProyecto + (iProyecto===-1?' (¡NO ENCONTRADA!)':' → "'+encabezados[iProyecto]+'"'));
    partes.push('Columna "trimestre" encontrada por nombre en índice: ' + iTrimestre + (iTrimestre===-1?' (¡NO ENCONTRADA!)':' → "'+encabezados[iTrimestre]+'"'));

    partes.push('\nIntentando correr procesarDataSemanal() completo...');
    const r = procesarDataSemanal();
    partes.push('✓ ÉXITO — se procesaron ' + r.filas + ' filas en "' + CONFIG_SEMANAL.HOJA_RESUMEN + '"');
  } catch (e) {
    partes.push('✗ ERROR: ' + e.message);
    partes.push('Detalle técnico: ' + (e.stack || '(sin más detalle)'));
  }
  return devolverYRegistrar_(partes);
}

/** Igual que diagnosticoRepCapex pero mostrado en una alerta del Sheet. */
function diagnosticoRepCapexConAviso() {
  SpreadsheetApp.getUi().alert(diagnosticoRepCapex());
}

/**
 * Lee la pestaña "Rep. Capex (Rev)" (generada por procesarDataSemanal) para la
 * pantalla "Detalle de Proyectos" del portal. Devuelve encabezados + filas tal
 * cual están en la hoja, para no depender de un orden fijo de columnas.
 */
function obtenerDetalleProyectos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG_SEMANAL.HOJA_RESUMEN);
  if (!hoja || hoja.getLastRow() < 2) return { encabezados: [], filas: [] };
  const ultimaCol = hoja.getLastColumn();
  const encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0].map(h => String(h));
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, ultimaCol).getValues()
    .map(fila => fila.map(celda => {
      if (Object.prototype.toString.call(celda) === '[object Date]') return formatearFecha(celda);
      return celda;
    }));
  return { encabezados: encabezados, filas: filas };
}

/**
 * Indicadores por Trimestre desde "Rep. Capex (Rev)": Fondeado (Autorizado) y
 * Pendiente de pago (Disponible), para el banner de Vista General.
 */
function obtenerIndicadoresQRepCapex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG_SEMANAL.HOJA_RESUMEN);
  if (!hoja || hoja.getLastRow() < 2) return {};
  const ultimaCol = hoja.getLastColumn();
  const encabezados = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
  const iTrimestre = buscarColumnaPorNombre_(encabezados, 'trimestre');
  if (iTrimestre === -1) return {};

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, ultimaCol).getValues();
  const porQ = {};
  datos.forEach(fila => {
    const q = String(fila[iTrimestre] || '').trim();
    if (!q) return;
    if (!porQ[q]) porQ[q] = { fondeado: 0, pendientePago: 0 };
    porQ[q].fondeado += Number(fila[4]) || 0;      // E = Autorizado
    porQ[q].pendientePago += Number(fila[9]) || 0; // J = Disponible
  });
  return porQ;
}

/**
 * Cuánto falta por agendar de pagar: OCs con saldo pendiente (no pagadas)
 * que todavía no tienen Fecha de pago propuesta cargada.
 */
function obtenerPendientePorAgendar(filtros, datosPre) {
  filtros = filtros || {};
  let datos = datosPre;
  if (!datos) {
    const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
    if (!hojaBase || hojaBase.getLastRow() < 2) return { monto: 0, cantidad: 0, montoAgendado: 0, cantidadAgendada: 0 };
    datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  }
  const agrupadas = agruparPorOC_(datos);

  let monto = 0, cantidad = 0, montoAgendado = 0, cantidadAgendada = 0;
  agrupadas.forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (!coincideFiltro_(filtros.dominioSp, (g.dominioSp||'').trim() || '(Sin dato)')) return;
    if (g.pendienteOC <= 0) return; // solo OCs que aún no están pagadas

    // "Agendado" = tiene una fecha de pago propuesta O está valorizada en cuotas
    // (misma condición que usa el Cash Flow para incluir el monto en la proyección)
    const tieneAgenda = estaAgendada_(g);   // misma regla que el panel "Órdenes sin agendar"
    if (tieneAgenda) { montoAgendado += g.pendienteOC; cantidadAgendada++; }
    else { monto += g.pendienteOC; cantidad++; }
  });
  return { monto, cantidad, montoAgendado, cantidadAgendada };
}

/**
 * Capex Solicitado — dato manual que se edita en la pestaña TechosYFondeos
 * usando el Tipo "Solicitado" (misma estructura: Tipo | Dominio | Proyecto | Monto | ...).
 */
function obtenerCapexSolicitado() {
  const hoja = crearTechosYFondeosSiFalta_();
  const ultimaFila = hoja.getLastRow();
  const porDominio = { 'P&S': 0, 'Tecnología': 0, 'Seguridad': 0 };
  if (ultimaFila < 2) return porDominio;
  hoja.getRange(2, 1, ultimaFila - 1, 5).getValues().forEach(fila => {
    const seccionTexto = String(fila[0] || '').trim();
    const matchNumero = seccionTexto.match(/\d+/);
    const seccion = matchNumero ? matchNumero[0] : seccionTexto;
    const titulo = String(fila[1] || '').trim().toLowerCase();
    if (seccion !== '1' || !titulo.includes('solicitad')) return;
    const dominio = String(fila[2] || '').trim();
    if (porDominio.hasOwnProperty(dominio)) porDominio[dominio] += (Number(fila[4]) || 0) * 1000000;
  });
  return porDominio;
}

/** DIAGNÓSTICO de TechosYFondeos — para ver por qué no está tomando los montos. */
function diagnosticoTechosYFondeos() {
  const partes = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('TechosYFondeos');
    if (!hoja) { partes.push('✗ NO existe la pestaña "TechosYFondeos"'); return devolverYRegistrar_(partes); }
    const ultimaFila = hoja.getLastRow();
    partes.push('Última fila: ' + ultimaFila + ' | Última columna: ' + hoja.getLastColumn());
    if (ultimaFila < 2) { partes.push('✗ No hay filas de datos'); return devolverYRegistrar_(partes); }

    partes.push('Encabezados (fila 1): ' + JSON.stringify(hoja.getRange(1,1,1,5).getValues()[0]));
    const datos = hoja.getRange(2, 1, ultimaFila - 1, 5).getValues();
    datos.forEach((fila, i) => {
      partes.push('Fila ' + (i+2) + ': Sección="' + fila[0] + '" | Título="' + fila[1] + '" | Dominio="' + fila[2] + '" | Monto(E)=' + fila[4]);
    });

    [1, 2, 3].forEach(n => {
      partes.push('\nSección ' + n + ' — títulos que lee el portal: ' + JSON.stringify(obtenerLineasSeccionCapex(n, '', null, '', null)));
    });

    partes.push('\nResultado de obtenerDistribucionCapex():');
    partes.push(JSON.stringify(obtenerDistribucionCapex()));
  } catch (e) {
    partes.push('✗ ERROR: ' + e.message);
  }
  return devolverYRegistrar_(partes);
}
function diagnosticoTechosYFondeosConAviso() {
  SpreadsheetApp.getUi().alert(diagnosticoTechosYFondeos());
}

/**
 * Devuelve las líneas (título + monto) de una Sección de TechosYFondeos,
 * agrupando por Título y sumando la columna E (en millones), filtrado
 * opcionalmente por Dominio (columna C). Dinámico: agrega o quita títulos
 * simplemente editando las filas de la hoja, sin tocar código.
 */
/** Lee TechosYFondeos y arma un iterador simple de filas de UNA sección de "líneas"
 * (excluye siempre las filas "-Velocimetro", que se leen aparte). filtros = {dominio, plazo, vista, trimestre}. */
/** Normaliza el valor de la columna A: "1", 1, "1.0", " 1 " o "Sección 1" -> "1"; "2-Velocimetro" -> "2-velocimetro". */
function normalizarSeccion_(v) {
  let s = String(v == null ? '' : v).replace(/\u00A0/g, ' ').trim().toLowerCase().replace(/^secci[oó]n\s*/, '');
  if (/^\d+(\.0+)?$/.test(s)) s = String(parseInt(s, 10));
  return s;
}

function filasSeccionCapex_(numeroSeccion, filtros) {
  filtros = filtros || {};
  const hoja = crearTechosYFondeosSiFalta_();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();
  const objetivo = normalizarSeccion_(numeroSeccion);
  let seccionActual = '';
  return datos.filter(fila => {
    const propia = normalizarSeccion_(fila[0]);
    if (propia) seccionActual = propia;            // columna A vacía: la fila hereda la sección de la fila de arriba
    if (seccionActual !== objetivo) return false;   // comparación EXACTA: "2" nunca incluye "2-velocimetro"
    const titulo = String(fila[1] || '').trim();
    if (!titulo) return false;
    // Columnas confirmadas con datos reales: C=Dominio, D=Trimestre, E=Plazo, F=(sin usar por ahora), G=Monto
    if (filtros.dominio && String(fila[2] || '').trim() !== filtros.dominio) return false;
    if (filtros.trimestre && String(fila[3] || '').trim() !== filtros.trimestre) return false;
    if (filtros.plazo && String(fila[4] || '').trim() !== filtros.plazo) return false;
    if (filtros.tituloContiene && titulo.toLowerCase().indexOf(filtros.tituloContiene.toLowerCase()) === -1) return false;
    return true;
  });
}

/** Líneas de una sección (1, 2 ó 3), agrupadas y sumadas por Título — para el listado de texto. */
function obtenerLineasSeccionCapex(numeroSeccion, dominioFiltro, trimestreFiltro, plazoFiltro, tituloContieneFiltro) {
  const filas = filasSeccionCapex_(numeroSeccion, { dominio: dominioFiltro, plazo: plazoFiltro, trimestre: trimestreFiltro, tituloContiene: tituloContieneFiltro });
  const orden = [];
  const sumaPorTitulo = {};
  filas.forEach(fila => {
    const titulo = String(fila[1]).trim();
    const monto = Number(fila[6]) || 0;   // columna G — ya viene en soles, no en millones
    if (!sumaPorTitulo.hasOwnProperty(titulo)) { sumaPorTitulo[titulo] = 0; orden.push(titulo); }
    sumaPorTitulo[titulo] += monto;
  });
  return orden.map(titulo => ({ titulo, monto: sumaPorTitulo[titulo] }));
}

/** La misma sección 1, pero agrupada por DOMINIO en vez de por título — para el anillo P&S/Tecnología/Seguridad.
 * tituloFiltro: si se indica, solo suma las filas de ese título (p.ej. "Capex Anual Autorizado"). */
function obtenerDistribucionDominioSeccion(numeroSeccion, tituloFiltro, plazoFiltro, vistaFiltro, trimestreFiltro, dominioFiltro) {
  const filas = filasSeccionCapex_(numeroSeccion, { plazo: plazoFiltro, vista: vistaFiltro, trimestre: trimestreFiltro, dominio: dominioFiltro });
  const porDominio = {};
  filas.forEach(fila => {
    const titulo = String(fila[1]).trim();
    if (tituloFiltro && titulo !== tituloFiltro) return;
    const dominio = String(fila[2] || '').trim() || '(Sin dominio)';
    const monto = Number(fila[6]) || 0;   // columna G — ya viene en soles, no en millones
    porDominio[dominio] = (porDominio[dominio] || 0) + monto;
  });
  return porDominio;
}

/** Lista de valores distintos de Trimestre (columna F) que existan en una Sección, para pintar los botones Q1/Q2/Q3/Q4 dinámicamente. */
function obtenerTrimestresSeccionCapex(numeroSeccion) {
  const filas = filasSeccionCapex_(numeroSeccion, {});
  const set = {};
  filas.forEach(fila => { const t = String(fila[3] || '').trim(); if (t) set[t] = true; });   // columna D
  return Object.keys(set).sort();
}

/**
 * Calcula el % de un velocímetro (Sección 2 ó 3) a partir de las filas "<N>-Velocimetro".
 * tituloNumerador / tituloDenominador son los Títulos EXACTOS (columna B) que trae esa sección:
 *   Sección 2: numerador "Capex Total OC", denominador "Capex Comprometido Total".
 *   Sección 3: numerador "Capex Pagado",   denominador "Capex Total".
 */
function obtenerVelocimetroSeccionCapex(numeroSeccion, tituloNumerador, tituloDenominador, dominioFiltro) {
  const filas = filasVelocimetroSeccionCapex_(numeroSeccion, dominioFiltro);
  const sumar = titulo => filas.filter(f => String(f[1]).trim() === titulo).reduce((s, f) => s + (Number(f[6]) || 0), 0);
  const numerador = sumar(tituloNumerador), denominador = sumar(tituloDenominador);   // columna G, ya en soles
  const pct = denominador > 0 ? (numerador / denominador) * 100 : 0;
  return { pct: pct, numerador: numerador, denominador: denominador };
}
function filasVelocimetroSeccionCapex_(numeroSeccion, dominioFiltro) {
  const hoja = crearTechosYFondeosSiFalta_();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  const datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();
  const objetivo = (String(numeroSeccion).trim() + '-velocimetro').toLowerCase();
  return datos.filter(fila => {
    if (String(fila[0] || '').trim().toLowerCase() !== objetivo) return false;
    if (dominioFiltro && String(fila[2] || '').trim() !== dominioFiltro) return false;
    return true;
  });
}

/** Devuelve los valores crudos de columna A (Sección) y B (Título) tal cual están en TechosYFondeos, para diagnóstico visible directo en el portal. */
function obtenerValoresCrudosTechosYFondeos() {
  const hoja = crearTechosYFondeosSiFalta_();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 3).getValues()
    .map(f => ({ seccion: String(f[0] || ''), titulo: String(f[1] || ''), dominio: String(f[2] || '') }));
}

/* ======================================================================
   SUB-BANNERS DE LÍDERES — hoja "Lideres". Cada bloque empieza en la fila
   donde la columna A tiene cualquier texto (marcador) y el título real
   está en la columna B de esa misma fila. Las filas siguientes (hasta la
   próxima fila marcada en A, o hasta 2 filas vacías seguidas) son los
   datos del bloque: A:F se muestran, D y E son editables, F es "Estado"
   (solo editable por líderes con Sub Rol "Maestro"), y H trae el nombre
   de quién más puede editar esa fila en particular.
   ====================================================================== */

const HOJA_LIDERES = 'Lideres';

/** Correos de todos los usuarios con Sub Rol "Maestro", para copiar en notificaciones. */
function obtenerCorreosMaestros_() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues()
    .filter(f => esMaestro_(f[4]))
    .map(f => f[0])
    .filter(Boolean);
}

/** Correos de los usuarios con Sub Rol "Líder" o "Maestro" (sin duplicados), para el resumen de sincronización. */
function obtenerCorreosLideresYMaestros_() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  if (!hoja || hoja.getLastRow() < 2) return [];
  const correos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues()
    .filter(f => esLider_(f[4]) || esMaestro_(f[4]))
    .map(f => String(f[0] || '').trim())
    .filter(Boolean);
  return [...new Set(correos)];
}

/**
 * Arma la lista de sub-banners leyendo la pestaña "Lideres" de una sola pasada.
 * Estructura de cada bloque:
 *   Fila del marcador: columna A trae el NOMBRE DEL BANNER.
 *   Fila siguiente:    encabezados de B a H (B..F datos, G Estado, H Responsable).
 *   Filas siguientes:  datos, hasta la próxima fila marcada en A o una fila vacía.
 *   Fila siguiente al final de los datos: se escribe ahí la sumatoria de E y F.
 */
function obtenerSubBanners() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_LIDERES);
  if (!hoja || hoja.getLastRow() < 3) return [];

  const ultimaFila = hoja.getLastRow();
  const datos = hoja.getRange(1, 1, ultimaFila, 9).getValues(); // A..I

  const bloques = [];
  let i = 0;
  while (i < datos.length) {
    const colA = String(datos[i][0] || '').trim();
    if (colA === '') { i++; continue; }

    const nombreBanner = colA;
    const filaTitulo = i + 1; // 1-based
    const filaEncabezados = filaTitulo + 1;
    const encabezados = (datos[filaTitulo] || []).slice(1, 8); // B..H de la fila de encabezados

    // Recorrer filas de datos hasta la próxima marca en A o una fila vacía
    const filasDatos = [];
    let j = filaTitulo + 1; // índice 0-based de la primera fila de datos
    while (j < datos.length) {
      const filaActual = datos[j];
      const esNuevaMarca = String(filaActual[0] || '').trim() !== '';
      const vacia = filaActual.slice(1, 8).every(c => c === '' || c === null);
      if (esNuevaMarca || vacia) break;
      filasDatos.push({
        fila: j + 1,
        b: filaActual[1], c: filaActual[2], d: filaActual[3], e: filaActual[4], f: filaActual[5],
        estado: filaActual[6] || '',
        responsable: filaActual[7] || '',
        completado: filaActual[8] === 'Sí' || filaActual[8] === true,
      });
      j++;
    }

    // Sumatoria de E y F — SOLO en memoria para mostrarla; nunca se escribe de
    // vuelta al Sheet (escribirla generaba una fila "fantasma" que en la
    // siguiente lectura se confundía con un dato real y corría todo hacia abajo).
    const sumaE = filasDatos.reduce((s, f) => s + (Number(f.e) || 0), 0);
    const sumaF = filasDatos.reduce((s, f) => s + (Number(f.f) || 0), 0);

    bloques.push({
      titulo: nombreBanner,
      encabezados: encabezados,
      filas: filasDatos,
      sumaE: sumaE,
      sumaF: sumaF,
    });

    i = j; // seguir buscando el próximo bloque desde donde se cortó
  }

  return bloques;
}

/** Guarda un valor editado en una celda de datos (columnas B a G) de un sub-banner. */
function guardarCeldaSubBanner(fila, columna, valor) {
  const u = obtenerSubRolVisitante();
  if (!u.esLider && !u.esMaestro) return { ok: false, error: 'No tienes acceso a esta sección.' };

  // Los montos (B a F) ya no son editables por nadie — vienen del Sheet tal cual.
  if (columna !== 'G') return { ok: false, error: 'Ese dato no es editable desde el portal.' };
  if (!u.esMaestro) return { ok: false, error: 'Solo un perfil Maestro puede cambiar el Estado.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_LIDERES);
  if (!hoja) return { ok: false, error: 'No encontré la pestaña "' + HOJA_LIDERES + '".' };

  hoja.getRange(fila, 7).setValue(valor); // G (Estado, siempre texto)

  return { ok: true };
}

/**
 * Marca/desmarca una fila como completada (columna I) y, si se marca como
 * completada, avisa por correo a todos los perfiles Maestro.
 */
function marcarFilaCompletada(fila, completado, tituloBloque, nombreHoja) {
  const u = obtenerSubRolVisitante();
  if (!u.esLider && !u.esMaestro) return { ok: false, error: 'No tienes acceso a esta sección.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja || HOJA_LIDERES);
  if (!hoja) return { ok: false, error: 'No encontré la pestaña "' + (nombreHoja || HOJA_LIDERES) + '".' };

  hoja.getRange(fila, 9).setValue(completado ? 'Sí' : ''); // columna I

  if (completado) {
    const correosMaestros = obtenerCorreosMaestros_();
    if (correosMaestros.length) {
      MailApp.sendEmail({
        to: correosMaestros.join(','),
        subject: 'Fila completada en "' + (tituloBloque || 'Seguimiento') + '"',
        body: u.nombre + ' marcó como completada la fila ' + fila + ' de "' + (tituloBloque || 'Seguimiento') + '".',
      });
    }
  }
  return { ok: true };
}

/**
 * Envía un correo RECORDATORIO al responsable (columna H) de una fila,
 * con copia a todos los perfiles Maestro, avisando que está pendiente.
 */
function notificarResponsableFila(fila, tituloBloque) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_LIDERES);
  if (!hoja) return { ok: false, error: 'No encontré la pestaña "' + HOJA_LIDERES + '".' };

  const nombreResponsable = String(hoja.getRange(fila, 8).getValue() || '').trim(); // columna H
  if (!nombreResponsable) return { ok: false, error: 'Esta fila no tiene responsable asignado en la columna H.' };

  const hojaUsuarios = ss.getSheetByName('Usuarios');
  let correoResponsable = '';
  if (hojaUsuarios && hojaUsuarios.getLastRow() >= 2) {
    const datosUsuarios = hojaUsuarios.getRange(2, 1, hojaUsuarios.getLastRow() - 1, 2).getValues();
    const encontrado = datosUsuarios.find(f => String(f[1]).trim().toLowerCase() === nombreResponsable.toLowerCase());
    if (encontrado) correoResponsable = encontrado[0];
  }
  if (!correoResponsable) return { ok: false, error: 'No encontré el correo de "' + nombreResponsable + '" en Usuarios.' };

  const correosMaestros = obtenerCorreosMaestros_();
  MailApp.sendEmail({
    to: correoResponsable,
    subject: 'Recordatorio: pendiente de completar — "' + tituloBloque + '"',
    body: 'Hola ' + nombreResponsable + ',\n\nEste es un recordatorio de que la fila ' + fila + ' de "' + tituloBloque + '" sigue pendiente de completar.\n\nPor favor revísala cuando puedas.',
    cc: correosMaestros.join(',') || undefined,
  });
  return { ok: true, correo: correoResponsable };
}

/* ======================================================================
   FONDEO DE CAPEX — hoja "Fondeo". Misma mecánica que Líderes: columna A
   trae el nombre de cada bloque, la fila de abajo son los encabezados
   (B..H), y las filas siguientes son los datos. La suma es solo de la
   columna F, y se escribe con la etiqueta "Capex Aprobado para fondear"
   en la fila inmediatamente después del último dato. G y H son semáforos
   (Aprobado/En revisión/Desestimado y Fondeado/En proceso/Desestimado),
   editables solo por perfiles Maestro.
   ====================================================================== */

const HOJA_FONDEO = 'Fondeo';

function obtenerSubBannersFondeo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_FONDEO);
  if (!hoja || hoja.getLastRow() < 3) return [];

  const ultimaFila = hoja.getLastRow();
  const datos = hoja.getRange(1, 1, ultimaFila, 11).getValues(); // A..K

  const bloques = [];
  let i = 0;
  while (i < datos.length) {
    const colA = String(datos[i][0] || '').trim();
    if (colA === '') { i++; continue; }

    const nombreBanner = colA;
    const filaTitulo = i + 1;
    const encabezados = (datos[filaTitulo] || []).slice(1, 11); // B..K

    const filasDatos = [];
    let j = filaTitulo + 1;
    while (j < datos.length) {
      const filaActual = datos[j];
      const esNuevaMarca = String(filaActual[0] || '').trim() !== '';
      const vacia = filaActual.slice(1, 11).every(c => c === '' || c === null);
      if (esNuevaMarca || vacia) break;
      filasDatos.push({
        fila: j + 1,
        b: filaActual[1], c: filaActual[2], d: filaActual[3], e: filaActual[4],
        dominio: filaActual[5] || '',    // F = Dominio
        trimestre: filaActual[6] || '',  // G = Trimestre
        f: filaActual[7],                // H = Valores
        g: filaActual[8] || '',          // I = Aprobación
        h: filaActual[9] || '',          // J = Fondeo
        k: filaActual[10],               // K = columna nueva (mismo formato que Valores)
      });
      j++;
    }

    // Suma SOLO en memoria (nunca se escribe al Sheet) — dos totales:
    // 1) "Capex Aprobado para fondear": Valores (H) de las filas con Aprobación (I) = "Aprobado"
    // 2) "Capex Fondeado": Valores (H) de las filas con Fondeo (J) = "Fondeado"
    const sumaF = filasDatos
      .filter(f => String(f.g || '').trim().toLowerCase() === 'aprobado')
      .reduce((s, f) => s + (Number(f.f) || 0), 0);
    const sumaFondeado = filasDatos
      .filter(f => String(f.h || '').trim().toLowerCase() === 'fondeado')
      .reduce((s, f) => s + (Number(f.f) || 0), 0);

    bloques.push({ titulo: nombreBanner, encabezados: encabezados, filas: filasDatos, sumaF: sumaF, sumaFondeado: sumaFondeado });
    i = j;
  }
  return bloques;
}

/** Guarda una celda editada en Fondeo. G y H (semáforos) solo Maestro. */
function guardarCeldaSubBannerFondeo(fila, columna, valor) {
  const u = obtenerSubRolVisitante();
  if (!u.esLider && !u.esMaestro) return { ok: false, error: 'No tienes acceso a esta sección.' };

  // Los montos, el Dominio y el Trimestre (B a H) ya no son editables por nadie — vienen del Sheet tal cual.
  if (columna !== 'I' && columna !== 'J') return { ok: false, error: 'Ese dato no es editable desde el portal.' };
  if (!u.esMaestro) return { ok: false, error: 'Solo un perfil Maestro puede cambiar este estado.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_FONDEO);
  if (!hoja) return { ok: false, error: 'No encontré la pestaña "' + HOJA_FONDEO + '".' };

  const colNum = columna === 'I' ? 9 : 10;
  hoja.getRange(fila, colNum).setValue(valor);
  return { ok: true };
}

/** DIAGNÓSTICO de la pestaña "Fondeo" — para ver exactamente cómo se están leyendo las columnas. */
function diagnosticoFondeo() {
  const partes = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(HOJA_FONDEO);
    if (!hoja) { partes.push('✗ NO existe la pestaña "' + HOJA_FONDEO + '"'); return devolverYRegistrar_(partes); }

    const ultimaFila = hoja.getLastRow();
    partes.push('Última fila: ' + ultimaFila + ' | Última columna: ' + hoja.getLastColumn());
    if (ultimaFila < 3) { partes.push('✗ Muy pocas filas para tener un bloque completo.'); return devolverYRegistrar_(partes); }

    const datos = hoja.getRange(1, 1, ultimaFila, 10).getValues(); // A..J
    datos.forEach((fila, i) => {
      partes.push('Fila ' + (i+1) + ': A="' + fila[0] + '" B="' + fila[1] + '" C="' + fila[2] + '" D="' + fila[3] + '" E="' + fila[4] + '" F(Dominio)="' + fila[5] + '" G(Trimestre)="' + fila[6] + '" H(Valores)="' + fila[7] + '" I(Aprobación)="' + fila[8] + '" J(Fondeo)="' + fila[9] + '"');
    });

    partes.push('\nResultado de obtenerSubBannersFondeo():');
    partes.push(JSON.stringify(obtenerSubBannersFondeo()));
  } catch (e) {
    partes.push('✗ ERROR: ' + e.message);
    partes.push('Detalle técnico: ' + (e.stack || '(sin más detalle)'));
  }
  return devolverYRegistrar_(partes);
}
function diagnosticoFondeoConAviso() {
  SpreadsheetApp.getUi().alert(diagnosticoFondeo());
}

/** Lee la pestaña "Overview" con la estructura A=indicador, B=título, C=DominioSP, D=Dominio, E=SDATOOL, F=Proyecto, G=Área, H=Trimestre, I=Tipo, J=Monto. Agrupa y filtra según filtros. */
function obtenerOverview(filtros) {
  filtros = filtros || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Overview');
  if (!hoja || hoja.getLastRow() < 2) return {};

  const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, 10).getValues();
  const resultado = {};
  datos.forEach(fila => {
    const indicador = String(fila[0]||'').trim();
    if (!indicador) return;
    if (filtros.dominioSp && filtros.dominioSp.length && !filtros.dominioSp.includes(String(fila[2]).trim())) return;
    if (filtros.dominio && filtros.dominio.length && !filtros.dominio.includes(String(fila[3]).trim())) return;
    if (filtros.sdatool && filtros.sdatool.length && !filtros.sdatool.includes(String(fila[4]).trim())) return;
    if (filtros.proyecto && filtros.proyecto.length && !filtros.proyecto.includes(String(fila[5]).trim())) return;
    if (filtros.area && filtros.area.length && !filtros.area.includes(String(fila[6]).trim())) return;
    if (filtros.trimestre && filtros.trimestre.length && !filtros.trimestre.includes(String(fila[7]).trim())) return;
    if (filtros.tipo && filtros.tipo.length && !filtros.tipo.includes(String(fila[8]).trim())) return;
    if (!resultado[indicador]) resultado[indicador] = 0;
    resultado[indicador] += Number(fila[9]) || 0;
  });
  return resultado;
}

/** Lee Rep.Capex (Rev) sumando columnas G (SOLPED+IGV), H (OC), I (Ejecutado) con los mismos filtros. */
function obtenerResumenRepCapex(filtros) {
  filtros = filtros || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG_SEMANAL.HOJA_RESUMEN);
  if (!hoja || hoja.getLastRow() < 2) return { solped: 0, oc: 0, ejecutado: 0 };

  const ultimaCol = hoja.getLastColumn();
  const encabezados = hoja.getRange(1,1,1,ultimaCol).getValues()[0];
  const datos = hoja.getRange(2,1,hoja.getLastRow()-1,ultimaCol).getValues();

  const iDomSp = buscarColumnaPorNombre_(encabezados,'dominio sp');
  const iDom = buscarColumnaPorNombre_(encabezados,'dominio');
  const iSdat = buscarColumnaPorNombre_(encabezados,'sdatool');
  const iProy = buscarColumnaPorNombre_(encabezados,'proyecto');
  const iArea = buscarColumnaPorNombre_(encabezados,'area');
  const iTrim = buscarColumnaPorNombre_(encabezados,'trimestre');
  const iTipo = buscarColumnaPorNombre_(encabezados,'tipo');

  let solped = 0, oc = 0, ejecutado = 0, disponible = 0, capexAutorizado = 0;
  datos.forEach(fila => {
    const coincide = (campo, idx) => !campo || !campo.length || idx===-1 || campo.includes(String(fila[idx]||'').trim());
    if (!coincide(filtros.dominioSp, iDomSp)) return;
    if (!coincide(filtros.dominio, iDom)) return;
    if (!coincide(filtros.sdatool, iSdat)) return;
    if (!coincide(filtros.proyecto, iProy)) return;
    if (!coincide(filtros.area, iArea)) return;
    if (!coincide(filtros.trimestre, iTrim)) return;
    if (!coincide(filtros.tipo, iTipo)) return;
    capexAutorizado += Number(fila[4]) || 0; // E
    solped += Number(fila[6]) || 0;          // G
    oc += Number(fila[7]) || 0;              // H
    ejecutado += Number(fila[8]) || 0;       // I
    disponible += Number(fila[9]) || 0;      // J
  });
  return { capexAutorizado, solped, oc, ejecutado, disponible };
}

/** Opciones de filtros disponibles para la pestaña Resumen, sacadas de Overview y Rep.Capex (Rev). */
function obtenerOpcionesFiltrosResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sets = { dominioSp:{}, dominio:{}, sdatool:{}, proyecto:{}, area:{}, tipo:{}, trimestre:{} };

  const hoja1 = ss.getSheetByName('Overview');
  if (hoja1 && hoja1.getLastRow()>=2) {
    hoja1.getRange(2,1,hoja1.getLastRow()-1,10).getValues().forEach(f => {
      if(f[2]) sets.dominioSp[f[2]]=1; if(f[3]) sets.dominio[f[3]]=1;
      if(f[4]) sets.sdatool[f[4]]=1; if(f[5]) sets.proyecto[f[5]]=1;
      if(f[6]) sets.area[f[6]]=1; if(f[7]) sets.trimestre[f[7]]=1; if(f[8]) sets.tipo[f[8]]=1;
    });
  }
  const hoja2 = ss.getSheetByName(CONFIG_SEMANAL.HOJA_RESUMEN);
  if (hoja2 && hoja2.getLastRow()>=2) {
    const enc = hoja2.getRange(1,1,1,hoja2.getLastColumn()).getValues()[0];
    const dat = hoja2.getRange(2,1,hoja2.getLastRow()-1,hoja2.getLastColumn()).getValues();
    const idx = { dominioSp: buscarColumnaPorNombre_(enc,'dominio sp'), dominio: buscarColumnaPorNombre_(enc,'dominio'), sdatool: buscarColumnaPorNombre_(enc,'sdatool'), proyecto: buscarColumnaPorNombre_(enc,'proyecto'), area: buscarColumnaPorNombre_(enc,'area'), trimestre: buscarColumnaPorNombre_(enc,'trimestre'), tipo: buscarColumnaPorNombre_(enc,'tipo') };
    dat.forEach(f => { Object.keys(idx).forEach(k => { if(idx[k]!==-1 && f[idx[k]]) sets[k][f[idx[k]]]=1; }); });
  }
  return Object.fromEntries(Object.entries(sets).map(([k,v]) => [k, Object.keys(v).sort()]));
}

/** Diagnóstico rápido: muestra cuántas columnas tiene BASE_CASHFLOW y qué hay en los encabezados. */
function diagnosticoColumnaValidacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.HOJA_BASE);
  if (!hoja) return 'No existe la hoja ' + CONFIG.HOJA_BASE;
  const nCols = hoja.getLastColumn();
  const encabezados = hoja.getRange(1, 1, 1, nCols).getValues()[0];
  const lineas = [
    'Columnas en BASE_CASHFLOW: ' + nCols,
    'ENCABEZADOS_BASE espera: ' + ENCABEZADOS_BASE.length,
    'Encabezados actuales: ' + encabezados.join(' | '),
    '',
    'COL_BASE.ESTADO_VALIDACION = ' + COL_BASE.ESTADO_VALIDACION,
    '¿Existe esa columna? ' + (nCols >= COL_BASE.ESTADO_VALIDACION ? 'SÍ' : 'NO — necesita sincronizar para crearla'),
  ];
  if (nCols >= 2) {
    const muestraFila2 = hoja.getRange(2, Math.max(1, nCols-5), 1, Math.min(6, nCols)).getValues()[0];
    lineas.push('Últimas celdas fila 2: ' + muestraFila2.join(' | '));
  }
  const resultado = lineas.join('\n');
  Logger.log(resultado);
  SpreadsheetApp.getUi().alert(resultado);
  return resultado;
}

/**
 * Diagnóstico puntual para el bug de "Pagos registrados por mes" no mostrando un dominio: muestra,
 * para hasta 8 OC Pagadas de cada dominio, el valor CRUDO de "Fecha de documento" (columna B) tal
 * cual lo devuelve getValues() — así se ve directo si llega como objeto Date real, como texto, o
 * vacío, sin tener que adivinar. Llamar desde el editor de Apps Script (▶) con esta función seleccionada.
 */
function diagnosticoFechaDocumentoPorDominio() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hoja || hoja.getLastRow() < 2) return 'BASE_CASHFLOW está vacía.';
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hoja.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos).filter(g => g.estado === 'Pagado');
  const ocsAptas = {};
  agrupadas.forEach(g => { ocsAptas[String(g.oc)] = true; });
  const porDominio = {};
  datos.forEach(fila => {
    if (!ocsAptas[String(fila[0])] || Number(fila[14]) <= 0) return;   // solo líneas con algo pagado, de una OC que sigue Pagada
    const dom = fila[5] || 'Sin dominio';
    (porDominio[dom] = porDominio[dom] || []).push({ oc: fila[0], valor: fila[1] });
  });
  const lineas = ['Valores crudos de "Fecha de documento" (columna B) por dominio, hasta 8 por dominio:', ''];
  Object.keys(porDominio).sort().forEach(dom => {
    lineas.push('— ' + dom + ' (' + porDominio[dom].length + ' líneas con algo pagado) —');
    porDominio[dom].slice(0, 8).forEach(f => {
      const v = f.valor;
      const tipo = v === '' || v === null ? 'VACÍO' : (Object.prototype.toString.call(v) === '[object Date]' ? 'Date real' : 'texto/otro (' + typeof v + ')');
      const parseado = parsearFechaDocumento_(v);
      lineas.push('  OC ' + f.oc + ': valor=' + JSON.stringify(v) + ' | tipo=' + tipo + ' | ¿se interpreta bien? ' + (parseado ? Utilities.formatDate(parseado, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'NO'));
    });
    lineas.push('');
  });
  const resultadoDiag = lineas.join('\n');
  Logger.log(resultadoDiag);
  return resultadoDiag;
}

/* =====================================================================
   CORREO SEMANAL AUTOMÁTICO — cada lunes a las 9:00 AM
   Se instala con instalarTriggerCorreoSemanal() desde el menú.
   ===================================================================== */

/** Instala (o reinstala) el trigger de correo semanal para los lunes a las 9 AM. */
function instalarTriggerCorreoSemanal() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'enviarCorreoSemanalValidacion')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('enviarCorreoSemanalValidacion')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log('Trigger instalado: cada lunes a las 9:00 AM.');
  try {
    SpreadsheetApp.getUi().alert('✅ Listo. Cada lunes a las 9:00 AM se enviará el correo de validación a los perfiles Maestro.');
  } catch(e) { /* si se llama desde trigger no hay UI */ }
}

/** Función que ejecuta el trigger semanal (lunes 9:00). Se mantiene con este nombre para no romper triggers ya instalados. */
function enviarCorreoSemanalValidacion() { return enviarReporteSemanalCore_('Automático'); }

/** Arma y envía el correo de validación de la semana anterior. origen: 'Automático' | 'Manual'. Devuelve un resumen. */
function enviarReporteSemanalCore_(origen) {
  const correosMaestros = obtenerCorreosMaestros_();
  if (!correosMaestros.length) return { ok: false, error: 'No hay perfiles Maestro con correo en la pestaña Usuarios.' };

  const hoy = new Date();
  const lunes = new Date(hoy);
  lunes.setDate(hoy.getDate() - ((hoy.getDay() + 6) % 7)); // lunes de esta semana
  lunes.setHours(0,0,0,0);
  const lunest = new Date(lunes);
  const lunesSemAnt = new Date(lunes); lunesSemAnt.setDate(lunes.getDate() - 7);
  const viernesSemAnt = new Date(lunesSemAnt); viernesSemAnt.setDate(lunesSemAnt.getDate() + 4);

  const fmtFecha = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const fmtMonto = n => 'S/ ' + (Math.round(Number(n)||0)).toLocaleString('es-PE');

  // Leer BASE_CASHFLOW
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return { ok: false, error: 'BASE_CASHFLOW está vacía.' };
  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow()-1, ENCABEZADOS_BASE.length).getValues();
  const agrupadas = agruparPorOC_(datos);

  // Clasificar pagos programados (cuotas y fechas propuestas) de la semana anterior
  const items = construirPagosProgramados_(agrupadas, leerValidacionCuotas_());
  const semAnt = [], pendientes = [], verificados = [], sinPago = [], parciales = [];
  items.forEach(it => {
    const fp = new Date(it.fechaPago); fp.setHours(0,0,0,0);
    if (fp < lunesSemAnt || fp > viernesSemAnt) return;
    const val = String(it.estado || '').toLowerCase();
    const g2 = { oc: it.oc + (it.esCuota ? ' (cuota ' + it.cuota + ')' : ''), sdatool: it.g.sdatool, proyecto: it.g.proyecto,
                 monto: it.monto, fechaPago: it.fechaPago, estadoValidacion: it.estado };
    semAnt.push(g2);
    if (val === 'verificado') verificados.push(g2);
    else if (val === 'revisado - sin pago') sinPago.push(g2);
    else if (val === 'pago parcial') parciales.push(g2);
    else pendientes.push(g2);
  });
  const sumar = arr => arr.reduce((t, g) => t + g.monto, 0);
  const totalSemAnt = sumar(semAnt), totalVer = sumar(verificados), totalPend = sumar(pendientes);
  const pct = totalSemAnt > 0 ? Math.round(totalVer / totalSemAnt * 100) : 0;

  const filaDetalle = g =>
    `<tr>
      <td style="padding:6px 10px;border-bottom:1px solid #eee;">${g.oc}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #eee;">${g.sdatool||''}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #eee;">${g.proyecto||''}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #eee;">${fmtMonto(g.monto)}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #eee;">${fmtFecha(new Date(g.fechaPago))}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #eee;font-weight:700;color:${g.estadoValidacion==='Verificado'?'#4C9270':'#B98A3D'};">${g.estadoValidacion||'Pendiente'}</td>
    </tr>`;

  const tablaHeader = `<table style="border-collapse:collapse;width:100%;font-size:13px;">
    <thead><tr style="background:#f0f4f8;">
      <th style="padding:8px 10px;text-align:left;">N° OC</th>
      <th style="padding:8px 10px;text-align:left;">SDATOOL</th>
      <th style="padding:8px 10px;text-align:left;">Proyecto</th>
      <th style="padding:8px 10px;text-align:left;">Monto</th>
      <th style="padding:8px 10px;text-align:left;">Fecha Pago</th>
      <th style="padding:8px 10px;text-align:left;">Validación</th>
    </tr></thead><tbody>`;

  const cuerpo = `
    <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">
      <div style="background:#0B1D33;color:#fff;padding:24px;border-radius:8px 8px 0 0;">
        <h2 style="margin:0;font-size:18px;">Control Presupuestal — Resumen semanal de validación</h2>
        <p style="margin:8px 0 0;opacity:.8;font-size:13px;">Semana ${fmtFecha(lunesSemAnt)} al ${fmtFecha(viernesSemAnt)}</p>
      </div>
      <div style="background:#fff;padding:24px;border:1px solid #e0e0e0;">
        <div style="display:flex;gap:16px;margin-bottom:24px;flex-wrap:wrap;">
          <div style="background:#EEF4FF;border-left:4px solid #1B5FAE;padding:12px 18px;border-radius:8px;flex:1;">
            <div style="font-size:11px;color:#1B5FAE;font-weight:700;">TOTAL SEMANA</div>
            <div style="font-size:22px;font-weight:800;color:#1B5FAE;">${fmtMonto(totalSemAnt)}</div>
            <div style="font-size:11px;color:#6B7A99;">${semAnt.length} órdenes</div>
          </div>
          <div style="background:#E4F1EA;border-left:4px solid #4C9270;padding:12px 18px;border-radius:8px;flex:1;">
            <div style="font-size:11px;color:#4C9270;font-weight:700;">VERIFICADO</div>
            <div style="font-size:22px;font-weight:800;color:#4C9270;">${fmtMonto(totalVer)}</div>
            <div style="font-size:11px;color:#6B7A99;">${verificados.length} órdenes — ${pct}%</div>
          </div>
          <div style="background:#FFF8EE;border-left:4px solid #B98A3D;padding:12px 18px;border-radius:8px;flex:1;">
            <div style="font-size:11px;color:#B98A3D;font-weight:700;">PENDIENTE</div>
            <div style="font-size:22px;font-weight:800;color:#B98A3D;">${fmtMonto(totalPend)}</div>
            <div style="font-size:11px;color:#6B7A99;">${pendientes.length} órdenes</div>
          </div>
        </div>
        ${pendientes.length > 0 ? `
        <h3 style="color:#B98A3D;margin-bottom:10px;">⚠️ Pendientes de verificar (${pendientes.length})</h3>
        ${tablaHeader}${pendientes.map(filaDetalle).join('')}</tbody></table>
        ` : (sinPago.length + parciales.length === 0 ? '<p style="color:#4C9270;font-weight:600;">✅ Todos los pagos de la semana fueron verificados.</p>' : '<p style="color:#4C9270;font-weight:600;">✅ No quedan pagos pendientes de revisar.</p>')}
        ${sinPago.length > 0 ? `
        <h3 style="color:#6B7A99;margin:20px 0 10px;">👁 Revisadas sin evidencia de pago (${sinPago.length} — ${fmtMonto(sumar(sinPago))})</h3>
        ${tablaHeader}${sinPago.map(filaDetalle).join('')}</tbody></table>
        ` : ''}
        ${parciales.length > 0 ? `
        <h3 style="color:#1B5FAE;margin:20px 0 10px;">◐ Con pago parcial (${parciales.length} — ${fmtMonto(sumar(parciales))})</h3>
        ${tablaHeader}${parciales.map(filaDetalle).join('')}</tbody></table>
        ` : ''}
        ${verificados.length > 0 ? `
        <h3 style="color:#4C9270;margin:20px 0 10px;">✅ Verificados (${verificados.length})</h3>
        ${tablaHeader}${verificados.map(filaDetalle).join('')}</tbody></table>
        ` : ''}
      </div>
      <div style="background:#f0f4f8;padding:12px 24px;border-radius:0 0 8px 8px;font-size:11px;color:#6B7A99;">
        Generado automáticamente cada lunes • Control Presupuestal P&S
      </div>
    </div>`;

  MailApp.sendEmail({
    to: correosMaestros.join(','),
    subject: `Validación de pagos — Semana ${fmtFecha(lunesSemAnt)}–${fmtFecha(viernesSemAnt)} (${pct}% verificado)`,
    htmlBody: cuerpo,
  });

  const res = { ok: true, destinatarios: correosMaestros.length, pagos: semAnt.length, pct: pct, semana: fmtFecha(lunesSemAnt) + ' – ' + fmtFecha(viernesSemAnt) };
  try {
    PropertiesService.getScriptProperties().setProperty('REPORTE_ULTIMO', JSON.stringify({
      fecha: Utilities.formatDate(new Date(), REPORTE_TZ, 'dd/MM/yyyy HH:mm'), origen: origen || 'Automático',
      por: origen === 'Manual' ? (obtenerCorreoVisitante_() || '') : '', destinatarios: res.destinatarios, pagos: res.pagos, pct: res.pct, semana: res.semana,
    }));
  } catch (e) {}
  return res;
}

/** Devuelve las últimas 10 OCs modificadas (por fecha de modificación, col B). */
function obtenerUltimasOCModificadas() {
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return [];
  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow()-1, ENCABEZADOS_BASE.length).getValues();
  const agrupadas = agruparPorOC_(datos);
  return agrupadas
    .filter(g => g.fechaMod)
    .sort((a,b) => new Date(b.fechaMod) - new Date(a.fechaMod))
    .slice(0, 10)
    .map(g => ({
      oc: g.oc,
      proyecto: g.proyecto || '',
      pendienteOC: g.pendienteOC,
      pagado: g.pagado,
      estadoValidacion: g.estadoValidacion || '',
    }));
}

/* ====================================================================
   ALERTAS DEL PORTAL — pestaña "Alertas"
   A=Mensaje, B=Inicio, C=Fin, D=BloquearEdicion(Sí/No), E=Activa(Sí/No)
   ==================================================================== */

function obtenerAlertas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Alertas');
  if (!hoja) {
    hoja = ss.insertSheet('Alertas');
    hoja.getRange(1,1,1,5).setValues([['Mensaje','Inicio (dd/mm/aaaa hh:mm)','Fin (dd/mm/aaaa hh:mm)','Bloquear edición (Sí/No)','Activa (Sí/No)']]);
    hoja.setFrozenRows(1);
    return { alertas: [], bloquearEdicion: false };
  }
  if (hoja.getLastRow() < 2) return { alertas: [], bloquearEdicion: false };

  const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, 5).getValues();
  const ahora = new Date();
  const alertasActivas = [];
  let bloquearEdicion = false;

  datos.forEach(fila => {
    const activa = String(fila[4]||'').trim().toLowerCase();
    if (activa !== 'sí' && activa !== 'si') return;
    const msg = String(fila[0]||'').trim();
    if (!msg) return;
    const inicio = fila[1] ? new Date(fila[1]) : null;
    const fin    = fila[2] ? new Date(fila[2]) : null;
    if (inicio && ahora < inicio) return; // aún no empieza
    if (fin && ahora > fin) return;       // ya terminó
    alertasActivas.push(msg);
    if (String(fila[3]||'').trim().toLowerCase().startsWith('s')) bloquearEdicion = true;
  });

  return { alertas: alertasActivas, bloquearEdicion };
}
/* ============================================================
   FUNCIONES NUEVAS — pegar AL FINAL del Código.gs existente
   (reemplaza cualquier versión anterior de estas funciones)
   ============================================================ */

/* ── SOLICITUD DE OC ─────────────────────────────────────── */
const HOJA_SOL_OC = 'SolicitudesOC';
const DEST_EMAIL_OC = 'mvillon@bbva.com';

function registrarSolicitudOC(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(HOJA_SOL_OC);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_SOL_OC);
    hoja.getRange(1,1,1,14).setValues([[
      'Fecha','Remitente','Tipo','N° Ref','Proyecto','Oficina','Cuenta',
      'Centro de Costo','Detalle','Monto sin IGV','Monto con IGV','Empresa',
      'Correos copia','Estado'
    ]]);
    hoja.setFrozenRows(1);
  }
  const correo = datos.remitenteCorreo || Session.getActiveUser().getEmail() || '';
  const nombre = datos.remitenteNombre || correo;
  const fecha  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  hoja.appendRow([
    fecha, correo, datos.tipo||'SOLPED', datos.numRef||'',
    datos.proyecto||'', datos.oficina||'', datos.cuenta||'',
    datos.centroCosto||'', datos.desc||'',
    Number(datos.montoSinIGV)||0, Number(datos.montoConIGV)||0,
    datos.empresa||'', datos.copias||'', 'Pendiente'
  ]);

  const maestros = obtenerCorreosMaestros_();
  const cc = [...new Set([...maestros, ...(datos.copias||'').split(',').map(s=>s.trim()).filter(Boolean)])].join(',');
  const fmtM = n => 'S/ ' + Math.round(Number(n)||0).toLocaleString('es-PE');

  MailApp.sendEmail({
    to: DEST_EMAIL_OC, cc: cc||undefined, replyTo: correo,
    subject: '[' + (datos.tipo||'SOLPED') + '] Solicitud — ' + datos.proyecto,
    htmlBody: '<div style="font-family:Arial;max-width:580px">'
      + '<div style="background:#0B1D33;color:#fff;padding:18px 22px;border-radius:8px 8px 0 0">'
      + '<h2 style="margin:0;font-size:16px">📦 ' + (datos.tipo||'SOLPED') + ' — Solicitud de compra</h2>'
      + '<p style="margin:4px 0 0;font-size:12px;opacity:.75">Enviado por ' + (nombre||correo) + ' · ' + fecha + '</p></div>'
      + '<div style="padding:20px 22px;border:1px solid #e0e0e0"><table style="font-size:13px;width:100%;border-collapse:collapse">'
      + '<tr><td style="color:#666;padding:6px 0;width:40%">Proyecto</td><td><b>' + datos.proyecto + '</b></td></tr>'
      + (datos.oficina ? '<tr><td style="color:#666;padding:6px 0">Oficina</td><td>' + datos.oficina + '</td></tr>' : '')
      + (datos.cuenta  ? '<tr><td style="color:#666;padding:6px 0">Cuenta</td><td>' + datos.cuenta  + '</td></tr>' : '')
      + (datos.centroCosto ? '<tr><td style="color:#666;padding:6px 0">Centro de Costo</td><td>' + datos.centroCosto + '</td></tr>' : '')
      + '<tr><td style="color:#666;padding:6px 0">Empresa</td><td>' + datos.empresa + '</td></tr>'
      + '<tr><td style="color:#666;padding:6px 0">Monto sin IGV</td><td>' + fmtM(datos.montoSinIGV) + '</td></tr>'
      + '<tr><td style="color:#666;padding:6px 0">Monto con IGV</td><td><b style="color:#1B5FAE;font-size:15px">' + fmtM(datos.montoConIGV) + '</b></td></tr>'
      + '</table>'
      + '<div style="background:#f5f7fa;border-radius:6px;padding:10px;margin-top:12px;font-size:13px">' + (datos.desc||'') + '</div>'
      + '</div></div>',
    attachments: datos.adjuntoBase64 ? [Utilities.newBlob(
      Utilities.base64Decode(datos.adjuntoBase64), datos.adjuntoMime||'application/octet-stream', datos.adjuntoNombre
    )] : undefined,
  });
  return { ok: true };
}

function obtenerResumenSolicitudesOC() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_SOL_OC);
  if (!hoja || hoja.getLastRow() < 2) return { total:0, pendiente:0, aprobada:0, rechazada:0, filas:[] };
  const datos = hoja.getRange(2,1,hoja.getLastRow()-1,14).getValues();
  let p=0,a=0,r=0;
  const filas = datos.map(f => {
    const est = String(f[13]||'Pendiente');
    if(est==='Pendiente')p++; else if(est==='Aprobada')a++; else if(est==='Rechazada')r++;
    return {
      fecha:       f[0] ? Utilities.formatDate(new Date(f[0]), Session.getScriptTimeZone(),'dd/MM/yyyy') : '',
      tipo:        f[2]||'SOLPED',
      proyecto:    f[4],
      empresa:     f[11],
      montoSinIGV: f[9],
      montoConIGV: f[10],
      estado:      est,
    };
  }).reverse().slice(0,30);
  return { total:datos.length, pendiente:p, aprobada:a, rechazada:r, filas };
}

function obtenerOpcionesGB() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const vals = hoja.getRange(2, 12, hoja.getLastRow()-1, 1).getValues();
  const set = {};
  vals.forEach(([v]) => { if(v) set[String(v).trim()] = 1; });
  return Object.keys(set).sort();
}

/* ── APROBAR TODAS EN BLOQUE ─────────────────────────────── */
function aprobarTodasOCs(listaOCs) {
  if (!listaOCs || !listaOCs.length) return { ok:true, count:0 };
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hoja) return { ok:false, error:'Sin hoja base' };
  const set = {};
  listaOCs.forEach(oc => { set[String(oc)] = true; });
  const col = hoja.getRange(2, COL_BASE.OC, hoja.getLastRow()-1, 1).getValues();
  const fechaValActual = hoja.getRange(2, COL_BASE.FECHA_PAGO_VALIDADA, col.length, 1).getValues();
  // Columnas Valorizado + cuotas (fecha/monto 1..5) para poder verificar también las cuotas de OCs valorizadas
  const anchoVal = COL_BASE.MONTO_PAGO_5 - COL_BASE.VALORIZADO + 1;
  const vals = hoja.getRange(2, COL_BASE.VALORIZADO, col.length, anchoVal).getValues();
  const cuotasPorVerificar = [];
  const yaVistas = {};
  const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let count = 0;
  for (let i = 0; i < col.length; i++) {
    const oc = String(col[i][0]);
    if (set[oc]) {
      hoja.getRange(i+2, COL_BASE.ESTADO_VALIDACION).setValue('Verificado');
      // Si nunca se cargó una Fecha de pago validada, se pone la de hoy — así este pago ya
      // queda con una fecha real y puede ubicarse en "Pagos registrados por mes" (antes se perdía
      // para siempre si la fuente nunca traía "Fecha de documento", que es el caso de casi toda la base).
      if (!fechaValActual[i][0]) hoja.getRange(i+2, COL_BASE.FECHA_PAGO_VALIDADA).setValue(hoy);
      count++;
      if (vals[i][0] === true && !yaVistas[oc]) {
        yaVistas[oc] = true;
        for (let n = 1; n <= 5; n++) {
          const fecha = vals[i][1 + 2 * (n - 1)], monto = Number(vals[i][2 + 2 * (n - 1)]) || 0;
          if (fecha && monto > 0) cuotasPorVerificar.push([oc, n]);
        }
      }
    }
  }
  marcarCuotasVerificadas_(cuotasPorVerificar);
  bumpCacheVer_();
  return { ok:true, count };
}

/** Marca como Verificado las cuotas indicadas [[oc, n], ...] en la hoja ValidacionCuotas (crea las filas que falten). */
function marcarCuotasVerificadas_(pares) {
  if (!pares || !pares.length) return 0;
  const h = obtenerHojaValCuotas_();
  const ult = h.getLastRow();
  const idx = {};
  const fechaValActual = {};
  if (ult >= 2) {
    h.getRange(2, 1, ult - 1, 3).getValues().forEach((f, i) => {
      const clave = String(f[0]) + '|' + Number(f[1]);
      idx[clave] = i + 2;
      fechaValActual[clave] = f[2];
    });
  }
  const usuario = obtenerCorreoVisitante_() || '', ahora = new Date();
  const hoy = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nuevas = [];
  pares.forEach(pr => {
    const clave = String(pr[0]) + '|' + pr[1];
    const fila = idx[clave];
    if (fila) {
      h.getRange(fila, 4).setValue('Verificado');
      h.getRange(fila, 6, 1, 2).setValues([[usuario, ahora]]);
      // Igual que en aprobarTodasOCs: si nunca se cargó Fecha validada para esta cuota, se pone la
      // de hoy — para que ese pago tenga alguna fecha y pueda ubicarse en "Pagos registrados por mes".
      if (!fechaValActual[clave]) h.getRange(fila, 3).setValue(hoy);
    }
    else nuevas.push([String(pr[0]), pr[1], hoy, 'Verificado', '', usuario, ahora]);
  });
  if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, 7).setValues(nuevas);
  return pares.length;
}

/* ── CARGA INICIAL (2 pasos) ─────────────────────────────── */

/**
 * PASO 1 — solo lee hoja Usuarios y Alertas. Sin BASE_CASHFLOW.
 * Responde en < 0.5s. No envía correos.
 */
function obtenerUsuarioYRolRapido() {
  try {
    const correo = Session.getActiveUser().getEmail() || '';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Usuarios');
    let nombre = correo, rol = '', subRol = '', autorizado = false;
    if (hoja && hoja.getLastRow() >= 2) {
      const filas = hoja.getRange(2,1,hoja.getLastRow()-1,5).getValues();
      const f = filas.find(r => String(r[0]).trim().toLowerCase() === correo.toLowerCase());
      if (f) { autorizado=true; nombre=f[1]||correo; rol=String(f[3]||'').toLowerCase(); subRol=String(f[4]||'').toLowerCase(); }
    }
    let alertas = { alertas:[], bloquearEdicion:false };
    try { alertas = obtenerAlertas(); } catch(e) {}
    return {
      autorizado, correo, nombre,
      esAdmin:  rol.includes('editor')||rol.includes('programador')||rol.includes('admin'),
      esLider:  subRol.includes('líder')||subRol.includes('lider'),
      esMaestro: subRol.includes('maestro'),
      alertas,
    };
  } catch(e) {
    return { autorizado:true, correo:'', nombre:'Usuario', esAdmin:false, esLider:false, esMaestro:false, alertas:{alertas:[],bloquearEdicion:false} };
  }
}

/**
 * PASO 2 — datos de Vista General. Se llama DESPUÉS de mostrar la UI.
 */
function obtenerDatosIniciales() {
  return conCache_('inicial', {}, () => {
    const filtros = {};
    const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
    const datos = (hojaBase && hojaBase.getLastRow() >= 2)
      ? hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues()
      : [];
    return {
      resumen:             obtenerResumenDashboard(filtros, datos),
      distribucion:        obtenerDistribucionCapex(),
      ordenesPendientes:   obtenerResumenOrdenesPendientes(filtros, datos),
      pendientePorAgendar: obtenerPendientePorAgendar(filtros, datos),
    };
  });
}

/* ── ÚLTIMAS OCS MODIFICADAS ─────────────────────────────── */

/** Aprueba o rechaza en bloque todos los cambios de fecha pendientes. */
function accionMasivaCambiosFecha(accion, filtros) {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede aprobar o rechazar cambios de fecha.' };
  return resolverCambiosFecha(obtenerCambiosFechaPendientes().map(c => c.fila), accion === 'aprobar');
}

/**
 * Aprueba o rechaza VARIOS cambios de fecha a la vez (todos o los seleccionados).
 * filas = números de fila en HistorialFechas. Lee la base una sola vez, así que es rápido aunque sean muchos.
 * Solo se procesan los que siguen "Pendiente"; si una OC tiene varios cambios, se aplican en orden y queda el último.
 */
function resolverCambiosFecha(filas, aprobar) {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede aprobar o rechazar cambios de fecha.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('HistorialFechas');
  if (!hoja || hoja.getLastRow() < 2) return { ok: true, count: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ultima = hoja.getLastRow();
    const hist = hoja.getRange(2, 1, ultima - 1, 6).getValues();
    const pendientes = (filas || []).map(Number).filter(f => f >= 2 && f <= ultima)
      .filter((f, i, a) => a.indexOf(f) === i).sort((a, b) => a - b)
      .filter(f => { const est = String(hist[f - 2][5] || '').trim(); return est === 'Pendiente' || est === ''; });
    if (!pendientes.length) return { ok: true, count: 0 };

    if (aprobar) {
      const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
      const col = hojaBase.getRange(2, COL_BASE.OC, hojaBase.getLastRow() - 1, 1).getValues();
      const indice = {};
      col.forEach((f, i) => { const k = String(f[0]); (indice[k] = indice[k] || []).push(i + 2); });
      pendientes.forEach(f => {
        const oc = String(hist[f - 2][0]), fechaNueva = hist[f - 2][2];
        (indice[oc] || []).forEach(r => hojaBase.getRange(r, COL_BASE.FECHA_PAGO).setValue(fechaNueva));
      });
      bumpCacheVer_();
    }
    pendientes.forEach(f => hoja.getRange(f, 6).setValue(aprobar ? 'Aprobado' : 'Rechazado'));
    return { ok: true, count: pendientes.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * KPIs para el Cash Flow — lee BASE_CASHFLOW completo y calcula:
 * total OC, cuántas tienen fecha de pago, montos, y cuántas están verificadas.
 * Respeta los filtros de dominio/dominioSp/sdatool/proyecto.
 */
function obtenerKPIsGantt(filtros) {
  filtros = filtros || {};
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hoja || hoja.getLastRow() < 2) return { total:0, conFecha:0, verificadas:0, montoTotal:0, montoConFecha:0 };

  const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, Math.min(ENCABEZADOS_BASE.length, hoja.getLastColumn())).getValues();
  const agrupadas = agruparPorOC_(datos);

  let total=0, conFecha=0, verificadas=0, montoTotal=0, montoConFecha=0;
  const vcMap = leerValidacionCuotas_();

  agrupadas.forEach(g => {
    // Aplicar filtros de dominio/sp/sdatool/proyecto
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (!coincideFiltro_(filtros.dominioSp, (g.dominioSp||'').trim()||'(Sin dato)')) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;

    const monto = (g.pendienteOC||0) + (g.pagado||0);
    total++;
    montoTotal += monto;

    // Valorizada: cuentan sus cuotas (agendada si tiene alguna cuota con fecha y monto; verificada si TODAS sus cuotas lo están)
    const cuotasOC = g.valorizado
      ? [1,2,3,4,5].filter(n => g['fechaPago'+n] && Number(g['montoPago'+n]) > 0)
      : [];
    const tieneAgenda = g.valorizado ? cuotasOC.length > 0 : !!g.fechaPago;
    if (tieneAgenda) {
      conFecha++;
      montoConFecha += monto;
    }
    const verificada = g.valorizado
      ? (cuotasOC.length > 0 && cuotasOC.every(n => String((vcMap[String(g.oc)+'|'+n]||{}).estado||'').toLowerCase() === 'verificado'))
      : (g.estadoValidacion||'').toLowerCase() === 'verificado';
    if (verificada) verificadas++;
  });

  return { total, conFecha, verificadas, montoTotal, montoConFecha };
}


/* ====================================================================
   VALIDACIÓN DE PAGOS POR PAGO PROGRAMADO (cuotas / pago parcial)
   Estados: Verificado | Pendiente | Desfasado | Revisado - sin pago | Pago parcial
   - OC simple o con pago parcial: se guarda en BASE_CASHFLOW (Fecha validada + Validación), como siempre.
   - OC valorizada: cada cuota se valida por separado en la hoja "ValidacionCuotas"
     (no toca la estructura de BASE_CASHFLOW, así que la sincronización no la afecta).
   ==================================================================== */
const HOJA_VAL_CUOTAS = 'ValidacionCuotas';
const ESTADOS_VALIDACION = ['Verificado', 'Pendiente', 'Desfasado', 'Revisado - sin pago', 'Pago parcial'];

function obtenerHojaValCuotas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA_VAL_CUOTAS);
  if (!h) {
    h = ss.insertSheet(HOJA_VAL_CUOTAS);
    h.getRange(1, 1, 1, 7).setValues([['N° OC', 'Cuota', 'Fecha validada', 'Estado', 'Comentario', 'Usuario', 'Actualizado']]);
    h.setFrozenRows(1);
  }
  return h;
}

/** Devuelve { 'OC|n': {fechaValidada, estado, comentario} } */
function leerValidacionCuotas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(HOJA_VAL_CUOTAS);   // no la crea solo por leer
  const mapa = {};
  if (!h || h.getLastRow() < 2) return mapa;
  h.getRange(2, 1, h.getLastRow() - 1, 5).getValues().forEach(f => {
    if (f[0] === '' || f[0] === null) return;
    mapa[String(f[0]) + '|' + Number(f[1])] = { fechaValidada: f[2], estado: f[3] || '', comentario: f[4] || '' };
  });
  return mapa;
}

/** campo: 'fechaValidada' | 'estado' | 'comentario'. Estado solo lo cambia un perfil Maestro. */
function guardarValidacionCuota(numeroOC, cuota, campo, valor) {
  const alertas = obtenerAlertas();
  const u = obtenerSubRolVisitante();
  if (alertas.bloquearEdicion && !u.esMaestro) return { ok: false, error: '🔒 El portal está en modo solo lectura. Contacta a un perfil Maestro.' };
  const colPorCampo = { fechaValidada: 3, estado: 4, comentario: 5 };
  const col = colPorCampo[campo];
  if (!col) return { ok: false, error: 'Campo no reconocido: ' + campo };
  if (campo === 'estado') {
    if (!u.esMaestro) return { ok: false, error: 'Solo un perfil Maestro puede cambiar la Validación.' };
    if (valor && ESTADOS_VALIDACION.indexOf(valor) === -1) return { ok: false, error: 'Estado no válido: ' + valor };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const h = obtenerHojaValCuotas_();
    const ult = h.getLastRow();
    let fila = -1;
    if (ult >= 2) {
      const claves = h.getRange(2, 1, ult - 1, 2).getValues();
      for (let i = 0; i < claves.length; i++) {
        if (String(claves[i][0]) === String(numeroOC) && Number(claves[i][1]) === Number(cuota)) { fila = i + 2; break; }
      }
    }
    if (fila === -1) { fila = ult + 1; h.getRange(fila, 1, 1, 2).setValues([[String(numeroOC), Number(cuota)]]); }
    h.getRange(fila, col).setValue(valor);
    h.getRange(fila, 6, 1, 2).setValues([[obtenerCorreoVisitante_() || '', new Date()]]);
    bumpCacheVer_();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Convierte las OCs agrupadas en "pagos programados" a validar:
 *  - OC valorizada  → una fila por cuota (fecha + monto de la cuota, validación propia).
 *  - OC simple      → una fila con su fecha propuesta. Si ya tiene pagos (pagado > 0) y aún hay saldo,
 *                     es un PAGO PARCIAL: el monto a validar es el diferencial (saldo pendiente).
 */
function construirPagosProgramados_(agrupadas, vcMap) {
  vcMap = vcMap || {};
  const items = [];
  agrupadas.forEach(g => {
    const pendiente = g.pendienteOC || 0, pagado = g.pagado || 0, totalOC = pendiente + pagado;
    if (g.valorizado) {
      const cuotas = [1, 2, 3, 4, 5]
        .map(n => ({ n: n, fecha: g['fechaPago' + n], monto: Number(g['montoPago' + n]) || 0 }))
        .filter(c => c.fecha && c.monto > 0);
      cuotas.forEach(c => {
        const v = vcMap[String(g.oc) + '|' + c.n] || {};
        items.push({ g: g, oc: g.oc, cuota: c.n, totalCuotas: cuotas.length, esCuota: true,
          tipo: 'Cuota ' + c.n + ' de ' + cuotas.length,
          monto: c.monto, fechaPago: c.fecha, fechaValidada: v.fechaValidada || '', estado: v.estado || '',
          comentario: v.comentario || '', pagado: pagado, totalOC: totalOC });
      });
    } else if (g.fechaPago) {
      const parcial = pagado > 0 && pendiente > 0;
      items.push({ g: g, oc: g.oc, cuota: 0, totalCuotas: 1, esCuota: false,
        tipo: parcial ? 'Pago parcial (diferencial)' : (pendiente > 0 ? 'Pago total' : 'Pagado'),
        monto: pendiente > 0 ? pendiente : pagado,
        fechaPago: g.fechaPago, fechaValidada: g.fechaPagoValidada || '', estado: g.estadoValidacion || '',
        comentario: '', pagado: pagado, totalOC: totalOC });
    }
  });
  return items;
}


/* ====================================================================
   PAGOS REGISTRADOS POR MES: lo pagado de enero hasta el mes en curso.
   Toma las líneas con Monto Pagado > 0 y las ubica en el mes de su "Fecha de documento"
   (columna B de la base). Respeta los filtros. Resultado en caché (se limpia al sincronizar).
   ==================================================================== */
function bumpCacheVer_() {
  try { PropertiesService.getScriptProperties().setProperty('CACHE_VER', String(Date.now())); } catch (e) {}   // nunca debe impedir que se guarde un dato
}
function hashCorto_(txt) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, txt);
  return b.map(x => ('0' + (x & 0xff).toString(16)).slice(-2)).join('');
}

/**
 * Año/mes de "ahora" para el negocio: se basa en CONFIG.FECHA_CORTE (dd/MM/yyyy), NO en el reloj
 * real del servidor. Así "Pagos registrados por mes" y los rangos de fecha del Cash Flow siguen
 * funcionando con el año de tus datos aunque el reloj real de Google no coincida.
 */
function obtenerFechaNegocioActual_() {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(CONFIG.FECHA_CORTE || '').trim());
  if (m) return { anio: Number(m[3]), mes: Number(m[2]) };
  const hoy = new Date();
  return { anio: hoy.getFullYear(), mes: hoy.getMonth() + 1 };
}

function obtenerPagosPorMes(filtros) {
  filtros = filtros || {};
  const tz = Session.getScriptTimeZone();
  const fn_ = obtenerFechaNegocioActual_();
  const anioActual = fn_.anio, mesActual = fn_.mes;
  let cache = null, clave = null;
  try {
    cache = CacheService.getScriptCache();
    const ver = PropertiesService.getScriptProperties().getProperty('CACHE_VER') || '0';
    clave = 'pagosmes_' + ver + '_' + anioActual + '_' + mesActual + '_' + hashCorto_(JSON.stringify(filtros));
    const hit = cache.get(clave);
    if (hit) return JSON.parse(hit);
  } catch (e) { cache = null; }
  const r = calcularPagosPorMes_(filtros, anioActual, mesActual, tz);
  try { if (cache && clave) cache.put(clave, JSON.stringify(r), 900); } catch (e) {}
  return r;
}

/** Convierte a Date lo que venga en "Fecha de documento": objeto Date real, texto ISO (yyyy-MM-dd),
 * o texto dd/MM/yyyy (o dd-MM-yyyy) — este último es el que rompía silenciosamente antes: si la
 * celda llegaba como texto en vez de fecha real (pasa seguido cuando el IMPORTRANGE de un dominio
 * usa un formato de columna distinto al de otro), `new Date("24/09/2026")` da "Invalid Date" y esa
 * fila caía sin avisar en "sin fecha" — por eso un dominio entero podía faltar en el gráfico aunque
 * sí tuviera pagos. Devuelve null si de verdad no se puede interpretar. */
function parsearFechaDocumento_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : v;
  const s = String(v).trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);              // yyyy-MM-dd
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/.exec(s);         // dd/MM/yyyy o dd-MM-yyyy
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function calcularPagosPorMes_(filtros, anioActual, mesActual, tz) {
  const anio = Number(filtros.pagosAnio) || anioActual;
  const mesDefecto = anio === anioActual ? mesActual : (anio < anioActual ? 12 : 1);
  const mesHasta = Math.min(12, Math.max(1, Number(filtros.pagosMesHasta) || mesDefecto));
  const NOMBRES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const meses = [];
  for (let m = 1; m <= mesHasta; m++) meses.push({ mes: m, etiqueta: NOMBRES[m - 1], total: 0, porDominio: {}, lineas: 0 });
  const vacio = { anio: anio, mesHasta: mesHasta, mesActual: anio === anioActual ? mesActual : 0, meses: meses,
                  resumen: { totalGeneral: 0, enRango: 0, otrosPeriodos: 0, sinFecha: 0, sinFechaPorDominio: {}, conRespaldo: 0 } };

  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return vacio;
  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();

  const ok = {};
  agruparPorOC_(datos).forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.estado, g.estado)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (!coincideFiltro_(filtros.dominioSp, (g.dominioSp || '').trim() || '(Sin dato)')) return;
    ok[String(g.oc)] = g;
  });

  const R = vacio.resumen;
  datos.forEach(fila => {
    const pagado = Number(fila[14]) || 0;                          // Monto Pagado (col 15)
    if (pagado <= 0) return;
    const g = ok[String(fila[0])];
    if (!g) return;
    R.totalGeneral += pagado;
    let d = parsearFechaDocumento_(fila[1]);                       // Fecha de documento (col B)
    let deRespaldo = false;
    if (!d) { d = parsearFechaDocumento_(fila[COL_BASE.FECHA_PAGO_VALIDADA - 1]); deRespaldo = !!d; }   // sin fecha de documento: usa la Fecha de pago validada si existe
    if (!d) { R.sinFecha += pagado; const domSF = g.dominio || 'Sin dominio'; R.sinFechaPorDominio[domSF] = (R.sinFechaPorDominio[domSF] || 0) + pagado; return; }
    if (deRespaldo) R.conRespaldo += pagado;
    const a = Number(Utilities.formatDate(d, tz, 'yyyy')), m = Number(Utilities.formatDate(d, tz, 'M'));
    if (a === anio && m >= 1 && m <= mesHasta) {
      const mm = meses[m - 1];
      const dom = g.dominio || 'Sin dominio';
      mm.total += pagado; mm.lineas++;
      mm.porDominio[dom] = (mm.porDominio[dom] || 0) + pagado;
      R.enRango += pagado;
    } else {
      R.otrosPeriodos += pagado;
    }
  });
  return vacio;
}

/* ====================================================================
   ELIMINAR OC CON APROBACIÓN
   Cualquier perfil solicita (con motivo). Un Maestro/admin aprueba o rechaza.
   Al aprobar, la OC queda con Estado manual = "Eliminado" y sale de TODO el portal
   (Detalle, Cash Flow, gráficos, KPIs). Se puede restaurar desde el panel de Maestro.
   Registro en la hoja "SolicitudesEliminacion".
   ==================================================================== */
const HOJA_SOL_ELIM = 'SolicitudesEliminacion';

function obtenerHojaSolElim_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA_SOL_ELIM);
  if (!h) {
    h = ss.insertSheet(HOJA_SOL_ELIM);
    h.getRange(1, 1, 1, 8).setValues([['N° OC', 'Solicitante', 'Fecha solicitud', 'Motivo', 'Estado', 'Resuelto por', 'Fecha resolución', 'Comentario']]);
    h.setFrozenRows(1);
  }
  return h;
}
function puedeAprobarEliminacion_() {
  const u = obtenerSubRolVisitante();
  return !!(u && u.esMaestro) || esUsuarioAdmin();
}
function fechaHoraTxt_(d) {
  if (!d) return '';
  return Object.prototype.toString.call(d) === '[object Date]' ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(d);
}
function leerSolicitudesEliminacion_() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_SOL_ELIM);
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, 8).getValues().map((f, i) => ({
    fila: i + 2, oc: String(f[0]), solicitante: f[1] || '', fecha: f[2], motivo: f[3] || '', estado: String(f[4] || ''),
    resueltoPor: f[5] || '', fechaRes: f[6], comentario: f[7] || '',
  })).filter(x => x.oc !== '');
}
function leerSolicitudesEliminacionPendientes_() {
  const mapa = {};
  leerSolicitudesEliminacion_().forEach(x => { if (x.estado === 'Pendiente') mapa[x.oc] = x; });
  return mapa;
}
/** Escribe el Estado manual en todas las líneas de la OC en BASE_CASHFLOW. */
function establecerEstadoManualOC_(numeroOC, valor) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hoja || hoja.getLastRow() < 2) return 0;
  const col = hoja.getRange(2, COL_BASE.OC, hoja.getLastRow() - 1, 1).getValues();
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    if (String(col[i][0]) !== String(numeroOC)) continue;
    hoja.getRange(i + 2, COL_BASE.ESTADO_MANUAL).setValue(valor);
    n++;
  }
  bumpCacheVer_();
  return n;
}

function solicitarEliminacionOC(numeroOC, motivo) {
  const u = obtenerSubRolVisitante();
  if (obtenerAlertas().bloquearEdicion && !(u && u.esMaestro)) return { ok: false, error: '🔒 El portal está en modo solo lectura. Contacta a un perfil Maestro.' };
  motivo = String(motivo || '').trim();
  if (motivo.length < 3) return { ok: false, error: 'Indica el motivo de la eliminación.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (leerSolicitudesEliminacionPendientes_()[String(numeroOC)]) return { ok: true, pendiente: true, mensaje: 'Esta OC ya tenía una solicitud pendiente.' };
    const h = obtenerHojaSolElim_();
    const correo = obtenerCorreoVisitante_() || '';
    h.appendRow([String(numeroOC), correo, new Date(), motivo, 'Pendiente', '', '', '']);
    if (puedeAprobarEliminacion_()) {           // quien aprueba es quien pide: se ejecuta directo, queda registrado
      const fila = h.getLastRow();
      establecerEstadoManualOC_(numeroOC, 'Eliminado');
      h.getRange(fila, 5, 1, 4).setValues([['Aprobada', correo, new Date(), 'Aprobación directa']]);
      return { ok: true, aprobada: true };
    }
    return { ok: true, pendiente: true };
  } finally {
    lock.releaseLock();
  }
}

function cancelarSolicitudEliminacion(numeroOC) {
  const sol = leerSolicitudesEliminacionPendientes_()[String(numeroOC)];
  if (!sol) return { ok: true };
  const correo = String(obtenerCorreoVisitante_() || '').toLowerCase();
  if (String(sol.solicitante).toLowerCase() !== correo && !puedeAprobarEliminacion_())
    return { ok: false, error: 'Solo quien la solicitó o un Maestro puede cancelar la solicitud.' };
  obtenerHojaSolElim_().getRange(sol.fila, 5, 1, 4).setValues([['Cancelada', correo, new Date(), '']]);
  return { ok: true };
}

/** OCs involucradas en el panel de eliminaciones (pendientes + últimas 15 aprobadas). */
function solicitudesEliminacionParaPanel_() {
  const sols = leerSolicitudesEliminacion_();
  return { pend: sols.filter(x => x.estado === 'Pendiente'), aprobadas: sols.filter(x => x.estado === 'Aprobada').reverse().slice(0, 15) };
}
function mapearEliminaciones_(sp, info) {
  const dato = x => ({ fila: x.fila, oc: x.oc, solicitante: x.solicitante, fecha: fechaHoraTxt_(x.fecha), motivo: x.motivo,
                       proyecto: (info[x.oc] || {}).proyecto || '', monto: (info[x.oc] || {}).saldo || 0,
                       resueltoPor: x.resueltoPor, fechaRes: fechaHoraTxt_(x.fechaRes) });
  return { ok: true, pendientes: sp.pend.map(dato), recientes: sp.aprobadas.filter(x => (info[x.oc] || {}).eliminada).map(dato) };
}
/** Para el panel de Maestro: pendientes de aprobar + últimas eliminaciones aprobadas (con opción de restaurar). */
function obtenerSolicitudesEliminacion() {
  if (!puedeAprobarEliminacion_()) return { ok: false, pendientes: [], recientes: [] };
  const sp = solicitudesEliminacionParaPanel_();
  const set = {};
  sp.pend.concat(sp.aprobadas).forEach(x => { set[x.oc] = true; });
  return mapearEliminaciones_(sp, infoLigeraOCs_(set));
}

function resolverEliminaciones(filas, aprobar, comentario) {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un Maestro puede aprobar o rechazar eliminaciones.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const h = obtenerHojaSolElim_();
    const correo = obtenerCorreoVisitante_() || '';
    let count = 0;
    (filas || []).forEach(fila => {
      const f = h.getRange(fila, 1, 1, 8).getValues()[0];
      if (String(f[4]) !== 'Pendiente') return;
      if (aprobar) establecerEstadoManualOC_(f[0], 'Eliminado');
      h.getRange(fila, 5, 1, 4).setValues([[aprobar ? 'Aprobada' : 'Rechazada', correo, new Date(), comentario || '']]);
      count++;
    });
    return { ok: true, count: count };
  } finally {
    lock.releaseLock();
  }
}

function restaurarOCEliminada(filaSolicitud) {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un Maestro puede restaurar una OC.' };
  const h = obtenerHojaSolElim_();
  const f = h.getRange(filaSolicitud, 1, 1, 8).getValues()[0];
  establecerEstadoManualOC_(f[0], '');
  h.getRange(filaSolicitud, 5, 1, 4).setValues([['Restaurada', obtenerCorreoVisitante_() || '', new Date(), '']]);
  return { ok: true };
}


/* ====================================================================
   ÓRDENES SIN AGENDAR (por área)
   Se toman las OC con saldo pendiente (Monto pendiente / OC > 0):
   - AGENDADA   = tiene Fecha de pago propuesta, o está valorizada con al menos una cuota con fecha y monto.
   - SIN AGENDAR = el resto. Monto = saldo pendiente de la OC.
   ==================================================================== */
function estaAgendada_(g) {
  if (g.fechaPago) return true;
  if (!g.valorizado) return false;
  for (let n = 1; n <= 5; n++) {
    if (g['fechaPago' + n] && Number(g['montoPago' + n]) > 0) return true;
  }
  return false;
}

function calcularAgendamientoPorArea_(filtros) {
  filtros = filtros || {};
  const vacio = { total: { ocs: 0, agendadas: 0, montoAgendado: 0, noAgendadas: 0, montoNoAgendado: 0, pctAgendado: 0 }, areas: [] };
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return vacio;
  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();

  const T = vacio.total;
  const porArea = {};
  agruparPorOC_(datos).forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.estado, g.estado)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (!coincideFiltro_(filtros.dominioSp, (g.dominioSp || '').trim() || '(Sin dato)')) return;
    const saldo = Number(g.pendienteOC) || 0;
    if (saldo <= 0) return;                                   // ya pagada / sin saldo: no necesita agenda
    const nombre = String(g.area || '').trim() || '(Sin área)';
    const A = porArea[nombre] || (porArea[nombre] = { area: nombre, agendadas: 0, montoAgendado: 0, noAgendadas: 0, montoNoAgendado: 0 });
    T.ocs++;
    if (estaAgendada_(g)) { A.agendadas++; A.montoAgendado += saldo; T.agendadas++; T.montoAgendado += saldo; }
    else { A.noAgendadas++; A.montoNoAgendado += saldo; T.noAgendadas++; T.montoNoAgendado += saldo; }
  });

  const pct = (a, b) => (a + b) > 0 ? (a / (a + b)) * 100 : 0;
  T.pctAgendado = pct(T.montoAgendado, T.montoNoAgendado);
  vacio.areas = Object.keys(porArea).map(k => {
    const A = porArea[k];
    A.pctAgendado = pct(A.montoAgendado, A.montoNoAgendado);
    return A;
  }).sort((a, b) => (b.montoNoAgendado - a.montoNoAgendado) || a.area.localeCompare(b.area));   // primero las áreas a las que más falta
  return vacio;
}

/**
 * Igual que calcularAgendamientoPorArea_ pero además arma el detalle de OC sin agendar (para el
 * correo) y una lista aparte con las OC del Trimestre 1Q que TODAVÍA tienen saldo pendiente de pago
 * — esas ya deberían estar cerradas a esta altura del año, así que se destacan aparte sin importar
 * si están agendadas o no.
 */
function calcularResumenAgendamientoParaCorreo_(filtros) {
  filtros = filtros || {};
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return { porArea: [], total: { agendadas: 0, montoAgendado: 0, noAgendadas: 0, montoNoAgendado: 0 }, prioridad1Q: [] };
  const datos = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();

  const porArea = {};
  const prioridad1Q = [];
  agruparPorOC_(datos).forEach(g => {
    if (!coincideFiltro_(filtros.proyecto, g.proyecto)) return;
    if (!coincideFiltro_(filtros.sdatool, g.sdatool)) return;
    if (!coincideFiltro_(filtros.estado, g.estado)) return;
    if (!coincideFiltro_(filtros.area, g.area)) return;
    if (!coincideFiltro_(filtros.trimestre, g.trimestre)) return;
    if (!coincideFiltro_(filtros.dominio, g.dominio)) return;
    if (!coincideFiltro_(filtros.dominioSp, (g.dominioSp || '').trim() || '(Sin dato)')) return;

    const saldo = Number(g.pendienteOC) || 0;
    const agendada = estaAgendada_(g);

    // Prioridad: 1Q con saldo pendiente, esté o no agendada — es lo más atrasado del año.
    if (String(g.trimestre || '').trim() === '1Q' && saldo > 0) {
      // Las OC valorizadas pueden tener hasta 5 cuotas, cada una con su propia fecha y monto —
      // por eso antes salía una raya (—): se buscaba una sola "Fecha de pago" que esas OC no tienen.
      let cuotas = [];
      if (g.valorizado) {
        for (let n = 1; n <= 5; n++) {
          const fc = g['fechaPago' + n], mc = Number(g['montoPago' + n]) || 0;
          if (fc && mc > 0) cuotas.push({ fecha: formatearFecha(fc), monto: mc });
        }
      }
      prioridad1Q.push({
        oc: g.oc, proyecto: g.proyecto || '', area: g.area || '(Sin área)', oficina: g.oficina || '',
        monto: saldo, agendada: agendada, valorizado: !!g.valorizado, cuotas: cuotas,
        fechaPago: (agendada && !g.valorizado) ? formatearFecha(g.fechaPago) : '',
      });
    }

    if (saldo <= 0) return;   // ya pagada: no necesita agenda, no entra a la tabla por área
    const nombre = String(g.area || '').trim() || '(Sin área)';
    const A = porArea[nombre] || (porArea[nombre] = { area: nombre, agendadas: 0, montoAgendado: 0, noAgendadas: 0, montoNoAgendado: 0, ocsSinAgendar: [], ocsDetalle: [] });
    A.ocsDetalle.push({ oc: g.oc, proyecto: g.proyecto || '', oficina: g.oficina || '', monto: saldo, agendada: agendada, fechaPago: agendada ? formatearFecha(g.fechaPago) : '' });
    if (agendada) { A.agendadas++; A.montoAgendado += saldo; }
    else { A.noAgendadas++; A.montoNoAgendado += saldo; A.ocsSinAgendar.push({ oc: g.oc, proyecto: g.proyecto || '', monto: saldo }); }
  });

  const listaAreas = Object.keys(porArea).map(k => porArea[k]).sort((a, b) => b.montoNoAgendado - a.montoNoAgendado);
  listaAreas.forEach(a => { a.ocsDetalle.sort((x, y) => (x.agendada - y.agendada) || (y.monto - x.monto)); });   // sin agendar primero, luego por monto
  const total = listaAreas.reduce((acc, a) => {
    acc.agendadas += a.agendadas; acc.montoAgendado += a.montoAgendado;
    acc.noAgendadas += a.noAgendadas; acc.montoNoAgendado += a.montoNoAgendado;
    return acc;
  }, { agendadas: 0, montoAgendado: 0, noAgendadas: 0, montoNoAgendado: 0 });

  prioridad1Q.sort((a, b) => b.monto - a.monto);
  return { porArea: listaAreas, total: total, prioridad1Q: prioridad1Q };
}

/**
 * Paso 1 (no envía nada): arma la misma info que el correo, para mostrarla en un modal de
 * confirmación antes de disparar el envío — mismo patrón que "Importar fechas de Hans" y
 * "Sincronizar OC".
 */
function previsualizarResumenAgendamiento(filtros) {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede enviar este resumen.' };
  const r = calcularResumenAgendamientoParaCorreo_(filtros || {});
  return {
    ok: true,
    porArea: r.porArea.map(a => ({ area: a.area, agendadas: a.agendadas, montoAgendado: a.montoAgendado, noAgendadas: a.noAgendadas, montoNoAgendado: a.montoNoAgendado })),
    total: r.total,
    prioridad1Q: r.prioridad1Q.slice(0, 150),
    cantPrioridad1Q: r.prioridad1Q.length,
    montoPrioridad1Q: r.prioridad1Q.reduce((s, x) => s + x.monto, 0),
    destinatarios: obtenerCorreosMaestros_().length,
  };
}

/**
 * Envía el resumen de OC pendientes de agendar por Área (con la sección de prioridad 1Q sin pagar)
 * SOLO a los perfiles Maestro — este tipo de comunicación de seguimiento siempre va nada más a
 * Maestro, nunca a Líder ni a otros perfiles. Solo un Maestro puede disparar el envío.
 */
/** Devuelve la URL del web app publicado (para armar enlaces "Ver en el portal" en los correos). */
function obtenerUrlPortal_() {
  try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; }
}

function enviarResumenAgendamientoPorAreaCorreo(filtros) {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede enviar este resumen.' };
  const destinos = obtenerCorreosMaestros_();
  if (!destinos.length) return { ok: false, error: 'No hay perfiles Maestro con correo cargado en la pestaña Usuarios.' };

  const r = calcularResumenAgendamientoParaCorreo_(filtros || {});
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const fmtM = v => 'S/ ' + Math.round(Number(v) || 0).toLocaleString('es-PE');
  const pct = (a, b) => (a + b) > 0 ? Math.round(a / (a + b) * 100) : 0;
  const urlPortal = obtenerUrlPortal_();
  const linkArea = area => urlPortal ? `${urlPortal}?area=${encodeURIComponent(area)}&sinFecha=1` : '';
  const link1Q = urlPortal ? `${urlPortal}?trimestre=1Q&pendientePago=1` : '';
  const enlace = (url, texto) => url ? `<a href="${url}" style="color:#2E6FBE;font-size:11.5px;font-weight:600;text-decoration:none;">${texto} →</a>` : '';

  const filaArea = a => `<tr>
      <td style="padding:6px 10px;">${escHtmlSrv_(a.area)}</td>
      <td style="padding:6px 10px;text-align:right;">${a.noAgendadas}</td>
      <td style="padding:6px 10px;text-align:right;">${fmtM(a.montoNoAgendado)}</td>
      <td style="padding:6px 10px;text-align:right;">${a.agendadas}</td>
      <td style="padding:6px 10px;text-align:right;">${fmtM(a.montoAgendado)}</td>
      <td style="padding:6px 10px;text-align:right;font-weight:700;">${pct(a.montoAgendado, a.montoNoAgendado)}%</td>
    </tr>`;
  // Solo las OC SIN agendar entran al detalle por área (las agendadas ya no necesitan seguimiento aquí).
  const filaOCDetalle = it => `<tr>
      <td style="padding:4px 10px;">${it.oc}</td>
      <td style="padding:4px 10px;">${escHtmlSrv_(it.proyecto)}</td>
      <td style="padding:4px 10px;">${escHtmlSrv_(it.oficina) || '—'}</td>
      <td style="padding:4px 10px;text-align:right;">${fmtM(it.monto)}</td>
    </tr>`;
  const seccionPorArea = a => {
    const sinAgendar = a.ocsDetalle.filter(x => !x.agendada);
    if (!sinAgendar.length) return '';
    return `
    <div style="display:flex;justify-content:space-between;align-items:baseline;margin:16px 0 6px;">
      <h4 style="font-size:12.5px;margin:0;color:#0B1D33;">${escHtmlSrv_(a.area)} — ${sinAgendar.length} OC pendientes de agendar</h4>
      ${enlace(linkArea(a.area), 'Ver en el portal')}
    </div>
    <table style="border-collapse:collapse;font-size:12px;width:100%;margin-bottom:6px;">
      <tr style="color:#6B7A99;font-size:10.5px;text-align:left;background:#FAFBFC;"><td style="padding:4px 10px;">N° OC</td><td style="padding:4px 10px;">Proyecto</td><td style="padding:4px 10px;">Oficina</td><td style="padding:4px 10px;text-align:right;">Monto</td></tr>
      ${sinAgendar.slice(0, 25).map(filaOCDetalle).join('')}
    </table>
    ${sinAgendar.length > 25 ? `<p style="font-size:11px;color:#6B7A99;margin:0 0 6px;">… y ${sinAgendar.length - 25} OC más de esta área. Usa el enlace de arriba para verlas todas.</p>` : ''}
  `;
  };
  const fila1Q = it => {
    let celdaFecha;
    if (it.valorizado && it.cuotas.length) {
      celdaFecha = it.cuotas.map((c, i) => `Cuota ${i+1}: ${c.fecha} (${fmtM(c.monto)})`).join('<br>');
    } else if (it.agendada) {
      celdaFecha = it.fechaPago || '—';
    } else {
      celdaFecha = '<span style="color:#B98A3D;font-weight:600;">Falta agendar</span>';
    }
    return `<tr>
      <td style="padding:5px 10px;">${it.oc}</td><td style="padding:5px 10px;">${escHtmlSrv_(it.proyecto)}</td>
      <td style="padding:5px 10px;">${escHtmlSrv_(it.area)}</td>
      <td style="padding:5px 10px;">${celdaFecha}</td>
      <td style="padding:5px 10px;text-align:right;">${fmtM(it.monto)}</td>
    </tr>`;
  };

  const cuerpo = `
    <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">
      <div style="background:#0B1D33;color:#fff;padding:18px 22px;border-radius:8px 8px 0 0;">
        <div style="font-size:12px;opacity:.75;">Portal CAPEX P&S · Detalle OC</div>
        <h2 style="margin:6px 0 0;font-size:17px;">📅 Resumen de OC pendientes de agendar — ${fecha}</h2>
      </div>
      <div style="background:#fff;padding:20px 22px;border:1px solid #e0e0e0;border-top:4px solid #2E6FBE;">
        <table style="border-collapse:collapse;font-size:13px;width:100%;margin-bottom:8px;">
          <tr style="color:#6B7A99;font-size:11px;text-align:left;background:#F7F9FB;">
            <td style="padding:6px 10px;">Área</td><td style="padding:6px 10px;text-align:right;">OC sin agendar</td><td style="padding:6px 10px;text-align:right;">Monto sin agendar</td>
            <td style="padding:6px 10px;text-align:right;">OC agendadas</td><td style="padding:6px 10px;text-align:right;">Monto agendado</td><td style="padding:6px 10px;text-align:right;">% avance</td>
          </tr>
          ${r.porArea.map(filaArea).join('')}
          <tr style="font-weight:700;background:#EEF3FA;">
            <td style="padding:6px 10px;">Total</td><td style="padding:6px 10px;text-align:right;">${r.total.noAgendadas}</td><td style="padding:6px 10px;text-align:right;">${fmtM(r.total.montoNoAgendado)}</td>
            <td style="padding:6px 10px;text-align:right;">${r.total.agendadas}</td><td style="padding:6px 10px;text-align:right;">${fmtM(r.total.montoAgendado)}</td><td style="padding:6px 10px;text-align:right;">${pct(r.total.montoAgendado, r.total.montoNoAgendado)}%</td>
          </tr>
        </table>
        <p style="font-size:11px;color:#6B7A99;margin:0 0 4px;">Debajo, el detalle de las OC que todavía faltan agendar, por área (las ya agendadas no se listan aquí — sí se cuentan en la tabla de arriba):</p>
        <div style="border-top:1px solid #e0e0e0;padding-top:6px;">
          ${r.porArea.map(seccionPorArea).join('')}
        </div>
        <div style="display:flex;justify-content:space-between;align-items:baseline;margin:18px 0 6px;">
          <h3 style="color:#C0392B;font-size:13px;margin:0;">⚠️ Prioridad: OC de 1Q que todavía no se han pagado${r.prioridad1Q.length ? ' (' + r.prioridad1Q.length + ' · ' + fmtM(r.prioridad1Q.reduce((s,x)=>s+x.monto,0)) + ')' : ''}</h3>
          ${r.prioridad1Q.length ? enlace(link1Q, 'Ver en el portal') : ''}
        </div>
        ${r.prioridad1Q.length ? `
        <table style="border-collapse:collapse;font-size:13px;width:100%;">
          <tr style="color:#6B7A99;font-size:11px;text-align:left;background:#FBF3F2;"><td style="padding:5px 10px;">N° OC</td><td style="padding:5px 10px;">Proyecto</td><td style="padding:5px 10px;">Área</td><td style="padding:5px 10px;">Fecha agendada</td><td style="padding:5px 10px;text-align:right;">Saldo pendiente</td></tr>
          ${r.prioridad1Q.slice(0, 40).map(fila1Q).join('')}
        </table>
        ${r.prioridad1Q.length > 40 ? `<p style="font-size:11.5px;color:#6B7A99;margin-top:8px;">… y ${r.prioridad1Q.length - 40} más. Usa el enlace de arriba para verlas todas.</p>` : ''}
        ` : '<p style="font-size:13px;color:#4C9270;">✅ No hay OC de 1Q con saldo pendiente de pago.</p>'}
      </div>
      <div style="background:#f0f4f8;padding:10px 22px;border-radius:0 0 8px 8px;font-size:11px;color:#6B7A99;">
        Enviado solo a perfiles Maestro · Portal CAPEX P&S
      </div>
    </div>`;

  MailApp.sendEmail({
    to: destinos.join(','),
    subject: `[Portal CAPEX] Resumen de agendamiento — ${r.total.noAgendadas} OC sin agendar (${fmtM(r.total.montoNoAgendado)})` + (r.prioridad1Q.length ? ` · ${r.prioridad1Q.length} de 1Q sin pagar` : ''),
    htmlBody: cuerpo,
  });

  return { ok: true, destinatarios: destinos.length, noAgendadas: r.total.noAgendadas, prioridad1Q: r.prioridad1Q.length };
}


/* ====================================================================
   PESTAÑA "APROBACIONES" (solo Maestro): todo lo pendiente en UNA llamada rápida.
   Lee únicamente las hojas de solicitudes (pequeñas) y 4 columnas de la base solo para las OC involucradas.
   ==================================================================== */
function infoLigeraOCs_(setOCs) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BASE);
  const info = {};
  if (!hoja || hoja.getLastRow() < 2) return info;
  const n = hoja.getLastRow() - 1;
  const ocs = hoja.getRange(2, COL_BASE.OC, n, 1).getValues();
  const proy = hoja.getRange(2, 4, n, 1).getValues();                       // Proyecto (col 4)
  const pend = hoja.getRange(2, 14, n, 1).getValues();                      // Monto pendiente / OC (col 14)
  const man = hoja.getRange(2, COL_BASE.ESTADO_MANUAL, n, 1).getValues();   // Estado manual (col 35)
  for (let i = 0; i < n; i++) {
    const k = String(ocs[i][0]);
    if (!setOCs[k]) continue;
    const it = info[k] || (info[k] = { proyecto: '', saldo: 0, eliminada: false });
    if (!it.proyecto) it.proyecto = proy[i][0] || '';
    it.saldo += Number(pend[i][0]) || 0;
    if (String(man[i][0]) === 'Eliminado') it.eliminada = true;
  }
  return info;
}

function obtenerAprobaciones() {
  const vacio = { ok: false, error: 'Solo un perfil Maestro puede ver las aprobaciones.', cambiosFecha: [], eliminaciones: { pendientes: [], recientes: [] } };
  if (!puedeAprobarEliminacion_()) return vacio;
  const raw = leerCambiosFechaPendientesRaw_();
  const sp = solicitudesEliminacionParaPanel_();
  const set = {};
  raw.forEach(c => { set[String(c.oc)] = true; });
  sp.pend.concat(sp.aprobadas).forEach(x => { set[x.oc] = true; });
  const info = Object.keys(set).length ? infoLigeraOCs_(set) : {};    // UNA sola lectura (4 columnas) para ambas listas
  return { ok: true, cambiosFecha: mapearCambiosFecha_(raw, info), eliminaciones: mapearEliminaciones_(sp, info) };
}

/** Solo cuenta (rápido): para el globo de la pestaña. */
function obtenerContadorAprobaciones() {
  if (!puedeAprobarEliminacion_()) return { ok: false, cf: 0, el: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cf = 0, el = 0;
  const h1 = ss.getSheetByName('HistorialFechas');
  if (h1 && h1.getLastRow() >= 2) {
    h1.getRange(2, 1, h1.getLastRow() - 1, 6).getValues().forEach(f => {
      const est = String(f[5] || '').trim();
      if (f[0] !== '' && (est === 'Pendiente' || est === '')) cf++;
    });
  }
  const h2 = ss.getSheetByName(HOJA_SOL_ELIM);
  if (h2 && h2.getLastRow() >= 2) {
    h2.getRange(2, 1, h2.getLastRow() - 1, 5).getValues().forEach(f => { if (f[0] !== '' && String(f[4]) === 'Pendiente') el++; });
  }
  return { ok: true, cf: cf, el: el };
}


/* ====================================================================
   CACHÉ DE RESULTADOS PESADOS (Validación de pagos, Gantt, OC sin agendar)
   Guarda el resultado unos minutos; cualquier edición o sincronización cambia CACHE_VER y lo invalida.
   Los resultados grandes se parten en trozos (cada valor de caché admite ~100 KB).
   ==================================================================== */
function conCache_(nombre, filtros, calcular) {
  let cache = null, base = null;
  try {
    cache = CacheService.getScriptCache();
    const ver = PropertiesService.getScriptProperties().getProperty('CACHE_VER') || '0';
    base = 'c_' + nombre + '_' + ver + '_' + hashCorto_(JSON.stringify(filtros || {}));
    const n = Number(cache.get(base + '_n'));
    if (n > 0) {
      const claves = [];
      for (let i = 0; i < n; i++) claves.push(base + '_' + i);
      const partes = cache.getAll(claves);
      let txt = '', completo = true;
      for (let i = 0; i < n; i++) { const t = partes[claves[i]]; if (t == null) { completo = false; break; } txt += t; }
      if (completo) return JSON.parse(txt);
    }
  } catch (e) { cache = null; }
  const r = calcular();
  try {
    if (cache && base) {
      const txt = JSON.stringify(r);
      if (txt.length <= 1800000) {
        const TAM = 30000, put = {};
        let n = 0;
        for (let i = 0; i < txt.length; i += TAM) put[base + '_' + (n++)] = txt.slice(i, i + TAM);
        put[base + '_n'] = String(n);
        cache.putAll(put, 600);
      }
    }
  } catch (e) {}
  return r;
}
function obtenerGantt(filtros) { return conCache_('gantt', filtros, () => calcularGantt_(filtros)); }
function obtenerValidacionPagos(filtros) { return conCache_('valpagos', filtros, () => calcularValidacionPagos_(filtros)); }
function obtenerAgendamientoPorArea(filtros) { return conCache_('agend', filtros, () => calcularAgendamientoPorArea_(filtros)); }

/** Semanas (lunes a domingo) que tienen pagos programados, para el filtro de Validación. */
function semanasDePagos_(items) {
  const tz = Session.getScriptTimeZone();
  const lunes = {};
  items.forEach(it => {
    const v = it.fechaPago;
    if (!v) return;
    const ymd = Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).slice(0, 10);
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ymd);
    if (!m) return;
    const t = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    lunes[t - ((new Date(t).getUTCDay() + 6) % 7) * 86400000] = true;
  });
  const dm = t => { const d = new Date(t); return ('0' + d.getUTCDate()).slice(-2) + '/' + ('0' + (d.getUTCMonth() + 1)).slice(-2); };
  return Object.keys(lunes).map(Number).sort((a, b) => a - b).map(t => ({
    inicio: new Date(t).toISOString().slice(0, 10),
    etiqueta: dm(t) + ' – ' + dm(t + 6 * 86400000),
  }));
}


/* ====================================================================
   REPORTE SEMANAL DE VALIDACIÓN — control desde el portal (solo Maestro)
   - Activar/desactivar el envío automático: cada lunes ~9:00 a.m. (hora de Lima).
   - Enviar ahora: manda el mismo reporte al instante.
   El correo sale de la cuenta con la que se ejecuta el portal y llega a los perfiles Maestro de la pestaña Usuarios.
   ==================================================================== */
const REPORTE_TZ = 'America/Lima';
const REPORTE_HANDLER = 'enviarCorreoSemanalValidacion';

function reporteTriggers_() {
  return ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === REPORTE_HANDLER);
}

/** Próximo lunes 9:00 (hora de Lima, UTC-5 fijo) como instante real. `ahora` es opcional (para pruebas). */
function proximoLunes9_(ahora) {
  const ahoraMs = (ahora || new Date()).getTime();
  const lima = new Date(ahoraMs - 5 * 3600000);                       // "hora de Lima" leída como UTC
  let dias = (1 - lima.getUTCDay() + 7) % 7;                         // días hasta el lunes (0 si hoy es lunes)
  const objetivo = d => Date.UTC(lima.getUTCFullYear(), lima.getUTCMonth(), lima.getUTCDate() + d, 9, 0);
  if (objetivo(dias) <= lima.getTime()) dias += 7;                    // hoy lunes, pero ya pasaron las 9:00
  return new Date(objetivo(dias) + 5 * 3600000);
}

function obtenerEstadoReporteSemanal() {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede ver esta configuración.' };
  const activo = reporteTriggers_().length > 0;
  let ultimo = null;
  try { ultimo = JSON.parse(PropertiesService.getScriptProperties().getProperty('REPORTE_ULTIMO') || 'null'); } catch (e) {}
  // Semana que cubre el reporte: la semana anterior (lunes a viernes)
  const hoy = new Date();
  const lunes = new Date(hoy); lunes.setDate(hoy.getDate() - ((hoy.getDay() + 6) % 7)); lunes.setHours(0, 0, 0, 0);
  const l = new Date(lunes); l.setDate(lunes.getDate() - 7);
  const v = new Date(l); v.setDate(l.getDate() + 4);
  const f = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return {
    ok: true, activo: activo,
    proximo: activo ? Utilities.formatDate(proximoLunes9_(), REPORTE_TZ, 'dd/MM/yyyy') + ' · 9:00 a.m.' : '',
    destinatarios: obtenerCorreosMaestros_().length,
    semana: f(l) + ' – ' + f(v),
    ultimo: ultimo,
  };
}

function activarReporteSemanal() {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede activar el reporte.' };
  reporteTriggers_().forEach(t => ScriptApp.deleteTrigger(t));         // evita duplicados
  ScriptApp.newTrigger(REPORTE_HANDLER)
    .timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).nearMinute(0).inTimezone(REPORTE_TZ)
    .create();
  return obtenerEstadoReporteSemanal();
}

function desactivarReporteSemanal() {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede desactivar el reporte.' };
  reporteTriggers_().forEach(t => ScriptApp.deleteTrigger(t));
  return obtenerEstadoReporteSemanal();
}

function enviarReporteAhora() {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede enviar el reporte.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return enviarReporteSemanalCore_('Manual');
  } finally {
    lock.releaseLock();
  }
}


/* ====================================================================
   FEEDBACK / MEJORAS / BUGS
   Cualquier perfil puede enviar un mensaje; solo lo reciben los perfiles Maestro (correo) y queda
   registrado en la hoja "Feedback". El correo lleva Reply-To del remitente para poder responderle directo.
   ==================================================================== */
const HOJA_FEEDBACK = 'Feedback';
const FEEDBACK_TIPOS = ['Feedback', 'Sugerencia de mejora', 'Reportar un bug'];
const FEEDBACK_MAX_POR_HORA = 5;

function escHtmlSrv_(t) {
  return String(t == null ? '' : t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/\n/g, '<br>');
}
/** Evita que un texto que empiece con = + - @ se interprete como fórmula al guardarlo en la hoja. */
function textoSeguroHoja_(t) {
  t = String(t == null ? '' : t);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function enviarFeedback(datos) {
  datos = datos || {};
  const tipo = FEEDBACK_TIPOS.indexOf(datos.tipo) !== -1 ? datos.tipo : 'Feedback';
  const mensaje = String(datos.mensaje || '').trim().slice(0, 2000);
  const detalle = String(datos.detalle || '').trim().slice(0, 2000);
  const pestana = String(datos.pestana || '').trim().slice(0, 80);
  const navegador = String(datos.navegador || '').trim().slice(0, 250);
  const versionPortal = String(datos.versionPortal || '').trim().slice(0, 30);
  if (mensaje.length < 5) return { ok: false, error: 'Escribe tu mensaje (mínimo 5 caracteres).' };

  const u = obtenerUsuarioActual();
  const quien = u.correo || u.nombre || 'anónimo';

  // Límite por persona para evitar abuso: máximo 5 mensajes por hora
  try {
    const cache = CacheService.getScriptCache();
    const clave = 'fb_' + hashCorto_(String(quien).toLowerCase());
    const n = Number(cache.get(clave)) || 0;
    if (n >= FEEDBACK_MAX_POR_HORA) return { ok: false, error: 'Ya enviaste varios mensajes seguidos. Inténtalo de nuevo en un rato, gracias.' };
    cache.put(clave, String(n + 1), 3600);
  } catch (e) {}

  const ahora = new Date();
  const fechaTxt = Utilities.formatDate(ahora, REPORTE_TZ, 'dd/MM/yyyy HH:mm');

  // 1) Registro en la hoja (queda aunque falle el correo)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA_FEEDBACK);
  if (!h) {
    h = ss.insertSheet(HOJA_FEEDBACK);
    h.getRange(1, 1, 1, 11).setValues([['Fecha', 'Nombre', 'Correo', 'Perfil', 'Área', 'Tipo', 'Pestaña', 'Mensaje', 'Detalle (bug)', 'Versión portal', 'Navegador']]);
    h.setFrozenRows(1);
  }
  h.appendRow([fechaTxt, textoSeguroHoja_(u.nombre), textoSeguroHoja_(u.correo), textoSeguroHoja_(u.subRol || u.rol), textoSeguroHoja_(u.area),
               tipo, textoSeguroHoja_(pestana), textoSeguroHoja_(mensaje), textoSeguroHoja_(detalle), textoSeguroHoja_(versionPortal), textoSeguroHoja_(navegador)]);

  // 2) Correo SOLO a los perfiles Maestro
  let entregado = false;
  try {
    const destinos = obtenerCorreosMaestros_();
    if (destinos.length) {
      const icono = tipo === 'Reportar un bug' ? '🐞' : (tipo === 'Sugerencia de mejora' ? '💡' : '💬');
      const color = tipo === 'Reportar un bug' ? '#B8403A' : (tipo === 'Sugerencia de mejora' ? '#B98A3D' : '#1B5FAE');
      const fila = (k, v) => v ? `<tr><td style="padding:4px 12px 4px 0;color:#6B7A99;font-size:12px;white-space:nowrap;vertical-align:top;">${k}</td><td style="padding:4px 0;font-size:13px;">${escHtmlSrv_(v)}</td></tr>` : '';
      const cuerpo = `
        <div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;">
          <div style="background:#0B1D33;color:#fff;padding:18px 22px;border-radius:8px 8px 0 0;">
            <div style="font-size:12px;opacity:.75;">Portal CAPEX P&S · Mensaje de un usuario</div>
            <h2 style="margin:6px 0 0;font-size:17px;">${icono} ${escHtmlSrv_(tipo)}</h2>
          </div>
          <div style="background:#fff;padding:20px 22px;border:1px solid #e0e0e0;border-top:4px solid ${color};">
            <div style="background:#F7F9FB;border-radius:8px;padding:14px 16px;font-size:14px;line-height:1.5;margin-bottom:14px;">${escHtmlSrv_(mensaje)}</div>
            ${detalle ? `<div style="font-size:12px;color:#6B7A99;font-weight:700;margin:4px 0;">Detalles del bug</div><div style="background:#FBF3F2;border-radius:8px;padding:12px 14px;font-size:13px;line-height:1.5;margin-bottom:14px;">${escHtmlSrv_(detalle)}</div>` : ''}
            <table style="border-collapse:collapse;">
              ${fila('De', u.nombre + (u.correo ? ' <' + u.correo + '>' : ''))}
              ${fila('Perfil / área', [u.subRol || u.rol, u.area].filter(Boolean).join(' · '))}
              ${fila('Pestaña', pestana)}
              ${fila('Fecha', fechaTxt)}
              ${fila('Versión portal', versionPortal)}
              ${fila('Navegador', navegador)}
            </table>
          </div>
          <div style="background:#f0f4f8;padding:10px 22px;border-radius:0 0 8px 8px;font-size:11px;color:#6B7A99;">
            Puedes responder este correo: le llegará directo a quien lo envió. También queda registrado en la hoja "${HOJA_FEEDBACK}".
          </div>
        </div>`;
      const opciones = {
        to: destinos.join(','),
        subject: `[Portal CAPEX · ${tipo}] ${u.nombre || 'Usuario'}: ${mensaje.replace(/\s+/g, ' ').slice(0, 70)}`,
        htmlBody: cuerpo,
      };
      if (u.correo) opciones.replyTo = u.correo;
      MailApp.sendEmail(opciones);
      entregado = true;
    }
  } catch (e) { Logger.log('Feedback: no se pudo enviar el correo: ' + e); }

  return { ok: true, entregado: entregado };
}


/* ====================================================================
   ÚLTIMAS SOLICITUDES DE CAMBIO DE FECHA (para Vista General)
   Muestra las últimas N (cualquier estado: Pendiente, Aprobado, Rechazado), con el proyecto.
   ==================================================================== */
function obtenerUltimosCambiosFecha(n) {
  n = Math.min(Math.max(Number(n) || 6, 1), 20);
  return conCache_('ultcambfecha_' + n, {}, () => {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HistorialFechas');
    if (!hoja || hoja.getLastRow() < 2) return [];
    const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 6).getValues()
      .map((f, i) => ({
        fila: i + 2, oc: f[0],
        fechaAnterior: formatearFecha(f[1]) || f[1],
        fechaNueva: formatearFecha(f[2]) || f[2],
        usuario: f[3], fechaCambioRaw: f[4], fechaCambio: formatearFecha(f[4]),
        estado: f[5] || 'Pendiente',
      }))
      .filter(x => x.oc !== '');
    filas.sort((a, b) => new Date(b.fechaCambioRaw) - new Date(a.fechaCambioRaw));
    const ultimos = filas.slice(0, n);
    const set = {};
    ultimos.forEach(x => { set[String(x.oc)] = true; });
    const info = ultimos.length ? infoLigeraOCs_(set) : {};
    return ultimos.map(x => ({
      oc: x.oc, proyecto: (info[String(x.oc)] || {}).proyecto || '',
      fechaAnterior: x.fechaAnterior, fechaNueva: x.fechaNueva,
      usuario: x.usuario, fechaCambio: x.fechaCambio, estado: x.estado,
    }));
  });
}


/* ====================================================================
   IMPORTAR FECHAS DESDE "OCHans_Historial"
   Un gestor (Hans) mantiene su propia pestaña con OC, monto y una posible fecha de facturación,
   porque no tiene tiempo de cargar cada OC a mano en el portal. Este botón cruza esa hoja con
   BASE_CASHFLOW y, por cada OC con una fecha nueva o distinta, crea una SOLICITUD DE CAMBIO DE
   FECHA pendiente de aprobar (la misma cola que ya se ve en Aprobaciones) — nunca escribe la
   fecha directo en BASE_CASHFLOW. Los datos de Hans tienen prioridad (se proponen tal cual,
   sin comparar quién tiene "la razón"), pero solo se APLICAN cuando un Maestro los aprueba.
   Estructura confirmada de "OCHans_Historial":
     A = Fecha/Hora, B = N° OC, C = Monto, D = Estado, E = Detalle (fecha posible de facturación),
     F = Tipo de Cambio, G = marca de tiempo de captura.
   Columna E vacía = esa fila no aporta fecha, se ignora. Si una OC aparece en varias filas, se
   toma la ÚLTIMA fila (la más reciente) de esa OC.
   ==================================================================== */
const HOJA_OCHANS = 'OCHans_Historial';
const COL_OCHANS_OC = 2, COL_OCHANS_FECHA = 5;   // B y E (1-based)

function mismaFecha_(a, b) {
  return (formatearFecha(a) || '') === (formatearFecha(b) || '');   // misma normalización que usa el resto del portal
}

/**
 * Calcula (SIN escribir nada) qué pasaría si se importa "OCHans_Historial" ahora mismo:
 *  - "aplicar": OC que hoy NO tienen fecha -> se rellenan directo, sin aprobación (no hay nada que pisar).
 *  - "aprobar": OC que YA tienen una fecha distinta -> quedan pendientes de aprobar (se está reemplazando algo).
 * Se usa dos veces con los mismos datos: una para la vista previa, otra para ejecutar de verdad.
 */
function calcularImportacionOCHans_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHans = ss.getSheetByName(HOJA_OCHANS);
  if (!hojaHans || hojaHans.getLastRow() < 2) return { error: 'No encuentro la pestaña "OCHans_Historial", o está vacía.' };

  const filasHans = hojaHans.getRange(2, 1, hojaHans.getLastRow() - 1, Math.max(COL_OCHANS_OC, COL_OCHANS_FECHA)).getValues();
  const porOC = {};   // última fila con fecha, por OC
  filasHans.forEach(f => {
    const oc = String(f[COL_OCHANS_OC - 1] || '').trim();
    const fecha = f[COL_OCHANS_FECHA - 1];
    if (!oc || !fecha) return;   // sin OC o sin fecha (columna E vacía): se ignora esa fila
    porOC[oc] = fecha;
  });

  const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
  if (!hojaBase || hojaBase.getLastRow() < 2) return { error: 'BASE_CASHFLOW está vacía.' };
  const datosBase = hojaBase.getRange(2, 1, hojaBase.getLastRow() - 1, Math.min(ENCABEZADOS_BASE.length, hojaBase.getLastColumn())).getValues();
  // Solo interesan las OC que TODAVÍA tienen saldo pendiente (si ya está pagada del todo, no hace
  // falta proponerle fecha). Esto es lo que evitaba que el conteo saliera inflado con OC ya cerradas.
  const todasLasOC = {}, porOCBase = {};
  agruparPorOC_(datosBase).forEach(g => {
    todasLasOC[String(g.oc)] = g;
    if ((g.pendienteOC || 0) > 0) porOCBase[String(g.oc)] = g;
  });

  const pendientes = {};   // OC -> última solicitud pendiente (para no duplicar)
  leerCambiosFechaPendientesRaw_().forEach(c => { pendientes[String(c.oc)] = c; });

  const aplicar = [], aprobar = [];
  let sinCambio = 0, noEncontradas = 0, yaPendientes = 0, sinSaldo = 0;
  Object.keys(porOC).forEach(oc => {
    const fechaNueva = porOC[oc];
    if (!(oc in todasLasOC)) { noEncontradas++; return; }
    const g = porOCBase[oc];
    if (!g) { sinSaldo++; return; }   // existe, pero ya no tiene saldo pendiente: no necesita fecha
    const fechaActual = g.fechaPago;
    if (mismaFecha_(fechaActual, fechaNueva)) { sinCambio++; return; }
    const yaPend = pendientes[oc];
    if (yaPend && mismaFecha_(yaPend.fechaNueva, fechaNueva)) { yaPendientes++; return; }
    const item = { oc: oc, fechaNueva: fechaNueva, proyecto: g.proyecto || '', area: g.area || '' };
    if (!fechaActual) aplicar.push(item);                                                   // sin fecha previa: se rellena directo
    else aprobar.push(Object.assign({}, item, { fechaActual: fechaActual }));               // ya tenía otra: pide aprobación
  });
  return { aplicar: aplicar, aprobar: aprobar, sinCambio: sinCambio, noEncontradas: noEncontradas, yaPendientes: yaPendientes, sinSaldo: sinSaldo, total: Object.keys(porOC).length };
}

/** Paso 1 (no escribe nada): cuántos cambios habría, para mostrar el aviso "¿Proceder?" antes de tocar algo. */
function previsualizarImportacionOCHans() {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede correr esta importación.' };
  const r = calcularImportacionOCHans_();
  if (r.error) return { ok: false, error: r.error };
  // Apps Script devuelve null al navegador si la respuesta trae objetos Date: se pasan a texto yyyy-MM-dd
  const limpiar = c => Object.assign({}, c, {
    fechaNueva: formatearFecha(c.fechaNueva),
    fechaActual: formatearFecha(c.fechaActual),
  });
  return {
    ok: true, total: r.total, aplicarDirecto: r.aplicar.length, pendientesAprobar: r.aprobar.length,
    sinCambio: r.sinCambio, noEncontradas: r.noEncontradas, yaPendientes: r.yaPendientes, sinSaldo: r.sinSaldo,
    aplicar: r.aplicar.map(limpiar), aprobar: r.aprobar.map(limpiar),   // detalle con proyecto, para mostrar la lista antes de confirmar
  };
}

/**
 * Paso 2 (ya confirmado por el Maestro): aplica directo las OC sin fecha previa, deja las demás
 * como solicitud pendiente de aprobar, y manda el correo de resumen a los perfiles Maestro.
 */
function ejecutarImportacionOCHans() {
  if (!puedeAprobarEliminacion_()) return { ok: false, error: 'Solo un perfil Maestro puede correr esta importación.' };
  const r = calcularImportacionOCHans_();
  if (r.error) return { ok: false, error: r.error };
  if (!r.aplicar.length && !r.aprobar.length) {
    return { ok: true, aplicadas: 0, pendientes: 0, sinCambio: r.sinCambio, noEncontradas: r.noEncontradas, yaPendientes: r.yaPendientes, sinSaldo: r.sinSaldo, total: r.total };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = ss.getSheetByName(CONFIG.HOJA_BASE);
  const colOC = hojaBase.getRange(2, COL_BASE.OC, hojaBase.getLastRow() - 1, 1).getValues();
  const indicesPorOC = {};   // TODAS las filas de cada OC (una OC puede repetirse en varias líneas)
  colOC.forEach((f, i) => { const oc = String(f[0]); (indicesPorOC[oc] = indicesPorOC[oc] || []).push(i + 2); });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const hojaHist = obtenerHojaHistorialFechas_();
    const usuario = 'OCHans_Historial (import)';
    const ahora = new Date();

    // Directo: se escribe en BASE_CASHFLOW y queda registrado como "Aprobado" (trazable, sin quedar pendiente)
    r.aplicar.forEach(c => {
      (indicesPorOC[c.oc] || []).forEach(fila => hojaBase.getRange(fila, COL_BASE.FECHA_PAGO).setValue(c.fechaNueva));
      hojaHist.appendRow([c.oc, '', c.fechaNueva, usuario, ahora, 'Aprobado']);
    });
    // Con fecha previa: solicitud pendiente, como cualquier otro cambio de fecha
    r.aprobar.forEach(c => {
      hojaHist.appendRow([c.oc, c.fechaActual, c.fechaNueva, usuario, ahora, 'Pendiente']);
    });
  } finally {
    lock.releaseLock();
  }
  bumpCacheVer_();

  try { enviarResumenImportacionOCHans_(r); } catch (eCorreo) { Logger.log('No se pudo enviar el correo de resumen de OCHans: ' + eCorreo); }

  return { ok: true, aplicadas: r.aplicar.length, pendientes: r.aprobar.length, sinCambio: r.sinCambio, noEncontradas: r.noEncontradas, yaPendientes: r.yaPendientes, sinSaldo: r.sinSaldo, total: r.total };
}

/** Correo a los perfiles Maestro con el detalle de lo agregado (directo) y lo que quedó pendiente de aprobar. */
function enviarResumenImportacionOCHans_(r) {
  const destinos = obtenerCorreosMaestros_();
  if (!destinos.length) return;
  const tz = Session.getScriptTimeZone();
  const f = v => formatearFecha(v) ? Utilities.formatDate(new Date(formatearFecha(v)), tz, 'dd/MM/yyyy') : '—';
  const filaAplicar = c => `<tr><td style="padding:5px 10px;">${c.oc}</td><td style="padding:5px 10px;">${escHtmlSrv_(c.proyecto)}</td><td style="padding:5px 10px;">${f(c.fechaNueva)}</td></tr>`;
  const filaAprobar = c => `<tr><td style="padding:5px 10px;">${c.oc}</td><td style="padding:5px 10px;">${escHtmlSrv_(c.proyecto)}</td><td style="padding:5px 10px;">${f(c.fechaActual)}</td><td style="padding:5px 10px;">${f(c.fechaNueva)}</td></tr>`;
  const cuerpo = `
    <div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;">
      <div style="background:#0B1D33;color:#fff;padding:18px 22px;border-radius:8px 8px 0 0;">
        <div style="font-size:12px;opacity:.75;">Portal CAPEX P&S · Importación de fechas</div>
        <h2 style="margin:6px 0 0;font-size:17px;">📥 Fechas importadas desde OCHans_Historial</h2>
      </div>
      <div style="background:#fff;padding:20px 22px;border:1px solid #e0e0e0;border-top:4px solid #2E6FBE;">
        ${r.aplicar.length ? `<h3 style="color:#4C9270;font-size:13px;">✅ Fechas agregadas automáticamente (${r.aplicar.length}) — no tenían fecha antes</h3>
          <table style="border-collapse:collapse;font-size:13px;width:100%;margin-bottom:16px;"><tr style="color:#6B7A99;font-size:11px;text-align:left;"><td style="padding:5px 10px;">N° OC</td><td style="padding:5px 10px;">Proyecto</td><td style="padding:5px 10px;">Fecha nueva</td></tr>${r.aplicar.map(filaAplicar).join('')}</table>` : ''}
        ${r.aprobar.length ? `<h3 style="color:#B98A3D;font-size:13px;">⏳ Cambios pendientes de aprobar (${r.aprobar.length}) — ya tenían otra fecha</h3>
          <table style="border-collapse:collapse;font-size:13px;width:100%;"><tr style="color:#6B7A99;font-size:11px;text-align:left;"><td style="padding:5px 10px;">N° OC</td><td style="padding:5px 10px;">Proyecto</td><td style="padding:5px 10px;">Fecha anterior</td><td style="padding:5px 10px;">Fecha nueva</td></tr>${r.aprobar.map(filaAprobar).join('')}</table>
          <p style="font-size:12px;color:#6B7A99;margin-top:10px;">Revísalas en el portal, pestaña Aprobaciones → Cambios de fecha de pago.</p>` : ''}
        ${!r.aplicar.length && !r.aprobar.length ? '<p style="font-size:13px;color:#4C9270;">No hubo fechas nuevas que aplicar esta vez.</p>' : ''}
      </div>
      <div style="background:#f0f4f8;padding:10px 22px;border-radius:0 0 8px 8px;font-size:11px;color:#6B7A99;">
        ${r.sinCambio} OC sin cambio · ${r.sinSaldo || 0} ya no tienen saldo pendiente · ${r.noEncontradas} OC de Hans no encontradas en BASE_CASHFLOW.
      </div>
    </div>`;
  MailApp.sendEmail({ to: destinos.join(','), subject: `[Portal CAPEX] Importación de fechas — ${r.aplicar.length} agregadas, ${r.aprobar.length} por aprobar`, htmlBody: cuerpo });
}

function obtenerHojaHistorialFechas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName('HistorialFechas');
  if (!h) {
    h = ss.insertSheet('HistorialFechas');
    h.getRange(1, 1, 1, 6).setValues([['N° OC', 'Fecha anterior', 'Fecha nueva', 'Usuario', 'Fecha del cambio', 'Estado']]);
    h.setFrozenRows(1);
  }
  return h;
}
