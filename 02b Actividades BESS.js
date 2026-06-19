/******************** FUNCIÓN PRINCIPAL ACTIVIDADES BESS ********************/
/**
 * dr_actividades_bess con lógica de mapeo dinámico.
 * Busca las columnas por nombre para mantener el mismo orden (15 columnas de datos + 3 de metadatos)
 * en la hoja destino, permitiendo compatibilidad plena sin importar qué columnas nuevas aparezcan.
 */
function dr_actividades_bess(carpetaID, id_archivo, nom_hoja, fecha) {
  const carpeta = DriveApp.getFolderById(carpetaID);
  const archivos = carpeta.getFiles();
  const hojaCarga = SpreadsheetApp.openById(id_archivo).getSheetByName(nom_hoja);

  hojaCarga.getFilter()?.remove();

  const hoy = new Date();
  const ayer = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - 1);
  const sinHoras = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const parseYYYYMMDD = (s) => {
    const [Y, M, D] = s.split('-').map(Number);
    return new Date(Y, (M - 1), D);
  };

  let desde = sinHoras(ayer), hasta = sinHoras(ayer);
  if (fecha) {
    if (typeof fecha === 'string') {
      const f = sinHoras(parseYYYYMMDD(fecha)); desde = f; hasta = f;
    } else if (fecha instanceof Date) {
      const f = sinHoras(fecha); desde = f; hasta = f;
    } else if (typeof fecha === 'object' && fecha.desde && fecha.hasta) {
      desde = sinHoras(parseYYYYMMDD(fecha.desde)); hasta = sinHoras(parseYYYYMMDD(fecha.hasta));
    }
  }

  const fechasAEliminar = [];
  const datosParaInsertar = [];
  const archivosProcesados = [];

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName();
    if (nombreArchivo.indexOf('DR') === -1) continue;
    if (archivo.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    const fm = archivo.getLastUpdated();
    const fmSin = new Date(fm.getFullYear(), fm.getMonth(), fm.getDate());
    if (fmSin < desde || fmSin > hasta) continue;

    Logger.log(`Procesando archivo: ${nombreArchivo}`);

    const regexNombreArchivo = /^(\w{5})_DR_(\d{8})/;
    const match = nombreArchivo.match(regexNombreArchivo);
    const proyecto = match ? match[1] : 'Proyecto desconocido';
    const fecha_nombre_archivo = match ? match[2] : 'Fecha desconocida';

    const archivoSS = SpreadsheetApp.openById(archivo.getId());
    const hoja = archivoSS.getSheets()[0];
    const valores = hoja.getDataRange().getValues();

    // Fecha desde M6 o desde celda identificadora
    let celdaFecha = null;
    for (let i = 0; i < valores.length; i++) {
      for (let j = 0; j < valores[i].length; j++) {
        if (String(valores[i][j]).trim().toLowerCase() === "fecha de daily report:") {
          celdaFecha = valores[i][j + 1];
          break;
        }
      }
      if (celdaFecha) break;
    }
    if (!celdaFecha) celdaFecha = hoja.getRange("M6").getValue();

    let fechaFormateada;
    if (celdaFecha && !isNaN(new Date(celdaFecha).getTime())) {
      fechaFormateada = formatearFechaSimple(new Date(celdaFecha));
    } else {
      Logger.log(`Fecha inválida/ausente en ${nombreArchivo} -> se omite.`);
      continue;
    }

    if (fecha_nombre_archivo && /^\d{8}$/.test(fecha_nombre_archivo)) {
      const fechaNombreDate = yyyymmddToDate(fecha_nombre_archivo);
      const fechaNombreForm = formatearFechaSimple(fechaNombreDate);
      if (fechaNombreForm !== fechaFormateada) {
        Logger.log(`Archivo con fecha NO consistente: ${nombreArchivo}`);
        continue;
      }
    }

    fechasAEliminar.push(fechaFormateada);

    // Buscar "Actividad" y "Observaciones"/"Firma"
    let filaInicio = -1, filaFin = -1;
    for (let i = 0; i < valores.length; i++) {
        let filaStr = valores[i].join("").toLowerCase();
        
        // Identificar la fila de encabezados evaluando 2 filas juntas (para sortear celdas combinadas)
        let dosFilasStr = filaStr + " " + (valores[i + 1] ? valores[i + 1].join("").toLowerCase() : "");
        if (filaInicio === -1 && dosFilasStr.includes("actividad") && dosFilasStr.includes("unidad") && dosFilasStr.includes("cantidad") && dosFilasStr.includes("acumulad")) {
            filaInicio = i;
        } 
        else if (filaInicio !== -1 && filaFin === -1 && filaStr.includes("firma")) {
            filaFin = i - 1;
        }
    }

    if (filaInicio === -1) {
      Logger.log(`No se encontró la cabecera de Actividades en ${nombreArchivo} -> se omite.`);
      continue;
    }
    if (filaFin === -1) {
      filaFin = valores.length - 1; // Fallback si no encuentra 'firma'
    }

    // Mapeo dinámico de columnas desde la fila que tiene los títulos
    let headers = valores[filaInicio];
    let colMap = {
      id: -1, actividad: -1, unidad: -1, cantidad: -1, pu: -1, valor_presupuesto: -1,
      cantidad_anterior: -1, cantidad_hoy: -1, cantidad_acumulada: -1,
      porcentaje_anterior: -1, porcentaje_hoy: -1, porcentaje_acumulado: -1,
      porcentaje_acumulado_liberado: -1, cantidad_hoy_liberado: -1, cantidad_acumulada_liberado: -1, empresa: -1
    };

    // Subtitulos pueden estar en la fila Inicio o en la de inmediatamente abajo (por Celdas Combinadas Verticalmente)
    for (let j = 0; j < valores[filaInicio].length; j++) {
      let h1 = String(valores[filaInicio][j] || "").trim().toLowerCase().replace(/\s+/g, ' ');
      let h2 = String(valores[filaInicio + 1]?.[j] || "").trim().toLowerCase().replace(/\s+/g, ' ');
      
      let match = (key) => h1 === key || h2 === key;
      let includes = (key) => h1.includes(key) || h2.includes(key);

      if (colMap.id === -1 && (h1 === "id" || h2 === "id")) colMap.id = j;
      else if (colMap.actividad === -1 && includes("actividad")) colMap.actividad = j;
      else if (colMap.unidad === -1 && includes("unidad")) colMap.unidad = j;
      else if (colMap.cantidad_anterior === -1 && includes("cantidad") && includes("anterior")) colMap.cantidad_anterior = j;
      else if (colMap.cantidad_hoy === -1 && includes("cantidad") && includes("hoy")) colMap.cantidad_hoy = j;
      else if (colMap.cantidad_acumulada === -1 && includes("cantidad") && includes("acumulad") && !includes("%") && !includes("liberado")) colMap.cantidad_acumulada = j;
      else if (colMap.cantidad === -1 && includes("cantidad") && !includes("anterior") && !includes("hoy") && !includes("acumulad") && !includes("liberado")) colMap.cantidad = j;
      else if (colMap.pu === -1 && (includes("pu") || includes("p.u"))) colMap.pu = j;
      else if (colMap.valor_presupuesto === -1 && includes("valor") && (includes("presupuesto") || includes("pp"))) colMap.valor_presupuesto = j;
      else if (colMap.porcentaje_anterior === -1 && includes("%") && includes("anterior")) colMap.porcentaje_anterior = j;
      else if (colMap.porcentaje_hoy === -1 && includes("%") && includes("hoy")) colMap.porcentaje_hoy = j;
      else if (colMap.porcentaje_acumulado === -1 && includes("%") && includes("acumulad") && !includes("liberado")) colMap.porcentaje_acumulado = j;
      else if (colMap.porcentaje_acumulado_liberado === -1 && includes("%") && includes("acumulad") && includes("liberado")) colMap.porcentaje_acumulado_liberado = j;
      else if (colMap.cantidad_hoy_liberado === -1 && includes("cantidad") && includes("hoy") && includes("liberado")) colMap.cantidad_hoy_liberado = j;
      else if (colMap.cantidad_acumulada_liberado === -1 && includes("cantidad") && includes("acumulad") && includes("liberado")) colMap.cantidad_acumulada_liberado = j;
      else if (colMap.empresa === -1 && includes("empresa")) colMap.empresa = j;
    }

    Logger.log(`[DEBUG] Archivo: ${nombreArchivo} | filaInicio=${filaInicio}, filaFin=${filaFin}`);
    Logger.log(`[DEBUG] FilaInicio: ${valores[filaInicio].join(" | ")}`);
    Logger.log(`[DEBUG] FilaInicio+1: ${valores[filaInicio+1] ? valores[filaInicio+1].join(" | ") : "N/A"}`);
    Logger.log(`[DEBUG] Mapa de columnas resuelto: ${JSON.stringify(colMap)}`);

    let contSkippedNaN = 0;
    // Procesar filas de datos usando el mapa de columnas (desde la fila 2 post-headers por si hay merges)
    for (let i = filaInicio + 2; i <= filaFin; i++) {
      let filaOriginal = valores[i];

      let acum = colMap.cantidad_acumulada !== -1 ? parseFloat(filaOriginal[colMap.cantidad_acumulada]) : NaN;
      if (isNaN(acum)) {
        contSkippedNaN++;
        continue; // Mismo filtro: se ignora si acumulado no es válido o está vacío
      }

      let empresa_val = colMap.empresa !== -1 ? String(filaOriginal[colMap.empresa]).trim() : "";


      // Estructura fija de 15 columnas
      let nuevaFila = new Array(15).fill("");

      if (colMap.id !== -1) nuevaFila[0] = filaOriginal[colMap.id] ? String(filaOriginal[colMap.id]).trim() : "";
      if (colMap.actividad !== -1) nuevaFila[1] = filaOriginal[colMap.actividad];
      if (colMap.unidad !== -1) nuevaFila[2] = filaOriginal[colMap.unidad];
      if (colMap.cantidad !== -1) nuevaFila[3] = filaOriginal[colMap.cantidad];
      if (colMap.pu !== -1) nuevaFila[4] = filaOriginal[colMap.pu];
      if (colMap.valor_presupuesto !== -1) nuevaFila[5] = filaOriginal[colMap.valor_presupuesto];
      if (colMap.cantidad_anterior !== -1) nuevaFila[6] = filaOriginal[colMap.cantidad_anterior];
      if (colMap.cantidad_hoy !== -1) nuevaFila[7] = filaOriginal[colMap.cantidad_hoy];
      if (colMap.cantidad_acumulada !== -1) nuevaFila[8] = filaOriginal[colMap.cantidad_acumulada];
      if (colMap.porcentaje_anterior !== -1) nuevaFila[9] = filaOriginal[colMap.porcentaje_anterior];
      if (colMap.porcentaje_hoy !== -1) nuevaFila[10] = filaOriginal[colMap.porcentaje_hoy];
      if (colMap.porcentaje_acumulado !== -1) nuevaFila[11] = filaOriginal[colMap.porcentaje_acumulado];
      if (colMap.porcentaje_acumulado_liberado !== -1) nuevaFila[12] = filaOriginal[colMap.porcentaje_acumulado_liberado];
      if (colMap.cantidad_hoy_liberado !== -1) nuevaFila[13] = filaOriginal[colMap.cantidad_hoy_liberado];
      if (colMap.cantidad_acumulada_liberado !== -1) nuevaFila[14] = filaOriginal[colMap.cantidad_acumulada_liberado];

      // Añadir metadata y empresa al final
      nuevaFila.push(fechaFormateada, proyecto, fecha_nombre_archivo, empresa_val);
      datosParaInsertar.push(nuevaFila);
    }

    Logger.log(`[DEBUG] Filas extraídas con éxito: ${datosParaInsertar.length}. Filas descartadas por no tener Acumulado (vacío o texto): ${contSkippedNaN}`);

    if (datosParaInsertar.length > 0) {
      archivosProcesados.push(nombreArchivo);
    }
  }

  // --- Limpieza ---
  if (fechasAEliminar.length > 0) {
    const fechasUnicas = Array.from(new Set(fechasAEliminar));
    const fechasSet = new Set(fechasUnicas);

    const dateCol = 16;  // Índice 15 base 1 = Columna 16
    const startRow = 2;

    const rowsToDelete = _findRowIndicesByDates(hojaCarga, dateCol, fechasSet, startRow);

    if (rowsToDelete.length > 0) {
      const ranges = _compactRowIndicesToRanges(rowsToDelete);
      _deleteRowRanges(hojaCarga, ranges);
      Logger.log(`Limpieza: borradas ${rowsToDelete.length} filas en ${ranges.length} bloque(s).`);
      Utilities.sleep(2000);
    }
  }

  // --- Inserción ---
  if (datosParaInsertar.length > 0) {
    const startRow = hojaCarga.getLastRow() + 1;
    hojaCarga.getRange(startRow, 1, datosParaInsertar.length, datosParaInsertar[0].length).setValues(datosParaInsertar);
    Logger.log(`✅ Insertadas ${datosParaInsertar.length} filas nuevas.`);
  }

  // Verificación de inconsistencias
  try {
    const discrepancias = [];
    datosParaInsertar.forEach((fila, idx) => {
      const anterior = parseFloat(fila[6]);
      const hoy = parseFloat(fila[7]);
      const acumulado = parseFloat(fila[8]);

      if (!isNaN(anterior) && !isNaN(hoy) && !isNaN(acumulado)) {
        const suma = +(anterior + hoy).toFixed(4);
        const esperado = +acumulado.toFixed(4);
        if (Math.abs(suma - esperado) > 0.0001) {
          discrepancias.push(`Fila nueva ${idx + 1}: ${fila[1]} (Anterior: ${anterior}, Hoy: ${hoy}, Acum: ${acumulado}, Fecha: ${fila[15]})`);
        }
      }
    });

    if (discrepancias.length > 0) {
      const cuerpo = [
        `Se detectaron discrepancias en el cálculo de acumulados en la hoja "${nom_hoja}".`,
        ``, ...discrepancias, ``
      ].join('\n');
      Logger.log(`⚠️ Discrepancias encontradas:\n${cuerpo}`);
    }
  } catch (e) {
    Logger.log(`Error verificando acumulados: ${e.message}`);
  }

  Logger.log(`Archivos procesados: ${archivosProcesados.join(', ')}`);
}
