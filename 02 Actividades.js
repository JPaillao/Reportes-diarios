/******************** FUNCIÓN PRINCIPAL ********************/

/**
 * dr_actividades con fecha de modificación parametrizable y ejecución optimizada.
 *
 * - carpetaID  : ID de la carpeta en Drive donde están los DR.
 * - id_archivo : ID del Spreadsheet destino (registro consolidado).
 * - nom_hoja   : Nombre de la hoja dentro del Spreadsheet destino.
 * - fecha      : Opcional: Date | 'YYYY-MM-DD' | {desde:'YYYY-MM-DD', hasta:'YYYY-MM-DD'}
 *                Si no se pasa, usa "ayer".
 *
 * Cambios clave:
 *  - Borrado por BLOQUES solo de las fechas afectadas (columna 16), no clear total.
 *  - Inserción en UN solo setValues con todas las filas nuevas (reduce recálculo).
 *  - Pausa condicional tras borrar (solo si hubo filas, 2 s) para permitir que cesen recálculos del consolidado.
 *  - Validación de consistencia: fecha en nombre (YYYYMMDD) debe calzar con M6 (dd-MM-yyyy).
 */
function dr_actividades(carpetaID, id_archivo, nom_hoja, fecha) {
  const carpeta = DriveApp.getFolderById(carpetaID);
  const archivos = carpeta.getFiles();
  const hojaCarga = SpreadsheetApp.openById(id_archivo).getSheetByName(nom_hoja);

  // Quitar filtro si existe
  hojaCarga.getFilter()?.remove();

  // --- Ventana de fechas (default = AYER) ---
  const hoy = new Date();
  const ayer = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - 1);

  const sinHoras = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const parseYYYYMMDD = (s) => {
    const [Y, M, D] = s.split('-').map(Number);
    return new Date(Y, (M - 1), D);
  };

  let desde = sinHoras(ayer);
  let hasta = sinHoras(ayer);

  if (fecha) {
    if (typeof fecha === 'string') {
      const f = sinHoras(parseYYYYMMDD(fecha));
      desde = f; hasta = f;
    } else if (fecha instanceof Date) {
      const f = sinHoras(fecha);
      desde = f; hasta = f;
    } else if (typeof fecha === 'object' && fecha.desde && fecha.hasta) {
      desde = sinHoras(parseYYYYMMDD(fecha.desde));
      hasta = sinHoras(parseYYYYMMDD(fecha.hasta));
    }
  }

  const fechasAEliminar = [];     // dd-MM-yyyy (desde M6)
  const datosParaInsertar = [];   // filas completas para setValues
  const archivosProcesados = [];  // logging

  // --- Recorrer archivos en la carpeta ---
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName();
    if (nombreArchivo.indexOf('DR') === -1) continue;
    if (archivo.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    const fm = archivo.getLastUpdated();
    const fmSin = new Date(fm.getFullYear(), fm.getMonth(), fm.getDate());
    if (fmSin < desde || fmSin > hasta) continue;

    Logger.log(`Procesando archivo: ${nombreArchivo}`);

    // Proyecto y fecha del nombre
    const regexNombreArchivo = /^(\w{5})_DR_(\d{8})/;
    const match = nombreArchivo.match(regexNombreArchivo);
    const proyecto = match ? match[1] : 'Proyecto desconocido';
    const fecha_nombre_archivo = match ? match[2] : 'Fecha desconocida';

    const archivoSS = SpreadsheetApp.openById(archivo.getId());
    const hoja = archivoSS.getSheets()[0];
    const valores = hoja.getDataRange().getValues();

    // Fecha operativa desde M6 (dd-MM-yyyy)
    const celdaFecha = hoja.getRange("M6").getValue();
    let fechaFormateada;
    if (celdaFecha && !isNaN(new Date(celdaFecha).getTime())) {
      fechaFormateada = formatearFechaSimple(new Date(celdaFecha));
      Logger.log(`M6 -> ${fechaFormateada}`);
    } else {
      Logger.log(`M6 inválida/ausente en ${nombreArchivo} -> se omite.`);
      continue;
    }

    // Validación: fecha del nombre vs M6
    if (fecha_nombre_archivo && /^\d{8}$/.test(fecha_nombre_archivo)) {
      const fechaNombreDate = yyyymmddToDate(fecha_nombre_archivo);
      const fechaNombreForm = formatearFechaSimple(fechaNombreDate);
      if (fechaNombreForm !== fechaFormateada) {
        const msg = [
          `Archivo con fecha NO consistente: ${nombreArchivo}`,
          `Proyecto: ${proyecto}`,
          `Fecha en nombre: ${fechaNombreForm} (de ${fecha_nombre_archivo})`,
          `Fecha en M6:     ${fechaFormateada}`,
          `Acción: NO se cargó este DR.`
        ].join('\n');
        Logger.log(msg);
        try {
          MailApp.sendEmail({
            to: "jpaillao@orion-power.com",
            subject: `🚨 DR con fecha inconsistente: ${proyecto} / ${nombreArchivo}`,
            body: `Hola,\n\n${msg}\n\nSaludos,\nDR Bot`
          });
        } catch (e) {
          Logger.log(`Error enviando correo por inconsistencia: ${e}`);
        }
        continue; // no cargar este archivo
      }
    }

    // Fecha para limpieza posterior en consolidado
    fechasAEliminar.push(fechaFormateada);

    // Buscar "Actividad" y "Observaciones"
    let filaInicio = -1, columnaInicio = -1, filaFin = -1;
    for (let i = 0; i < valores.length; i++) {
      for (let j = 0; j < valores[i].length; j++) {
        const cel = typeof valores[i][j] === 'string' ? valores[i][j].trim().toLowerCase() : '';
        if (cel === 'actividad' && filaInicio === -1) {
          filaInicio = i;
          columnaInicio = j - 1;
        } else if (cel === 'observaciones' && filaFin === -1) {
          filaFin = i - 3;
        }
      }
    }

    if (filaInicio === -1 || filaFin === -1 || filaInicio >= filaFin) {
      Logger.log(`No se encontró bloque Actividad/Observaciones en ${nombreArchivo} -> se omite.`);
      continue;
    }

    const datos = hoja.getRange(filaInicio + 2, 1, filaFin - filaInicio - 1, valores[0].length).getValues();
    const datosFiltrados = datos
      .filter(fila => {
        const acumulado = parseFloat(fila[8]);
        return !isNaN(acumulado); // incluye 0 válidos
      })
      .map(fila => {
        fila[0] = fila[0] ? fila[0].toString().trim() : "";
        fila.push(fechaFormateada, proyecto, fecha_nombre_archivo);
        return fila;
      });

    if (datosFiltrados.length > 0) {
      datosParaInsertar.push(...datosFiltrados); // ✅ reemplazo de concat
      archivosProcesados.push(nombreArchivo);
    }
  } // while archivos

  Logger.log(`Fechas a eliminar: ${JSON.stringify(Array.from(new Set(fechasAEliminar)))}`);

  // --- LIMPIEZA: borrar SOLO filas con fechas afectadas (columna 16) en BLOQUES ---
  if (fechasAEliminar.length > 0) {
    const fechasUnicas = Array.from(new Set(fechasAEliminar));
    const fechasSet = new Set(fechasUnicas);

    const dateCol = 16;  // actividades: la fecha destino está en la columna 16 (índice 15)
    const startRow = 2;  // encabezado en fila 1

    const rowsToDelete = _findRowIndicesByDates(hojaCarga, dateCol, fechasSet, startRow);

    if (rowsToDelete.length > 0) {
      const ranges = _compactRowIndicesToRanges(rowsToDelete);
      _deleteRowRanges(hojaCarga, ranges);
      Logger.log(`Limpieza actividades: borradas ${rowsToDelete.length} filas en ${ranges.length} bloque(s) para fechas ${JSON.stringify(fechasUnicas)}.`);
      Utilities.sleep(2000); // ✅ solo si hubo borrado
    } else {
      Logger.log(`Limpieza actividades: no había filas previas para fechas ${JSON.stringify(fechasUnicas)}.`);
    }
  }

  // --- INSERCIÓN: un solo setValues con TODO lo nuevo ---
  if (datosParaInsertar.length > 0) {
    const startRow = hojaCarga.getLastRow() + 1;
    const numRows = datosParaInsertar.length;
    const numCols = datosParaInsertar[0].length;

    hojaCarga.getRange(startRow, 1, numRows, numCols).setValues(datosParaInsertar);
    Logger.log(`✅ Insertadas ${numRows} filas nuevas en una sola operación.`);
  } else {
    Logger.log(`Sin datos nuevos para insertar.`);
  }

  ///////// Bloque de revisión de cantidades ///////////////////////////

  // --- VERIFICACIÓN DE CONSISTENCIA: CantidadAnterior + Hoy === Acumulado ---

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
          discrepancias.push(
            `Fila nueva ${idx + 1}: ${fila[1]} (Anterior: ${anterior}, Hoy: ${hoy}, Acum: ${acumulado}, Fecha: ${fila[15]})`
          );
        }
      }
    });

    if (discrepancias.length > 0) {
      const cuerpo = [
        `Se detectaron discrepancias en el cálculo de acumulados en la hoja "${nom_hoja}" del archivo consolidado.`,
        ``,
        ...discrepancias,
        ``,
        `Por favor revisar estas actividades.`
      ].join('\n');

      MailApp.sendEmail({
        to: "jpaillao@orion-power.com",
        subject: `⚠️ Discrepancias en acumulados de actividades - ${nom_hoja}`,
        body: cuerpo
      });

      Logger.log(`⚠️ Discrepancias encontradas:\n${cuerpo}`);
    }
  } catch (e) {
    Logger.log(`Error verificando acumulados: ${e.message}`);
  }



  Logger.log(`Archivos procesados: ${archivosProcesados.join(', ')}`);
}



// /** Anterior
//  * dr_actividades con fecha de modificación parametrizable.
//  * - carpetaID  : ID de la carpeta en Drive donde están los DR.
//  * - id_archivo : ID del Spreadsheet destino (registro consolidado).
//  * - nom_hoja   : Nombre de la hoja dentro del Spreadsheet destino.
//  * - fecha      : Opcional. Puede ser:
//  *                a) Date            (ej: new Date(2025,7,10))
//  *                b) 'YYYY-MM-DD'    (ej: '2025-08-10')
//  *                c) {desde:'YYYY-MM-DD', hasta:'YYYY-MM-DD'}  // Rango opcional
//  *              Si no se pasa, usa "ayer".
//  *
//  * Mantiene TODO igual salvo el filtro de fecha de modificación.
//  */
// function dr_actividades(carpetaID, id_archivo, nom_hoja, fecha) {
//   const carpeta = DriveApp.getFolderById(carpetaID);
//   const archivos = carpeta.getFiles();
//   const hojaCarga = SpreadsheetApp.openById(id_archivo).getSheetByName(nom_hoja);

//   // --- Resolver ventana de fechas (truncadas a día) ---
//   const hoy = new Date();
//   const ayer = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - 1);

//   /** Normaliza a fecha sin horas (local) */
//   const sinHoras = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());

//   /** Parse de 'YYYY-MM-DD' -> Date sin horas */
//   const parseYYYYMMDD = (s) => {
//     const [Y, M, D] = s.split('-').map(Number);
//     return new Date(Y, (M - 1), D);
//   };

//   let desde = sinHoras(ayer);
//   let hasta = sinHoras(ayer);

//   if (fecha) {
//     if (typeof fecha === 'string') {
//       const f = sinHoras(parseYYYYMMDD(fecha));
//       desde = f; hasta = f;
//     } else if (fecha instanceof Date) {
//       const f = sinHoras(fecha);
//       desde = f; hasta = f;
//     } else if (typeof fecha === 'object' && fecha.desde && fecha.hasta) {
//       desde = sinHoras(parseYYYYMMDD(fecha.desde));
//       hasta = sinHoras(parseYYYYMMDD(fecha.hasta));
//     }
//   }

//   const fechasAEliminar = [];
//   let datosParaInsertar = [];
//   let archivosProcesados = [];

//   // Quitar filtro si existe (igual que tu código)
//   hojaCarga.getFilter()?.remove();

//   while (archivos.hasNext()) {
//     const archivo = archivos.next();
//     const nombreArchivo = archivo.getName();
//     if (nombreArchivo.indexOf('DR') === -1) continue; // misma condición

//     const fm = archivo.getLastUpdated();
//     const fmSin = new Date(fm.getFullYear(), fm.getMonth(), fm.getDate());

//     // Mantener sólo archivos con lastUpdated dentro de [desde, hasta]
//     if (fmSin < desde || fmSin > hasta) continue;

//     Logger.log(`Procesando archivo: ${nombreArchivo}`);

//     // === Bloque original de parseo, intacto (solo envuelto aquí) ===
//     const regexNombreArchivo = /^(\w{5})_DR_(\d{8})/;
//     const match = nombreArchivo.match(regexNombreArchivo);
//     const proyecto = match ? match[1] : 'Proyecto desconocido';
//     const fecha_nombre_archivo = match ? match[2] : 'Fecha desconocida';

//     const archivoSS = SpreadsheetApp.openById(archivo.getId());
//     const hoja = archivoSS.getSheets()[0];
//     const valores = hoja.getDataRange().getValues();

//     const celdaFecha = hoja.getRange("M6").getValue();
//     let fechaFormateada;
//     if (celdaFecha && !isNaN(new Date(celdaFecha).getTime())) {
//       fechaFormateada = formatearFechaSimple(new Date(celdaFecha));
//       Logger.log(`La fecha en la celda M6 es ${new Date(celdaFecha)} Convertida a ${fechaFormateada}`);
//       fechasAEliminar.push(fechaFormateada);
//     } else {
//       Logger.log(`La fecha en la celda M6 no es válida o está ausente en el archivo ${nombreArchivo}`);
//       continue;
//     }

//     // Buscar "Actividad" y "Observaciones" (igual que tu lógica)
//     let filaInicio = -1, columnaInicio = -1, filaFin = -1;
//     for (let i = 0; i < valores.length; i++) {
//       for (let j = 0; j < valores[i].length; j++) {
//         const celda = typeof valores[i][j] === 'string' ? valores[i][j].trim().toLowerCase() : '';
//         if (celda === 'actividad' && filaInicio === -1) {
//           filaInicio = i;
//           columnaInicio = j - 1;
//         } else if (celda === 'observaciones' && filaFin === -1) {
//           filaFin = i - 3;
//         }
//       }
//     }

//     if (filaInicio === -1 || filaFin === -1 || filaInicio >= filaFin) {
//       Logger.log(`Error: No se encontraron las palabras clave "Actividad" u "Observaciones" en el archivo ${nombreArchivo}`);
//       continue;
//     }

//     const datos = hoja.getRange(filaInicio + 2, 1, filaFin - filaInicio - 1, valores[0].length).getValues();
//     const datosFiltrados = datos
//       .filter(fila => {
//         const acumulado = parseFloat(fila[8]);
//         return !isNaN(acumulado); // Incluye 0 válidos
//       })
//       .map(fila => {
//         fila[0] = fila[0] ? fila[0].toString().trim() : "";
//         fila.push(fechaFormateada, proyecto, fecha_nombre_archivo);
//         return fila;
//       });

//     if (datosFiltrados.length > 0) {
//       datosParaInsertar = datosParaInsertar.concat(datosFiltrados);
//       archivosProcesados.push(nombreArchivo);
//     }
//   }

//   Logger.log(`Fechas a eliminar: ${JSON.stringify(fechasAEliminar)}`);

//   // Borrar datos existentes para esas fechas (idéntico a tu flujo)
//   if (fechasAEliminar.length > 0) {
//     const datosHoja = hojaCarga.getDataRange().getValues();
//     const nuevasFilas = datosHoja.filter(fila => {
//       const fechaCelda = fila[15] ? formatearFechaSimple(new Date(fila[15])) : '';
//       return !fechasAEliminar.includes(fechaCelda);
//     });

//     hojaCarga.clear(); // Borra todo, incluyendo encabezados
//     if (nuevasFilas.length > 0) {
//       hojaCarga.getRange(1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
//     }

//     Logger.log(`Se eliminaron todas las filas con fechas: ${JSON.stringify(fechasAEliminar)}`);
//   }

//   // Insertar nuevos datos (igual que tu código)
//   if (datosParaInsertar.length > 0) {
//     const rangoInsertar = hojaCarga.getRange(hojaCarga.getLastRow() + 1, 1, datosParaInsertar.length, datosParaInsertar[0].length);
//     rangoInsertar.setValues(datosParaInsertar);
//     // Logger.log(`Datos a insertar: ${JSON.stringify(datosParaInsertar)}`);
//     Logger.log(`Datos insertados correctamente en Google Sheets.`);
//   }

//   Logger.log(`Archivos procesados: ${JSON.stringify(archivosProcesados)}`);
// }

// function dr_actividades(carpetaID, id_archivo, nom_hoja) {
//   const carpeta = DriveApp.getFolderById(carpetaID);
//   const archivos = carpeta.getFiles();
//   const hojaCarga = SpreadsheetApp.openById(id_archivo).getSheetByName(nom_hoja);
//   const hoy = new Date();
//   const hoySinHoras = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
//   const fechasAEliminar = [];
//   let datosParaInsertar = [];
//   let archivosProcesados = [];

//   // Elimina el filtro si existe
//   hojaCarga.getFilter()?.remove();

//   while (archivos.hasNext()) {
//     const archivo = archivos.next();
//     const nombreArchivo = archivo.getName();
//     const fechaModificacion = archivo.getLastUpdated();
//     const fechaModificacionSinHoras = new Date(fechaModificacion.getFullYear(), fechaModificacion.getMonth(), fechaModificacion.getDate());

//     if (nombreArchivo.indexOf('DR') !== -1 && fechaModificacionSinHoras.getTime() === hoySinHoras.getTime()) {
//       Logger.log(`Procesando archivo: ${nombreArchivo}`);

//       const regexNombreArchivo = /^(\w{5})_DR_(\d{8})/;
//       const match = nombreArchivo.match(regexNombreArchivo);
//       const proyecto = match ? match[1] : 'Proyecto desconocido';
//       const fecha_nombre_archivo = match ? match[2] : 'Fecha desconocida';

//       const archivoSS = SpreadsheetApp.openById(archivo.getId());
//       const hoja = archivoSS.getSheets()[0];
//       const valores = hoja.getDataRange().getValues();

//       const celdaFecha = hoja.getRange("M6").getValue();
//       let fechaFormateada;
//       if (celdaFecha && !isNaN(new Date(celdaFecha).getTime())) {
//         fechaFormateada = formatearFechaSimple(new Date(celdaFecha));
//         Logger.log(`La fecha en la celda M6 es ${new Date(celdaFecha)} Convertida a ${fechaFormateada}`);
//         fechasAEliminar.push(fechaFormateada);
//       } else {
//         Logger.log(`La fecha en la celda M6 no es válida o está ausente en el archivo ${nombreArchivo}`);
//         continue;
//       }

//       // Buscar "Actividad" y "Observaciones"
//       let filaInicio = -1, columnaInicio = -1, filaFin = -1;
//       for (let i = 0; i < valores.length; i++) {
//         for (let j = 0; j < valores[i].length; j++) {
//           const celda = typeof valores[i][j] === 'string' ? valores[i][j].trim().toLowerCase() : '';
//           if (celda === 'actividad' && filaInicio === -1) {
//             filaInicio = i;
//             columnaInicio = j - 1;
//           } else if (celda === 'observaciones' && filaFin === -1) {
//             filaFin = i - 3;
//           }
//         }
//       }

//       if (filaInicio === -1 || filaFin === -1 || filaInicio >= filaFin) {
//         Logger.log(`Error: No se encontraron las palabras clave "Actividad" u "Observaciones" en el archivo ${nombreArchivo}`);
//         continue;
//       }

//       const datos = hoja.getRange(filaInicio + 2, 1, filaFin - filaInicio - 1, valores[0].length).getValues();
//       const datosFiltrados = datos
//         .filter(fila => {
//           const acumulado = parseFloat(fila[8]);
//           return !isNaN(acumulado); // Incluye 0 válidos
//         })
//         .map(fila => {
//           fila[0] = fila[0] ? fila[0].toString().trim() : "";
//           fila.push(fechaFormateada, proyecto, fecha_nombre_archivo);
//           return fila;
//         });

//       if (datosFiltrados.length > 0) {
//         datosParaInsertar = datosParaInsertar.concat(datosFiltrados);
//         archivosProcesados.push(nombreArchivo);
//       }
//     }
//   }

//   Logger.log(`Fechas a eliminar: ${JSON.stringify(fechasAEliminar)}`);

//   // Borrar datos existentes para esas fechas
//   if (fechasAEliminar.length > 0) {
//     const datosHoja = hojaCarga.getDataRange().getValues();
//     const nuevasFilas = datosHoja.filter(fila => {
//       const fechaCelda = fila[15] ? formatearFechaSimple(new Date(fila[15])) : '';
//       return !fechasAEliminar.includes(fechaCelda);
//     });

//     hojaCarga.clear(); // Borra todo, incluyendo encabezados
//     if (nuevasFilas.length > 0) {
//       hojaCarga.getRange(1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
//     }

//     Logger.log(`Se eliminaron todas las filas con fechas: ${JSON.stringify(fechasAEliminar)}`);
//   }

//   // Insertar los nuevos datos
//   if (datosParaInsertar.length > 0) {
//     const rangoInsertar = hojaCarga.getRange(hojaCarga.getLastRow() + 1, 1, datosParaInsertar.length, datosParaInsertar[0].length);
//     rangoInsertar.setValues(datosParaInsertar);
//     Logger.log(`Datos a insertar: ${JSON.stringify(datosParaInsertar)}`);
//     Logger.log(`Datos insertados correctamente en Google Sheets.`);
//   }

//   Logger.log(`Archivos procesados: ${JSON.stringify(archivosProcesados)}`);
// }
// }
