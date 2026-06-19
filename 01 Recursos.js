/******************** FUNCIÓN PRINCIPAL ********************/
/**
 * dr_recursos con fecha de modificación parametrizable.
 *
 * - carpetaID  : ID de la carpeta con los DR del proyecto.
 * - id_archivo : ID del Spreadsheet destino (consolidado).
 * - nom_hoja   : Nombre de la hoja dentro del Spreadsheet destino.
 * - fecha      : (opcional) 'YYYY-MM-DD' | Date | {desde:'YYYY-MM-DD', hasta:'YYYY-MM-DD'}.
 *                Si no se pasa, usa "ayer".
 *
 * ⚙️ Versión estable (misma lógica de rangos y salida que la original)
 * 🔧 Ajustes mínimos:
 *    - sleep() solo si hubo borrado real.
 *    - push(...array) en vez de concat().
 */
function dr_recursos(carpetaID, id_archivo, nom_hoja, fecha) {
  const carpeta = DriveApp.getFolderById(carpetaID);
  const archivos = carpeta.getFiles();
  const hojaCarga = SpreadsheetApp.openById(id_archivo).getSheetByName(nom_hoja);
  if (!hojaCarga) {
    throw new Error(`No se encontró la hoja de destino nombrada "${nom_hoja}" en el archivo consolidado.`);
  }
  hojaCarga.getFilter()?.remove();

  const hoy = new Date();
  const ayer = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - 1);
  const sinHoras = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const parseYYYYMMDD = (s) => {
    const [Y, M, D] = s.split('-').map(Number);
    return new Date(Y, M - 1, D);
  };

  let desde = sinHoras(ayer), hasta = sinHoras(ayer);
  if (fecha) {
    if (typeof fecha === 'string') {
      const f = sinHoras(parseYYYYMMDD(fecha)); desde = f; hasta = f;
    } else if (fecha instanceof Date) {
      const f = sinHoras(fecha); desde = f; hasta = f;
    } else if (typeof fecha === 'object' && fecha.desde && fecha.hasta) {
      desde = sinHoras(parseYYYYMMDD(fecha.desde));
      hasta = sinHoras(parseYYYYMMDD(fecha.hasta));
    }
  }

  const fechasAEliminar = [];
  const data_to_insert = [];
  const archivos_leidos = [];

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName();
    if (!/DR/i.test(nombreArchivo)) continue;
    if (archivo.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    const fechaMod = archivo.getLastUpdated();
    const fmSin = new Date(fechaMod.getFullYear(), fechaMod.getMonth(), fechaMod.getDate());
    if (fmSin < desde || fmSin > hasta) continue;

    Logger.log(`Procesando archivo: ${nombreArchivo}`);

    const regexNombreArchivo = /^(\w{5})_DR_(\d{8})/;
    const match = nombreArchivo.match(regexNombreArchivo);
    const proyecto = match ? match[1] : 'Proyecto desconocido';
    const fecha_nombre_archivo = match ? match[2] : 'Fecha desconocida';

    const archivoSS = SpreadsheetApp.openById(archivo.getId());
    const hoja = archivoSS.getSheets()[0];

    const celdaFecha = hoja.getRange("M6").getValue();
    let fechaFormateada;
    if (celdaFecha && !isNaN(new Date(celdaFecha).getTime())) {
      fechaFormateada = formatearFechaSimple(new Date(celdaFecha));
    } else {
      Logger.log(`M6 inválida/ausente en ${nombreArchivo} -> se omite.`);
      continue;
    }

    if (/^\d{8}$/.test(fecha_nombre_archivo)) {
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
        continue;
      }
    }



    fechasAEliminar.push(fechaFormateada);

    // --- Detectar inicio de recursos ---
    const valores = hoja.getDataRange().getValues();
    let filaInicio = -1, columnaInicio = -1;
    for (let i = 0; i < valores.length; i++) {
      for (let j = 0; j < valores[i].length; j++) {
        if (typeof valores[i][j] === 'string' && /(?:tecton|bos)\s*-\s*gastos generales/i.test(valores[i][j])) {
          filaInicio = i; columnaInicio = j; break;
        }
      }
      if (filaInicio !== -1) break;
    }
    if (filaInicio === -1) continue;

    const numfil = 22;
    const dfRango1 = hoja.getRange(filaInicio + 2, columnaInicio + 1, numfil, 2).getValues();
    const dfRango2 = hoja.getRange(filaInicio + 2, columnaInicio + 4, numfil, 3).getValues();
    const dfRango3 = hoja.getRange(filaInicio + 2, columnaInicio + 8, numfil, 2).getValues();

    const constante1 = "Gastos generales";
    const constante2 = "Terreno";
    const constante3 = "Maquinaria";
    const paral = hoja.getRange("M12").getValue();
    const aff = hoja.getRange("M17").getValue();

    const dfValores1 = dfRango1.map(f => [...f, constante1]);
    const dfValores2 = dfRango2.map(f => [f[0], f[2], constante2]);
    const dfValores3 = dfRango3.map(f => [...f, constante3]);

    const ValoresConcatenados = [];
    ValoresConcatenados.push(...dfValores1, ...dfValores2, ...dfValores3);

    const filtrados = ValoresConcatenados
      .filter(f => f[0] && f[0].toString().trim() !== "")
      .filter(f => f[1] && parseFloat(f[1]) !== 0);

    filtrados.push(["Días con paralización de faena", paral, ""]);
    filtrados.push(["Avance físico financiero acumulado", aff, ""]);
    filtrados.forEach(f => f.push(fechaFormateada, proyecto, fecha_nombre_archivo));

    data_to_insert.push(...filtrados);
    archivos_leidos.push(nombreArchivo);
  }

  // --- Limpieza ---
  if (fechasAEliminar.length > 0) {
    const fechasUnicas = Array.from(new Set(fechasAEliminar));
    const fechasSet = new Set(fechasUnicas);
    const dateCol = 4, startRow = 2;
    const rowsToDelete = _findRowIndicesByDates(hojaCarga, dateCol, fechasSet, startRow);
    if (rowsToDelete.length > 0) {
      const ranges = _compactRowIndicesToRanges(rowsToDelete);
      _deleteRowRanges(hojaCarga, ranges);
      Logger.log(`🧹 Borradas ${rowsToDelete.length} filas (${ranges.length} bloque(s)).`);
      Utilities.sleep(2000);
    } else {
      Logger.log(`Sin filas previas para fechas ${JSON.stringify(fechasUnicas)}.`);
    }
  }

  // --- Inserción ---
  if (data_to_insert.length > 0) {
    const startRow = hojaCarga.getLastRow() + 1;
    hojaCarga.getRange(startRow, 1, data_to_insert.length, data_to_insert[0].length)
      .setValues(data_to_insert);
    Logger.log(`✅ Insertadas ${data_to_insert.length} filas nuevas.`);
  } else {
    Logger.log(`Sin datos nuevos para insertar.`);
  }

  Logger.log(`Archivos procesados: ${archivos_leidos.join(', ')}`);
}



// /******************** FUNCIÓN PRINCIPAL ********************/

// /**
//  * dr_recursos con fecha de modificación parametrizable.
//  *
//  * - carpetaID  : ID de la carpeta con los DR del proyecto.
//  * - id_archivo : ID del Spreadsheet destino (consolidado).
//  * - nom_hoja   : Nombre de la hoja dentro del Spreadsheet destino.
//  * - fecha      : (opcional) 'YYYY-MM-DD' | Date | {desde:'YYYY-MM-DD', hasta:'YYYY-MM-DD'}.
//  *                Si no se pasa, usa "ayer".
//  *
//  * Cambios clave:
//  *  - Limpieza: se BORRAN SOLO las filas de las fechas afectadas (col 4) en BLOQUES contiguos (no 1x1).
//  *  - Inserción: se inserta TODO lo nuevo en UNA sola llamada setValues (un “golpe”).
//  *  - Pausa: sleep(5000) tras el borrado, antes de insertar (da tiempo a que cese recálculo).
//  *  - Se mantiene la validación de consistencia fecha nombre (YYYYMMDD) vs M6 (dd-MM-yyyy).
//  */
// function dr_recursos(carpetaID, id_archivo, nom_hoja, fecha) {
//   // --- Setup base ---
//   const carpeta = DriveApp.getFolderById(carpetaID);
//   const archivos = carpeta.getFiles();
//   const hojaCarga = SpreadsheetApp.openById(id_archivo).getSheetByName(nom_hoja);

//   // Quitar filtro si existe (no toca formatos/validaciones).
//   hojaCarga.getFilter()?.remove();

//   // --- Resolver "ayer" por defecto + parametrización de fecha/rango ---
//   const hoy = new Date();
//   const ayer = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - 1);

//   const sinHoras = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate()); // trunc local
//   const parseYYYYMMDD = (s) => {
//     const [Y, M, D] = s.split('-').map(Number);
//     return new Date(Y, M - 1, D);
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

//   // --- Acumuladores ---
//   const fechasAEliminar = []; // dd-MM-yyyy (desde M6 de los DR válidos)
//   const data_to_insert = [];  // [Recurso, Cantidad, Tipo, Fecha(dd-MM-yyyy), Proyecto, YYYYMMDD_nombre]
//   const archivos_leidos = []; // logging

//   // --- Recorrido de archivos DR ---
//   while (archivos.hasNext()) {
//     const archivo = archivos.next();
//     const nombreArchivo = archivo.getName();
//     if (nombreArchivo.indexOf('DR') === -1) continue; // mismo filtro por nombre

//     // Filtrar por lastUpdated truncado a día en [desde, hasta]
//     const fechaMod = archivo.getLastUpdated();
//     const fmSin = new Date(fechaMod.getFullYear(), fechaMod.getMonth(), fechaMod.getDate());
//     if (fmSin < desde || fmSin > hasta) continue;

//     Logger.log(`Procesando archivo: ${nombreArchivo}`);

//     // Proyecto y fecha del nombre
//     const regexNombreArchivo = /^(\w{5})_DR_(\d{8})/;
//     const match = nombreArchivo.match(regexNombreArchivo);
//     const proyecto = match ? match[1] : 'Proyecto desconocido';
//     const fecha_nombre_archivo = match ? match[2] : 'Fecha desconocida';

//     // Abrir DR
//     const archivoSS = SpreadsheetApp.openById(archivo.getId());
//     const hoja = archivoSS.getSheets()[0];
//     const valores = hoja.getDataRange().getValues();

//     // Fecha operativa M6
//     const celdaFecha = hoja.getRange("M6").getValue();
//     let fechaFormateada;
//     if (celdaFecha && !isNaN(new Date(celdaFecha).getTime())) {
//       fechaFormateada = formatearFechaSimple(new Date(celdaFecha)); // dd-MM-yyyy
//     } else {
//       Logger.log(`M6 inválida/ausente en ${nombreArchivo} -> se omite.`);
//       continue;
//     }
//     Logger.log(`Fecha M6 (formateada): ${fechaFormateada}`);

//     // Consistencia nombre vs M6
//     if (fecha_nombre_archivo && /^\d{8}$/.test(fecha_nombre_archivo)) {
//       const fechaNombreDate = yyyymmddToDate(fecha_nombre_archivo);
//       const fechaNombreForm = formatearFechaSimple(fechaNombreDate);
//       if (fechaNombreForm !== fechaFormateada) {
//         const msg = [
//           `Archivo con fecha NO consistente: ${nombreArchivo}`,
//           `Proyecto: ${proyecto}`,
//           `Fecha en nombre: ${fechaNombreForm} (de ${fecha_nombre_archivo})`,
//           `Fecha en M6:     ${fechaFormateada}`,
//           `Acción: NO se cargó este DR.`
//         ].join('\n');
//         Logger.log(msg);
//         try {
//           MailApp.sendEmail({
//             to: "jpaillao@orion-power.com",
//             subject: `🚨 DR con fecha inconsistente: ${proyecto} / ${nombreArchivo}`,
//             body: `Hola,\n\n${msg}\n\nSaludos,\nDR Bot`
//           });
//         } catch (e) {
//           Logger.log(`Error enviando correo por inconsistencia: ${e}`);
//         }
//         continue; // no cargar este archivo
//       }
//     }

//     // Guardar fecha a eliminar en destino (col 4)
//     fechasAEliminar.push(fechaFormateada);

//     // Detectar inicio de recursos (tu regex original)
//     let filaInicio = -1, columnaInicio = -1, matchCount = 0;
//     for (let i = 0; i < valores.length; i++) {
//       for (let j = 0; j < valores[i].length; j++) {
//         if (typeof valores[i][j] === 'string' && /(?:tecton|bos)\s*-\s*gastos generales/i.test(valores[i][j])) {
//           matchCount++;
//           if (matchCount === 1) { filaInicio = i; columnaInicio = j; break; }
//         }
//       }
//       if (matchCount === 1) break;
//     }
//     if (filaInicio === -1 || columnaInicio === -1) {
//       Logger.log(`No se encontró el inicio de los recursos en ${nombreArchivo} -> se omite.`);
//       continue;
//     }

//     // Captura de rangos (layout original)
//     const numfil = 22;
//     const dfRango1 = hoja.getRange(filaInicio + 2, columnaInicio + 1, numfil, 2).getValues();
//     const dfRango2 = hoja.getRange(filaInicio + 2, columnaInicio + 4, numfil, 3).getValues();
//     const dfRango3 = hoja.getRange(filaInicio + 2, columnaInicio + 8, numfil, 2).getValues();

//     const constante1 = "Gastos generales";
//     const constante2 = "Terreno";
//     const constante3 = "Maquinaria";
//     const paral = hoja.getRange("M12").getValue(); // días de paralización

//     // Armar y filtrar valores
//     const dfValores1 = dfRango1.map(fila => [...fila, constante1]);
//     const dfValores2 = dfRango2.map(fila => [fila[0], fila[2], constante2]); // elimina intermedia
//     const dfValores3 = dfRango3.map(fila => [...fila, constante3]);

//     let ValoresConcatenados = [].concat(dfValores1, dfValores2, dfValores3);
//     ValoresConcatenados = ValoresConcatenados.filter(fila => fila[0] && fila[0].toString().trim() !== "");
//     ValoresConcatenados = ValoresConcatenados.filter(fila => fila[1] && parseFloat(fila[1]) !== 0);

//     // Fila de paralización
//     ValoresConcatenados.push(["Días con paralización de faena", paral, ""]);

//     // Enriquecer con metadatos y acumular
//     if (ValoresConcatenados.length > 0) {
//       ValoresConcatenados.forEach(fila => {
//         // Estructura destino: [Recurso, Cantidad, Tipo, Fecha(dd-MM-yyyy), Proyecto, FechaNombreArchivo(YYYYMMDD)]
//         fila.push(fechaFormateada);      // Col 4: fecha operativa (M6) formateada
//         fila.push(proyecto);             // Col 5
//         fila.push(fecha_nombre_archivo); // Col 6: YYYYMMDD del nombre
//       });
//       data_to_insert.push(...ValoresConcatenados);
//       archivos_leidos.push(nombreArchivo);
//     }
//   } // fin while archivos

//   // --- LIMPIEZA: borrar SOLO las filas de las fechas afectadas (columna 4) en BLOQUES ---
//   if (fechasAEliminar.length > 0) {
//     const fechasUnicas = Array.from(new Set(fechasAEliminar));
//     const fechasSet = new Set(fechasUnicas);

//     const dateCol = 4;   // recursos: la fecha destino está en la columna 4
//     const startRow = 2;  // mantener encabezado de la fila 1 (ajusta si no tienes header)

//     const rowsToDelete = _findRowIndicesByDates(hojaCarga, dateCol, fechasSet, startRow);

//     if (rowsToDelete.length > 0) {
//       const ranges = _compactRowIndicesToRanges(rowsToDelete);
//       _deleteRowRanges(hojaCarga, ranges);
//       Logger.log(`Limpieza: borradas ${rowsToDelete.length} filas en ${ranges.length} bloque(s) para fechas ${JSON.stringify(fechasUnicas)}.`);
//     } else {
//       Logger.log(`Limpieza: no había filas previas para fechas ${JSON.stringify(fechasUnicas)}.`);
//     }

//     // Pausa de cortesía para que cesen recálculos/volátiles del consolidado antes de insertar
//     Utilities.sleep(5000);
//   }

//   // --- INSERCIÓN: un solo "golpe" (setValues) con TODO lo nuevo ---
//   if (data_to_insert.length > 0) {
//     const startRow = hojaCarga.getLastRow() + 1;
//     const numRows = data_to_insert.length;
//     const numCols = data_to_insert[0].length;

//     // Si fueran muchísimas filas, podrías trocear (chunks) aquí.
//     hojaCarga.getRange(startRow, 1, numRows, numCols).setValues(data_to_insert);
//     Logger.log(`Insertadas ${numRows} filas nuevas (una sola operación).`);
//   } else {
//     Logger.log(`No hay nuevas filas para insertar.`);
//   }

//   Logger.log(`Archivos procesados: ${JSON.stringify(archivos_leidos)}`);
// }