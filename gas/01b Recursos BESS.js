/******************** FUNCIÓN PRINCIPAL RECURSOS BESS ********************/
/**
 * dr_recursos_bess con fecha de modificación parametrizable.
 * Lee recursos de hasta 3 columnas diferentes (ej. OPW, Tecton, Otros) y los aplana.
 * Genera la misma salida que el original añadiendo una 7ma columna "Empresa" al final.
 */
function dr_recursos_bess(carpetaID, id_archivo, nom_hoja, fecha) {
  const carpeta = DriveApp.getFolderById(carpetaID);
  const archivos = carpeta.getFiles();
  const hojaCarga = SpreadsheetApp.openById(id_archivo).getSheetByName(nom_hoja);
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
    const valores = hoja.getDataRange().getValues();

    // Encontrar fecha M6 o equivalente buscando la celda que dice "Fecha de daily report:"
    let celdaFecha = null;
    let paral = "", aff = "";

    // Búsquedas generales en toda la hoja para indicadores
    for (let i = 0; i < valores.length; i++) {
      for (let j = 0; j < valores[i].length; j++) {
        let v = String(valores[i][j]).trim().toLowerCase();
        if (v === "fecha de daily report:") {
          celdaFecha = valores[i][j + 1];
        }
        if (v.includes("paralización de faena")) paral = valores[i][j + 1];
        if (v.includes("financiero acumulado")) aff = valores[i][j + 1];
      }
    }

    // Fallback original M6 si no encuentra
    if (!celdaFecha) celdaFecha = hoja.getRange("M6").getValue();

    let fechaFormateada;
    if (celdaFecha && !isNaN(new Date(celdaFecha).getTime())) {
      fechaFormateada = formatearFechaSimple(new Date(celdaFecha));
    } else {
      Logger.log(`Fecha inválida/ausente en ${nombreArchivo} -> se omite.`);
      continue;
    }

    if (/^\d{8}$/.test(fecha_nombre_archivo)) {
      const fechaNombreDate = yyyymmddToDate(fecha_nombre_archivo);
      const fechaNombreForm = formatearFechaSimple(fechaNombreDate);
      if (fechaNombreForm !== fechaFormateada) {
        Logger.log(`Archivo con fecha NO consistente: ${nombreArchivo}`);
        continue;
      }
    }

    fechasAEliminar.push(fechaFormateada);

    // Detección dinámica de bloques de recursos (restringido a la fila de "Gastos generales" para evitar falsos positivos en metadata superior)
    let coordGG = null, coordTerr = null, coordMaq = null;
    for (let i = 0; i < Math.min(valores.length, 15); i++) {
      for (let j = 0; j < valores[i].length; j++) {
        let val = typeof valores[i][j] === 'string' ? valores[i][j].trim().toLowerCase() : '';
        
        if (!coordGG && val.includes("gastos generales")) {
          coordGG = { r: i, c: j };
        }
        
        // Buscar Terreno y Maquinaria SOLO en la misma fila que Gastos Generales
        if (coordGG && i === coordGG.r) {
          if (!coordTerr && j > coordGG.c && val.includes("terreno")) {
            coordTerr = { r: i, c: j };
          }
          if (!coordMaq && j > coordGG.c && val.includes("maquinaria")) {
            coordMaq = { r: i, c: j };
          }
        }
      }
    }

    // Constructor de aplanado
    const filtrados = [];

    function extractSection(coord, constantType) {
      if (!coord) return;

      let empresas = [];
      let startC = coord.c + 1;
      // Saltar columnas vacías que puedan existir entre el título de sección y las empresas
      while (startC < valores[coord.r].length && String(valores[coord.r][startC] || "").trim() === "") {
        startC++;
      }

      for (let c = startC; c < startC + 10; c++) {
        if (c >= valores[coord.r].length) break;
        let title = String(valores[coord.r][c] || "").trim();
        if (title.toLowerCase() === "total" || title === "") {
          break;
        }
        empresas.push({ name: title, offset: c - coord.c });
      }

      let i = coord.r + 1;
      while (i < valores.length) {
        let stop = false;
        for (let j = 0; j < valores[i].length; j++) {
           let val = String(valores[i][j]).trim().toLowerCase();
           if (val === "total" || val.includes("actividades de obra") || val.includes("control de avance")) {
               stop = true; break;
           }
        }
        if (stop) break;

        let recurso = typeof valores[i][coord.c] === 'string' ? valores[i][coord.c].trim() : String(valores[i][coord.c] || "").trim();
        
        // Si el recurso está vacío, puede que haya celdas combinadas y el texto esté en la columna adyacente
        if (recurso === "" && coord.c + 1 < valores[i].length) {
           let nextRecurso = String(valores[i][coord.c + 1] || "").trim();
           if (nextRecurso !== "" && isNaN(parseFloat(nextRecurso))) {
               recurso = nextRecurso;
           }
        }

        if (recurso !== "") {
          empresas.forEach(emp => {
            let q = parseFloat(valores[i][coord.c + emp.offset]);
            if (!isNaN(q) && q !== 0) {
              filtrados.push([recurso, q, constantType, "", "", "", emp.name]);
            }
          });
        }
        i++;
        if (i > coord.r + 60) break; // Límite de seguridad
      }
    }

    extractSection(coordGG, "Gastos generales");
    extractSection(coordTerr, "Terreno");
    extractSection(coordMaq, "Maquinaria");

    // Filas fijas
    filtrados.push(["Días con paralización de faena", paral || "", "", "", "", "", ""]);

    // Completar metadata [..., Fecha, Proyecto, YYYYMMDD, Empresa]
    filtrados.forEach(f => {
      f[3] = fechaFormateada;
      f[4] = proyecto;
      f[5] = fecha_nombre_archivo;
    });

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
    }
  }

  // --- Inserción ---
  if (data_to_insert.length > 0) {
    const startRow = hojaCarga.getLastRow() + 1;
    hojaCarga.getRange(startRow, 1, data_to_insert.length, data_to_insert[0].length)
      .setValues(data_to_insert);
    Logger.log(`✅ Insertadas ${data_to_insert.length} filas nuevas.`);
  }

  Logger.log(`Archivos procesados: ${archivos_leidos.join(', ')}`);
}
