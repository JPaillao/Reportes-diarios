function dr_recursos_esp(carpetaID, id_archivo, nom_hoja) {

  const archivos = obtenerArchivosRecursivos(carpetaID); // Ahora es un array de archivos

  // Hoja de carga donde se verificará si ya se cargó la fecha
  const hojaCarga = SpreadsheetApp.openById(id_archivo).getSheetByName(nom_hoja);

  if (!hojaCarga) {
    Logger.log(`No se pudo acceder a la hoja: ${nom_hoja}`);
    return;
  }
  Logger.log(`Hoja de carga abierta: ${nom_hoja}`);

  const hoy = new Date();
  const hoySinHoras = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());

  // Recorrer todos los archivos en la lista de archivos
  archivos.forEach(archivo => {
    let listadoValores = [];
    Logger.log(`Archivo detectado: ${archivo.getName()}`);
    const nombreArchivo = archivo.getName();

    // Comprobar si el archivo ha sido modificado hoy y su nombre contiene "Daily"
    if (nombreArchivo.indexOf('Daily') !== -1) {
      const archivoSS = SpreadsheetApp.openById(archivo.getId());
      const hoja = archivoSS.getSheets()[0];
      const valores = hoja.getDataRange().getValues();

      const patrones = {
        mo: /PERSONAL ACTIVO TECTON/i,
        ma: /MAQUINARIA/i,
        op: /PERSONAL EN TERRENO ORION POWER/i,
        to: /TOTAL PERSONAL/i
      };

      let filasPatrones = {
        mo: null,
        ma: null,
        op: null,
        to: null
      };

      // Buscar las filas que contienen los patrones
      for (let i = 0; i < valores.length; i++) {
        for (let j = 0; j < valores[i].length; j++) {
          if (typeof valores[i][j] === 'string') {
            for (let key in patrones) {
              if (patrones[key].test(valores[i][j]) && filasPatrones[key] === null) {
                filasPatrones[key] = i;
                Logger.log(`Primera coincidencia para ${key} encontrada en la fila ${i + 1}`);
              }
            }
          }
        }
      }

      // Verificar si hay coincidencias en filas MO y TO
      if (filasPatrones.mo !== null && filasPatrones.to !== null) {
        const filaInicio = filasPatrones.mo;
        const filaFin = filasPatrones.to;
        const datosEntreMOyTO = valores.slice(filaInicio, filaFin + 1);

        Logger.log('Datos completos entre MO y TO:');
        datosEntreMOyTO.forEach(fila => Logger.log(fila));

        // Extraer el proyecto y la fecha del nombre del archivo
        const regexNombreArchivo = /^Daily Report (\w+) (\d{4})-(\d{2})-(\d{2})/;
        const match = archivo.getName().match(regexNombreArchivo);
        const fechaArchivo = new Date(match[2], match[3] - 1, match[4]);
        if (isNaN(fechaArchivo.getTime())) {
          Logger.log(`La fecha extraída del archivo es inválida: ${match[2]}-${match[3]}-${match[4]}`);
          return; // Saltar al siguiente archivo
        }

        const fechaFormateada = fechaArchivo;
        const proyecto = match ? match[1] : 'Proyecto desconocido';

        Logger.log(`Fecha formateada: ${fechaFormateada}, Proyecto: ${proyecto}`);

        // Reunir los datos de las columnas necesarias
        datosEntreMOyTO.forEach(fila => {
          const valorColumna6 = fila[5]; // Columna 6 (índice 5)
          const valorColumna5 = fila[4]; // Columna 5 (índice 4)

          if (valorColumna6 !== '' && valorColumna6 !== null && valorColumna6 !== undefined) {
            listadoValores.push([valorColumna6, valorColumna5]);
          }
        });

        // Asegurarse de que los datos a insertar tengan las metas de fecha y proyecto
        if (listadoValores.length > 0) {
          listadoValores = listadoValores.map(fila => {
            fila.push(fechaFormateada);
            fila.push(proyecto);
            return fila;
          });

          const rangoFechas = hojaCarga.getRange(2, 3, hojaCarga.getLastRow() - 1, 1).getValues();
          Logger.log(rangoFechas);

          let filasAEliminar = [];
          for (let i = 0; i < rangoFechas.length; i++) {
            const fechaExistente = new Date(rangoFechas[i][0]);

            if (fechaExistente.getTime() === fechaFormateada.getTime()) {
              filasAEliminar.push(i + 2);
            }
          }

          if (filasAEliminar.length > 0) {
            filasAEliminar.reverse();
            filasAEliminar.forEach(fila => {
              hojaCarga.deleteRow(fila);
            });
            Logger.log(`Filas eliminadas para la fecha ${fechaFormateada}`);
          } else {
            Logger.log(`No se encontraron filas para eliminar con la fecha ${fechaFormateada}`);
          }

          const ultimaFila = hojaCarga.getLastRow();
          const rangoInsertar = hojaCarga.getRange(ultimaFila + 1, 1, listadoValores.length, 4);
          rangoInsertar.setValues(listadoValores);
          Logger.log(`Proceso completado correctamente. Valores insertados verticalmente.`);
        } else {
          Logger.log('No se encontraron coincidencias para ambos patrones (MO y TO).');
        }
      }
    }
  });
}
