/**
 * Versión para script externo que respeta fórmulas desde la columna T en adelante.
 * 20260609
 */
function recopilarDRConsolidadoS() {
  const ssId = "1E0FylabCiHlOnFL2UrFPHTBJG0wTo9cwaVM5khh51IU"; //https://docs.google.com/spreadsheets/d/1E0FylabCiHlOnFL2UrFPHTBJG0wTo9cwaVM5khh51IU/edit
  const ss = SpreadsheetApp.openById(ssId);
  const sheetEstatus = ss.getSheetByName("Estatus_DR");
  const sheetDestino = ss.getSheetByName("Proyectos_activos_carga");

  if (!sheetEstatus || !sheetDestino) {
    Logger.log("Error: No se encontraron las hojas necesarias.");
    return;
  }

  const lastRowEstatus = sheetEstatus.getLastRow();
  if (lastRowEstatus < 2) return;
  const dataEstatus = sheetEstatus.getRange(2, 1, lastRowEstatus - 1, 5).getValues();

  let masterData = [];

  dataEstatus.forEach((row) => {
    const estado = row[1]; // Columna B
    const url = row[2];    // Columna C
    const rango = row[3];  // Columna D
    
    if (estado === "Activo" && url && rango) {
      try {
        const externalSS = SpreadsheetApp.openByUrl(url);
        const sourceValues = externalSS.getRange(rango).getValues();
        
        // FILTRAR Y RECORTAR: 
        // 1. Filtramos filas vacías (columna B / índice 1).
        // 2. Usamos .map(r => r.slice(0, 19)) para asegurar que solo traemos hasta la columna S.
        // 3. Formateamos los objetos Date a string para que la API REST no los suba como '2026-03-16T03:00:00.000Z'
        const filteredAndSliced = sourceValues
          .filter(r => r[1] !== "" && r[1] !== null)
          .map(r => r.slice(0, 19).map(cell => {
             if (cell instanceof Date) {
               return Utilities.formatDate(cell, "America/Santiago", "dd-MM-yyyy");
             }
             return cell;
          })); 
        
        if (filteredAndSliced.length > 0) {
          masterData = masterData.concat(filteredAndSliced);
        }
      } catch (e) {
        Logger.log("Error en " + url + ": " + e.message);
      }
    }
  });

  if (masterData.length > 0) {
    const lastRowDestino = sheetDestino.getLastRow();
    const sheetName = sheetDestino.getName();
    const token = ScriptApp.getOAuthToken();
    
    // LIMPIEZA VÍA API:
    // Al usar la REST API evitamos que el script de Apps Script se quede congelado
    // esperando el recalculo completo de la fórmula FILTER en 'proyectos_activos'.
    if (lastRowDestino > 1) {
      const clearRange = `${sheetName}!A2:S${lastRowDestino}`;
      const urlClear = `https://sheets.googleapis.com/v4/spreadsheets/${ssId}/values/${encodeURIComponent(clearRange)}:clear`;
      try {
        UrlFetchApp.fetch(urlClear, {
          method: "post",
          headers: { Authorization: "Bearer " + token },
          muteHttpExceptions: true
        });
      } catch (e) {
        Logger.log("Fallback limpieza tradicional. Error API: " + e.message);
        sheetDestino.getRange(2, 1, lastRowDestino - 1, 19).clearContent();
      }
    }
    
    // ESCRITURA VÍA API:
    // Insertamos los datos únicamente en el bloque A2:S mediante HTTP PUT
    const endRow = 1 + masterData.length;
    const writeRange = `${sheetName}!A2:S${endRow}`;
    const urlUpdate = `https://sheets.googleapis.com/v4/spreadsheets/${ssId}/values/${encodeURIComponent(writeRange)}?valueInputOption=USER_ENTERED`;
    
    try {
      const response = UrlFetchApp.fetch(urlUpdate, {
        method: "put",
        headers: { Authorization: "Bearer " + token },
        contentType: "application/json",
        payload: JSON.stringify({ values: masterData }),
        muteHttpExceptions: true
      });
      Logger.log("Respuesta API Escritura: " + response.getResponseCode());
    } catch (e) {
      Logger.log("Fallback escritura tradicional. Error API: " + e.message);
      sheetDestino.getRange(2, 1, masterData.length, 19).setValues(masterData);
    }
    
    Logger.log("Consolidado actualizado hasta columna S. Filas: " + masterData.length);
  }
}