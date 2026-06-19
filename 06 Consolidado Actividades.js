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
        const filteredAndSliced = sourceValues
          .filter(r => r[1] !== "" && r[1] !== null)
          .map(r => r.slice(0, 19)); 
        
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
    
    try {
      // LIMPIEZA VÍA SERVICIO AVANZADO:
      if (lastRowDestino > 1) {
        const clearRange = `${sheetName}!A2:S${lastRowDestino}`;
        Sheets.Spreadsheets.Values.clear({}, ssId, clearRange);
      }
      
      // ESCRITURA VÍA SERVICIO AVANZADO:
      const endRow = 1 + masterData.length;
      const writeRange = `${sheetName}!A2:S${endRow}`;
      const resource = { values: masterData };
      
      Sheets.Spreadsheets.Values.update(resource, ssId, writeRange, { valueInputOption: "USER_ENTERED" });
      Logger.log("Escritura exitosa vía Servicio Avanzado Sheets.");
      
    } catch (e) {
      Logger.log("Fallback tradicional. Motivo: " + e.message);
      // Si el servicio avanzado no está habilitado, ejecutamos la vía tradicional
      if (lastRowDestino > 1) {
        sheetDestino.getRange(2, 1, lastRowDestino - 1, 19).clearContent();
      }
      sheetDestino.getRange(2, 1, masterData.length, 19).setValues(masterData);
    }
    
    Logger.log("Consolidado actualizado hasta columna S. Filas: " + masterData.length);
  }
}