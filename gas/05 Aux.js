/** Convierte 'YYYYMMDD' (string) a Date (local) */
function yyyymmddToDate(yyyymmdd) {
  if (!yyyymmdd || yyyymmdd.length !== 8) return null;
  const Y = parseInt(yyyymmdd.substring(0, 4), 10);
  const M = parseInt(yyyymmdd.substring(4, 6), 10) - 1;
  const D = parseInt(yyyymmdd.substring(6, 8), 10);
  return new Date(Y, M, D);
}

/**
 * Compacta índices de filas (1-based) en rangos contiguos: [[start, count], ...]
 * Ej: [5,6,12,20,21,22] -> [[5,2], [12,1], [20,3]]
 */
function _compactRowIndicesToRanges(rows1based) {
  if (!rows1based.length) return [];
  const rows = rows1based.slice().sort((a,b) => a-b);
  const ranges = [];
  let start = rows[0], prev = rows[0], count = 1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i] === prev + 1) { count++; prev = rows[i]; }
    else { ranges.push([start, count]); start = rows[i]; prev = rows[i]; count = 1; }
  }
  ranges.push([start, count]);
  return ranges;
}

/**
 * Devuelve índices de filas (1-based) cuya fecha en la columna `dateCol`
 * coincide con cualquiera en `fechasSet` (formato dd-MM-yyyy).
 * Empieza en `startRow` (usa 2 si tienes encabezado en la fila 1).
 */
function _findRowIndicesByDates(sheet, dateCol, fechasSet, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  const num = lastRow - startRow + 1;

  const vals = sheet.getRange(startRow, dateCol, num, 1).getValues().map(r => r[0]);

  const toKey = (v) => {
    if (!v) return '';
    if (v instanceof Date) return formatearFechaSimple(v); // dd-MM-yyyy
    if (typeof v === 'string') {
      if (/^\d{2}-\d{2}-\d{4}$/.test(v)) return v; // ya está normalizado
      const d = new Date(v);
      return isNaN(d.getTime()) ? '' : formatearFechaSimple(d);
    }
    const d = new Date(v);
    return isNaN(d.getTime()) ? '' : formatearFechaSimple(d);
  };

  const rows = [];
  for (let i = 0; i < vals.length; i++) {
    const key = toKey(vals[i]);
    if (key && fechasSet.has(key)) rows.push(startRow + i);
  }
  return rows;
}

/** Borra rangos de filas (1-based) de atrás hacia adelante para no desalinear índices. */
function _deleteRowRanges(sheet, ranges) {
  for (let i = ranges.length - 1; i >= 0; i--) {
    const [start, count] = ranges[i];
    sheet.deleteRows(start, count);
  }
}