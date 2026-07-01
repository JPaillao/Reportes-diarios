function formatearFechaSimple(fecha) {
  if (!(fecha instanceof Date)) return fecha;
  const anio = fecha.getFullYear();
  const mes = String(fecha.getMonth() + 1).padStart(2, '0'); // Meses empiezan desde 0
  const dia = String(fecha.getDate()).padStart(2, '0');
  return `${dia}-${mes}-${anio}`;
}

// Función para obtener todos los archivos en una carpeta y sus subcarpetas
function obtenerArchivosRecursivos(carpetaID) {
  const carpeta = DriveApp.getFolderById(carpetaID);
  let archivos = [];

  // Obtener archivos de la carpeta actual
  const archivosDirectos = carpeta.getFiles();
  while (archivosDirectos.hasNext()) {
    archivos.push(archivosDirectos.next());
  }

  // Obtener archivos de cada subcarpeta
  const subcarpetas = carpeta.getFolders();
  while (subcarpetas.hasNext()) {
    const subcarpeta = subcarpetas.next();
    archivos = archivos.concat(obtenerArchivosRecursivos(subcarpeta.getId()));
  }

  return archivos;
}