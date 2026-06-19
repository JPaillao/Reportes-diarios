function ejecutarRecursosPorProyecto() {
  const ID_RECURSOS = '1GMoZdbxfBRSEuf0NrHvw1wZ93eVYe65w_SASAVLmXNc';
  const proyectos = [
    // { nombre: 'COSEC', carpetaID: '1BWmzBDHtlmEc-XM2uaxV6GJIrCIpi1vG' },
    // { nombre: 'GUIND', carpetaID: '1dg08Y8sTTwhN48PyIdXL9B5jBs9NfH1u' },
    // { nombre: 'SANRA', carpetaID: '1QUHHNtraAdamVI-9wtQfjXeqaI0g5q_l' },
    { nombre: 'QUILO', carpetaID: '1kKyOSDUUdnrq1iGgCA3_zqepUa2d2ucF' },
    { nombre: 'LVIOL', carpetaID: '1z90w2KJdu6pn_WUl2jr0gQhhvTASigcp' },
    { nombre: 'SJCHI', carpetaID: '1qNn58hhJP3C63DKxguySdrxh-_dDD0dI' },
    { nombre: 'CFDAL', carpetaID: '1SntflxGmVW1DwTEEM865M4AlLD4QrD3j' }
  ];

  proyectos.forEach((proyecto, i) => {
    const lock = LockService.getScriptLock();
    try {
      Logger.log(`🔐 Adquiriendo lock para ${proyecto.nombre}...`);
      lock.waitLock(30000);

      Logger.log(`📁 Iniciando RECURSOS para ${proyecto.nombre}`);
      if (proyecto.nombre === 'CFDAL') {
        dr_recursos_bess(proyecto.carpetaID, ID_RECURSOS, proyecto.nombre);
      } else {
        dr_recursos(proyecto.carpetaID, ID_RECURSOS, proyecto.nombre);
      }
      Logger.log(`✅ Recursos OK para ${proyecto.nombre}`);

      // Pausa entre proyectos (evita saturar la hoja destino)
      if (i < proyectos.length - 1) {
        Logger.log("⏳ Pausa 4 segundos antes del siguiente proyecto...");
        Utilities.sleep(4000);
      }

    } catch (error) {
      const msg = `❌ Error en RECURSOS de ${proyecto.nombre}: ${error.message}`;
      Logger.log(msg);
      try {
        MailApp.sendEmail({
          to: "jpaillao@orion-power.com",
          subject: `🚨 Error en procesamiento de RECURSOS (${proyecto.nombre})`,
          body: `Se produjo un error al procesar los RECURSOS del proyecto ${proyecto.nombre}:\n\n${error.message}`
        });
      } catch (mailError) {
        Logger.log(`⚠️ Error enviando correo: ${mailError}`);
      }
    } finally {
      lock.releaseLock();
    }
  });
}


function ejecutarActividadesPorProyecto() {
  const ID_ACTIVIDADES = '1pjcCzixMEtVPc5jHkk-106aOgDSkLw5drlQz41rIdes';
  const proyectos = [
    // { nombre: 'COSEC', carpetaID: '1BWmzBDHtlmEc-XM2uaxV6GJIrCIpi1vG' },
    // { nombre: 'GUIND', carpetaID: '1dg08Y8sTTwhN48PyIdXL9B5jBs9NfH1u' },
    // { nombre: 'SANRA', carpetaID: '1QUHHNtraAdamVI-9wtQfjXeqaI0g5q_l' },
    { nombre: 'QUILO', carpetaID: '1kKyOSDUUdnrq1iGgCA3_zqepUa2d2ucF' },
    { nombre: 'LVIOL', carpetaID: '1z90w2KJdu6pn_WUl2jr0gQhhvTASigcp' },
    { nombre: 'SJCHI', carpetaID: '1qNn58hhJP3C63DKxguySdrxh-_dDD0dI' },
    { nombre: 'CFDAL', carpetaID: '1SntflxGmVW1DwTEEM865M4AlLD4QrD3j' }
  ];

  proyectos.forEach((proyecto, i) => {
    const lock = LockService.getScriptLock();
    try {
      Logger.log(`🔐 Adquiriendo lock para ${proyecto.nombre}...`);
      lock.waitLock(30000);

      Logger.log(`📝 Iniciando ACTIVIDADES para ${proyecto.nombre}`);
      if (proyecto.nombre === 'CFDAL') {
        dr_actividades_bess(proyecto.carpetaID, ID_ACTIVIDADES, proyecto.nombre);
      } else {
        dr_actividades(proyecto.carpetaID, ID_ACTIVIDADES, proyecto.nombre);
      }
      Logger.log(`✅ Actividades OK para ${proyecto.nombre}`);

      // Pausa entre proyectos (evita saturar la hoja destino)
      if (i < proyectos.length - 1) {
        Logger.log("⏳ Pausa 4 segundos antes del siguiente proyecto...");
        Utilities.sleep(4000);
      }

    } catch (error) {
      const msg = `❌ Error en ACTIVIDADES de ${proyecto.nombre}: ${error.message}`;
      Logger.log(msg);
      try {
        MailApp.sendEmail({
          to: "jpaillao@orion-power.com",
          subject: `🚨 Error en procesamiento de ACTIVIDADES (${proyecto.nombre})`,
          body: `Se produjo un error al procesar las ACTIVIDADES del proyecto ${proyecto.nombre}:\n\n${error.message}`
        });
      } catch (mailError) {
        Logger.log(`⚠️ Error enviando correo: ${mailError}`);
      }
    } finally {
      lock.releaseLock();
    }
  });
}
// function ejecutarDrPorProyecto() {
//   // 🗂️ IDs fijos de consolidado de recursos y actividades
//   const ID_RECURSOS = '1GMoZdbxfBRSEuf0NrHvw1wZ93eVYe65w_SASAVLmXNc';
//   const ID_ACTIVIDADES = '1pjcCzixMEtVPc5jHkk-106aOgDSkLw5drlQz41rIdes'; //hoja exclusiva para carga de datos, existe otra para compilar para evitar sobre procesameinto con cada ejecución

//   const proyectos = [
//     // { nombre: 'COSEC', carpetaID: '1BWmzBDHtlmEc-XM2uaxV6GJIrCIpi1vG' },
//     // { nombre: 'GUIND', carpetaID: '1dg08Y8sTTwhN48PyIdXL9B5jBs9NfH1u' },
//     { nombre: 'SANRA', carpetaID: '1QUHHNtraAdamVI-9wtQfjXeqaI0g5q_l' },
//     { nombre: 'QUILO', carpetaID: '1kKyOSDUUdnrq1iGgCA3_zqepUa2d2ucF' },
//     { nombre: 'LVIOL', carpetaID: '1z90w2KJdu6pn_WUl2jr0gQhhvTASigcp' }
//   ];

//   proyectos.forEach((proyecto, i) => {
//     const lock = LockService.getScriptLock();
//     try {
//       Logger.log(`🔐 Adquiriendo lock para ${proyecto.nombre}...`);
//       lock.waitLock(30000);

//       Logger.log(`📁 Iniciando RECURSOS para ${proyecto.nombre}`);
//       dr_recursos(proyecto.carpetaID, ID_RECURSOS, proyecto.nombre);

//       Utilities.sleep(5000); // pausa de 5 seg para no colapsar hoja

//       Logger.log(`📝 Iniciando ACTIVIDADES para ${proyecto.nombre}`);
//       dr_actividades(proyecto.carpetaID, ID_ACTIVIDADES, proyecto.nombre);

//       Logger.log(`✅ Proyecto ${proyecto.nombre} procesado correctamente.`);

//       if (i < proyectos.length - 1) {
//         Logger.log("⏳ Pausa 2 segundos antes del siguiente proyecto...");
//         Utilities.sleep(2000);
//       }

//     } catch (error) {
//       Logger.log(`❌ Error en ${proyecto.nombre}: ${error.message}`);
//       MailApp.sendEmail({
//         to: "jpaillao@orion-power.com",
//         subject: `🚨 Error en procesamiento de ${proyecto.nombre}`,
//         body: `Se produjo un error al procesar ${proyecto.nombre}:\n\n${error.message}`
//       });
//     } finally {
//       lock.releaseLock();
//     }
//   });
// }





// function ejecutarDrRecursosPorProyecto() {
//   const proyectos = [
//     // { nombre: 'LMAIT', carpetaID: '1t86Ko9_J4e-YkVwt56pXBBHPHYEwfCzL', idArchivo: '1r0cw0C_bpT4yqeEONV1d-iUDVNvZaZYU1O0UmtvvndA', nomHoja: 'LMAIT' },
//     { nombre: 'COSEC', carpetaID: '1BWmzBDHtlmEc-XM2uaxV6GJIrCIpi1vG', idArchivo: '1r0cw0C_bpT4yqeEONV1d-iUDVNvZaZYU1O0UmtvvndA', nomHoja: 'COSEC' },
//     { nombre: 'GUIND', carpetaID: '1dg08Y8sTTwhN48PyIdXL9B5jBs9NfH1u', idArchivo: '1r0cw0C_bpT4yqeEONV1d-iUDVNvZaZYU1O0UmtvvndA', nomHoja: 'GUIND' },
//     { nombre: 'SANRA', carpetaID: '1QUHHNtraAdamVI-9wtQfjXeqaI0g5q_l', idArchivo: '1r0cw0C_bpT4yqeEONV1d-iUDVNvZaZYU1O0UmtvvndA', nomHoja: 'SANRA' },
//     { nombre: 'QUILO', carpetaID: '1kKyOSDUUdnrq1iGgCA3_zqepUa2d2ucF', idArchivo: '1r0cw0C_bpT4yqeEONV1d-iUDVNvZaZYU1O0UmtvvndA', nomHoja: 'QUILO' },
//     { nombre: 'LVIOL', carpetaID: '1z90w2KJdu6pn_WUl2jr0gQhhvTASigcp', idArchivo: '1r0cw0C_bpT4yqeEONV1d-iUDVNvZaZYU1O0UmtvvndA', nomHoja: 'LVIOL' }
//   ];

//   proyectos.forEach((proyecto, i) => {
//     try {
//       Logger.log(`Iniciando proceso para el proyecto: ${proyecto.nombre}`);
//       dr_recursos(proyecto.carpetaID, proyecto.idArchivo, proyecto.nomHoja);
//       Logger.log(`Proceso completado correctamente para el proyecto: ${proyecto.nombre}`);

//       // ⏱️ Pausa de 2 segundos entre proyectos (excepto el último)
//       if (i < proyectos.length - 1) {
//         Logger.log("Esperando 2 segundos antes de continuar con el siguiente proyecto...");
//         Utilities.sleep(2000);
//       }

//     } catch (error) {
//       Logger.log(`Error en el proyecto ${proyecto.nombre}: ${error.message}`);
//       MailApp.sendEmail({
//         to: "jpaillao@orion-power.com",
//         subject: `🚨 Error en proyecto ${proyecto.nombre} (Recursos)`,
//         body: `Ocurrió un error al procesar los recursos del proyecto ${proyecto.nombre}:\n\n${error.message}`
//       });
//     }
//   });
// }

// function ejecutarDrActividadesPorProyecto() {
//   const proyectos = [
//  // { nombre: 'LMAIT', carpetaID: '1t86Ko9_J4e-YkVwt56pXBBHPHYEwfCzL', idArchivo: '1E0FylabCiHlOnFL2UrFPHTBJG0wTo9cwaVM5khh51IU', nomHoja: 'LMAIT' },
//     { nombre: 'COSEC', carpetaID: '1BWmzBDHtlmEc-XM2uaxV6GJIrCIpi1vG', idArchivo: '1E0FylabCiHlOnFL2UrFPHTBJG0wTo9cwaVM5khh51IU', nomHoja: 'COSEC' },
//     { nombre: 'GUIND', carpetaID: '1dg08Y8sTTwhN48PyIdXL9B5jBs9NfH1u', idArchivo: '1E0FylabCiHlOnFL2UrFPHTBJG0wTo9cwaVM5khh51IU', nomHoja: 'GUIND' },
//     { nombre: 'SANRA', carpetaID: '1QUHHNtraAdamVI-9wtQfjXeqaI0g5q_l', idArchivo: '1E0FylabCiHlOnFL2UrFPHTBJG0wTo9cwaVM5khh51IU', nomHoja: 'SANRA' },
//     { nombre: 'QUILO', carpetaID: '1kKyOSDUUdnrq1iGgCA3_zqepUa2d2ucF', idArchivo: '1E0FylabCiHlOnFL2UrFPHTBJG0wTo9cwaVM5khh51IU', nomHoja: 'QUILO' },
//     { nombre: 'LVIOL', carpetaID: '1z90w2KJdu6pn_WUl2jr0gQhhvTASigcp', idArchivo: '1E0FylabCiHlOnFL2UrFPHTBJG0wTo9cwaVM5khh51IU', nomHoja: 'LVIOL' }
//   ];

//   proyectos.forEach((proyecto, i) => {
//     try {
//       Logger.log(`Iniciando proceso para el proyecto: ${proyecto.nombre}`);
//       dr_actividades(proyecto.carpetaID, proyecto.idArchivo, proyecto.nomHoja);
//       Logger.log(`Proceso completado correctamente para el proyecto: ${proyecto.nombre}`);

//       if (i < proyectos.length - 1) {
//         Logger.log("Pausa de 2 segundos entre proyectos...");
//         Utilities.sleep(2000);
//       }
//     } catch (error) {
//       Logger.log(`Error en el proyecto ${proyecto.nombre}: ${error.message}`);
//       MailApp.sendEmail({
//         to: "jpaillao@orion-power.com",
//         subject: `🚨 Error en proyecto ${proyecto.nombre}`,
//         body: `Ocurrió un error al procesar el proyecto ${proyecto.nombre}:\n\n${error.message}`
//       });
//     }
//   });
// }





// function ejecutarDrRecursos_RUCAP() {
//   const proyectos = [
//     { nombre: 'RUCAP', carpetaID: '14tYr-ACIG0X_BsNDQT-Lbz0NO-uG6aIE', idArchivo: '1r0cw0C_bpT4yqeEONV1d-iUDVNvZaZYU1O0UmtvvndA', nomHoja: 'RUCAP' }
//   ]; 

//   // Ejecutar para cada proyecto
//   proyectos.forEach(proyecto => {
//     try {
//       Logger.log(`Iniciando proceso para el proyecto: ${proyecto.nombre}`);
//       // Llamada a la función que realiza el procesamiento
//       dr_recursos_esp(proyecto.carpetaID, proyecto.idArchivo, proyecto.nomHoja);
    
//       Logger.log(`Proceso completado correctamente para el proyecto: ${proyecto.nombre}`);
//     } catch (error) {
//       // Si ocurre un error, lo logueamos junto con el nombre del proyecto
//       Logger.log(`Error en el proyecto ${proyecto.nombre}: ${error.message}`);
//     }
//   });
// }

