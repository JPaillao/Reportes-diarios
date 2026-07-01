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
