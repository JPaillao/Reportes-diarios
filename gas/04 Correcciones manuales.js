/////Ejecutar reporte manualmente con fecha específica

 // 🗂️ IDs fijos de consolidado de recursos y actividades
const ID_RECURSOS = '1GMoZdbxfBRSEuf0NrHvw1wZ93eVYe65w_SASAVLmXNc';
const ID_ACTIVIDADES = '1pjcCzixMEtVPc5jHkk-106aOgDSkLw5drlQz41rIdes'; //hoja exclusiva para carga de datos, existe otra para compilar para evitar sobre procesameinto con cada ejecución

function corregir(){
  // dr_actividades("1QUHHNtraAdamVI-9wtQfjXeqaI0g5q_l", ID_ACTIVIDADES, 'SANRA', {desde:'2025-09-29', hasta:'2025-10-02'})
  // dr_actividades("1kKyOSDUUdnrq1iGgCA3_zqepUa2d2ucF", ID_ACTIVIDADES, 'QUILO', {desde:'2025-10-01', hasta:'2025-10-02'})
  // dr_recursos("1kKyOSDUUdnrq1iGgCA3_zqepUa2d2ucF",ID_ACTIVIDADES,'QUILO','2025-09-30')
    // dr_actividades("1z90w2KJdu6pn_WUl2jr0gQhhvTASigcp","1E0FylabCiHlOnFL2UrFPHTBJG0wTo9cwaVM5khh51IU",'LVIOL',{desde:'2025-08-13', hasta:'2025-08-14'})
    
    // dr_recursos("1kKyOSDUUdnrq1iGgCA3_zqepUa2d2ucF",ID_RECURSOS,'QUILO','2025-10-15')
    // dr_recursos("1z90w2KJdu6pn_WUl2jr0gQhhvTASigcp",ID_RECURSOS,'LVIOL','2025-10-15')
    // dr_recursos("1QUHHNtraAdamVI-9wtQfjXeqaI0g5q_l",ID_RECURSOS,'SANRA','2025-10-15')
    // dr_actividades("1qNn58hhJP3C63DKxguySdrxh-_dDD0dI", ID_ACTIVIDADES, 'SJCHI', {desde:'2026-03-16', hasta:'2026-03-20'})
    dr_actividades_bess("1SntflxGmVW1DwTEEM865M4AlLD4QrD3j", ID_ACTIVIDADES, 'CFDAL', {desde:'2026-05-18', hasta:'2026-05-26'})
    // dr_recursos_bess("1SntflxGmVW1DwTEEM865M4AlLD4QrD3j", ID_RECURSOS, 'CFDAL', {desde:'2026-04-05', hasta:'2026-05-26'})
    }